# -*- coding: utf-8 -*-
import io
import re
import datetime as dt
from copy import copy
from pathlib import Path
from typing import Dict, Tuple, List, Optional

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

# ======================
# 固定配置（按你模板/谈判表）
# ======================
SHEET_OUT = "引入产品详细信息"
SHEET_NEG_DEFAULT = "产品谈判记录表"

# 只允许写入这些列（其它列保持模板原样/公式/格式）
WRITE_FIELDS = {
    "品牌",
    "型号",
    "品类",
    "CPU型号",
    "网络制式",
    "摄像头",
    "屏幕",
    "电池",
    "预计采购票面价（元）",
    "预计零售价（元）",
    "合同预计数量（台）",
}

# 供应商名称固定写入 S 列（你要求“必须在 S 列”）
SUPPLIER_COL_S = 19  # S = 19（1-indexed）

# ======================
# Streamlit UI
# ======================
st.set_page_config(page_title="OPPO 引入回填", layout="wide")
st.title("OPPO 引入回填（上传2个文件 → 一键生成）")
st.caption("✅ 自动识别谈判表/入库表（顺序随意）｜✅ 规格+价格+数量结合识别｜✅ 只写指定字段，其余保持模板原样")


# ======================
# 工具函数
# ======================
def norm_text(x) -> str:
    if x is None:
        return ""
    return str(x).strip().replace("　", " ").replace("\u00a0", " ").strip()


def safe_set(cell, value):
    if isinstance(cell, MergedCell):
        return
    cell.value = value


def merged_top_left_value(ws, r: int, c: int):
    """合并单元格：取合并区域左上角"""
    try:
        for mr in ws.merged_cells.ranges:
            if mr.min_row <= r <= mr.max_row and mr.min_col <= c <= mr.max_col:
                return ws.cell(mr.min_row, mr.min_col).value
    except Exception:
        pass
    return ws.cell(r, c).value


def read_cell(ws, r: int, c: int):
    return merged_top_left_value(ws, r, c)


def normalize_model_name(x) -> str:
    """用于 token 抽取前的清洗（不改变谈判表原始型号文本）"""
    s = norm_text(x).upper()
    s = s.replace("（", "(").replace("）", ")")
    for bad in ["全网通", "移动", "联通", "电信", "公开版", "定制", "TD-LTE", "LTE", "NR"]:
        s = s.replace(bad, "")
    return s.strip()


def extract_model_token(s: str) -> str:
    """
    从字符串中提取型号主体 token，如：
    - A6x（PLT140 6G+128G） => PLT140
    - Watch X3（OWW261）     => OWW261
    - 15T（PLZ110 16G+512G） => PLZ110
    """
    s = normalize_model_name(s)
    m = re.search(r"[A-Z]{2,}\d+[A-Z0-9]*", s)
    return m.group(0) if m else s


# ======================
# 模板文件（内置仓库）
# ======================
def find_template_path() -> Path:
    candidates = [
        Path("template.xlsx"),
        Path("template(1).xlsx"),
        Path("assets/template.xlsx"),
    ]
    for p in candidates:
        if p.exists():
            return p
    raise RuntimeError("仓库内未找到模板文件：template.xlsx / template(1).xlsx / assets/template.xlsx")


# ======================
# 自动识别文件类型（关键）
# ======================
def identify_excel_type(file_like) -> str:
    """
    返回:
      - 'negotiation'  谈判记录表
      - 'inbound'      入库资料信息表
      - 'unknown'
    """
    try:
        file_like.seek(0)
        wb = load_workbook(file_like, read_only=True, data_only=True)
        sheets = wb.sheetnames

        # 1) 强规则：sheet 名匹配
        if SHEET_NEG_DEFAULT in sheets:
            return "negotiation"

        # 2) 扫前 30 行：是否包含谈判关键字段
        for s in sheets[:3]:
            ws = wb[s]
            found_quote = False
            found_model = False
            for row in ws.iter_rows(min_row=1, max_row=30, min_col=1, max_col=80, values_only=True):
                for v in row:
                    if isinstance(v, str):
                        t = v.strip()
                        if "供应商报价（元/台）" in t:
                            found_quote = True
                        if t == "型号":
                            found_model = True
            if found_quote and found_model:
                return "negotiation"

        # 3) 扫前 80 行：是否像入库表结构化表头
        inbound_keys = ["CP型号", "电池容量", "屏幕尺寸", "主摄像头物理像素", "次摄像头物理像素"]
        for s in sheets[:3]:
            ws = wb[s]
            for r in range(1, min(81, ws.max_row + 1)):
                vals = []
                for c in range(1, min(ws.max_column, 200) + 1):
                    v = ws.cell(r, c).value
                    if isinstance(v, str) and v.strip():
                        vals.append(v.strip())
                if not vals:
                    continue
                hits = 0
                for k in inbound_keys:
                    if any(k in x for x in vals):
                        hits += 1
                if hits >= 2:
                    return "inbound"

    except Exception:
        pass

    return "unknown"


def split_two_files(files) -> Tuple:
    if len(files) != 2:
        raise RuntimeError("请一次上传 2 个文件：谈判记录表 + 入库资料信息表（顺序随意）。")

    f1, f2 = files[0], files[1]
    t1 = identify_excel_type(f1)
    t2 = identify_excel_type(f2)

    f1.seek(0); f2.seek(0)

    if t1 == "negotiation" and t2 == "inbound":
        return f1, f2
    if t2 == "negotiation" and t1 == "inbound":
        return f2, f1

    raise RuntimeError(
        f"文件类型识别失败：文件1={t1}，文件2={t2}。\n"
        "请确认：\n"
        "1）谈判表包含 sheet=产品谈判记录表 或 字段“供应商报价（元/台）/型号”。\n"
        "2）入库表包含 CP型号/电池容量/屏幕尺寸 等表头。"
    )


# ======================
# 读取谈判表：主数据（型号/价格/零售价/品牌）
# ======================
def read_negotiation_items(neg_file_like, sheet: str = SHEET_NEG_DEFAULT) -> pd.DataFrame:
    neg_file_like.seek(0)
    df = pd.read_excel(neg_file_like, sheet_name=sheet, header=3)

    # 取有效行：供应商报价非空
    df_items = df[df["供应商报价（元/台）"].notna()].copy()

    # token 给规格用，但价格/数量必须用“本行”
    df_items["__token__"] = df_items["型号"].astype(str).map(extract_model_token)

    return df_items


# ======================
# 读取谈判表数量：K-P 列六家公司（按 每一行型号 + 每一列供应商）
# ======================
def build_qty_map_and_suppliers(neg_file_like) -> Tuple[Dict[Tuple[str, str], int], List[str]]:
    neg_file_like.seek(0)
    wb = load_workbook(neg_file_like, data_only=True)
    ws = wb[SHEET_NEG_DEFAULT]

    header_row = None
    model_col = None
    for r in range(1, 31):
        for c in range(1, 80):
            v = ws.cell(r, c).value
            if isinstance(v, str) and norm_text(v) == "型号":
                header_row = r
                model_col = c
                break
        if header_row:
            break
    if not header_row or not model_col:
        raise RuntimeError("谈判记录表中未找到表头“型号”，请确认表结构未变。")

    supplier_cols = list(range(11, 17))  # K..P
    suppliers_raw = [norm_text(ws.cell(header_row, c).value) for c in supplier_cols]
    suppliers = [s for s in suppliers_raw if s]
    if len(suppliers) == 0:
        raise RuntimeError("谈判记录表 K-P 列未识别到供应商名称（表头行）。")

    qty_map: Dict[Tuple[str, str], int] = {}
    for r in range(header_row + 1, ws.max_row + 1):
        model = norm_text(ws.cell(r, model_col).value)
        if not model:
            continue
        for idx, c in enumerate(supplier_cols):
            if idx >= len(suppliers):
                break
            supplier = suppliers[idx]
            v = ws.cell(r, c).value
            if v is None or str(v).strip() == "":
                continue
            try:
                qty = int(float(v))
            except Exception:
                continue
            qty_map[(model, supplier)] = qty

    return qty_map, suppliers


# ======================
# 入库表：结构化表识别（按 token 建规格映射）
# ======================
def try_parse_inbound_as_table(inbound_file_like) -> Tuple[Dict[str, dict], List[dict]]:
    inbound_file_like.seek(0)
    wb = load_workbook(inbound_file_like, read_only=True, data_only=True)

    model_headers = ["型号", "终端型号", "产品型号", "机型", "终端型号/机型"]
    col_variants = {
        "cpu": ["CP型号", "CPU型号", "CPU"],
        "cam_main": ["主摄像头物理像素（万像素）", "主摄像头物理像素", "主摄像头像素（万像素）", "主摄像头像素"],
        "cam_sub": ["次摄像头物理像素（万像素）", "次摄像头物理像素", "副摄像头像素（万像素）", "次摄像头像素"],
        "screen": ["屏幕尺寸（英寸）", "屏幕尺寸", "屏幕尺寸(英寸)"],
        "battery": ["电池容量（mAH）", "电池容量", "电池容量(mAh)", "电池容量（mAh）"],
        "net": ["终端制式（TD-LTE/TD-SCDMA）", "终端制式", "网络制式", "制式"],
    }

    specs_map: Dict[str, dict] = {}
    debug_rows: List[dict] = []

    for sheet in wb.sheetnames:
        ws = wb[sheet]

        header_row = None
        header_values: Dict[str, int] = {}

        # 找表头行：前80行内包含“型号”且至少2个关键列
        for r in range(1, min(81, ws.max_row + 1)):
            row_texts = {}
            for c in range(1, min(ws.max_column, 200) + 1):
                v = ws.cell(r, c).value
                if isinstance(v, str) and v.strip():
                    row_texts[v.strip()] = c

            model_col = None
            for mh in model_headers:
                if mh in row_texts:
                    model_col = row_texts[mh]
                    break
            if not model_col:
                continue

            found = {"model": model_col}
            hit_count = 0
            for key, names in col_variants.items():
                for nm in names:
                    if nm in row_texts:
                        found[key] = row_texts[nm]
                        hit_count += 1
                        break

            if hit_count >= 2:
                header_row = r
                header_values = found
                break

        if not header_row:
            continue

        model_col = header_values["model"]
        for r in range(header_row + 1, ws.max_row + 1):
            mval = read_cell(ws, r, model_col)
            if mval is None or str(mval).strip() == "":
                continue

            token = extract_model_token(str(mval))
            if not token:
                continue

            sp = {}
            for key in ["cpu", "cam_main", "cam_sub", "screen", "battery", "net"]:
                c = header_values.get(key)
                sp[key] = read_cell(ws, r, c) if c else None

            non_empty = sum(
                1 for k in ["cpu", "screen", "battery", "cam_main", "cam_sub"]
                if sp.get(k) not in [None, ""]
            )

            # 至少2个字段有值才收录，避免误命中无关表
            if non_empty >= 2:
                # 同 token 多次出现：保留字段更全的
                if token not in specs_map:
                    specs_map[token] = sp
                else:
                    old = specs_map[token]
                    old_non_empty = sum(
                        1 for k in ["cpu", "screen", "battery", "cam_main", "cam_sub"]
                        if old.get(k) not in [None, ""]
                    )
                    if non_empty > old_non_empty:
                        specs_map[token] = sp

                debug_rows.append({"mode": "table", "sheet": sheet, "token": token, "score": non_empty, **sp})

    return specs_map, debug_rows


def format_common_fields(specs: dict):
    cpu = specs.get("cpu")
    cam_main = specs.get("cam_main")
    cam_sub = specs.get("cam_sub")

    camera = None
    if cam_main or cam_sub:
        main_txt = str(cam_main).strip() if cam_main is not None else ""
        sub_txt = str(cam_sub).strip() if cam_sub is not None else ""
        if main_txt and sub_txt:
            camera = f"主摄{main_txt}，次摄{sub_txt}"
        elif main_txt:
            camera = f"主摄{main_txt}"
        elif sub_txt:
            camera = f"次摄{sub_txt}"

    screen = specs.get("screen")
    screen_txt = f"{screen}英寸" if screen is not None and str(screen).strip() != "" else None

    battery = specs.get("battery")
    battery_txt = str(battery).strip() if battery is not None and str(battery).strip() != "" else None

    net_raw = specs.get("net")
    net_txt = None
    if isinstance(net_raw, str):
        up = net_raw.upper()
        if "NR" in up or "5" in up:
            net_txt = "5G"
        elif "LTE" in up or "4" in up:
            net_txt = "4G"
        else:
            net_txt = net_raw.strip()

    return cpu, camera, screen_txt, battery_txt, net_txt


# ======================
# 核心：写入模板（规格 + 价格 + 数量 联动）
# ======================
def fill_template(
    template_stream: io.BytesIO,
    df_items: pd.DataFrame,
    specs_map: Dict[str, dict],
    debug_rows: List[dict],
    qty_map: Dict[Tuple[str, str], int],
    suppliers: List[str],
) -> bytes:
    wb = load_workbook(template_stream)
    ws = wb[SHEET_OUT]

    # 找表头行（包含 品牌/型号/CPU型号）
    header_row = None
    for r in range(1, 80):
        row_vals = [ws.cell(r, c).value for c in range(1, 150)]
        if "品牌" in row_vals and "型号" in row_vals and "CPU型号" in row_vals:
            header_row = r
            break
    if header_row is None:
        raise RuntimeError("未找到模板表头行（包含“品牌/型号/CPU型号”）。请确认模板未变更。")

    header_to_col = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if isinstance(v, str) and v.strip():
            header_to_col[v.strip()] = c

    start_row = header_row + 1
    example_row = start_row

    model_col = header_to_col.get("型号")
    if not model_col:
        raise RuntimeError("模板中找不到“型号”列。")

    per_item_rows = max(1, len(suppliers))
    total_needed_rows = len(df_items) * per_item_rows

    # 模板已有可用行数（以型号列连续非空判断）
    last = start_row - 1
    for r in range(start_row, ws.max_row + 1):
        if ws.cell(r, model_col).value is None:
            break
        last = r
    existing = max(1, last - start_row + 1)

    # 不够行：插入并复制样式
    if existing < total_needed_rows:
        ws.insert_rows(start_row + existing, amount=total_needed_rows - existing)
        for i in range(existing, total_needed_rows):
            tgt_r = start_row + i
            for c in range(1, ws.max_column + 1):
                src = ws.cell(example_row, c)
                tgt = ws.cell(tgt_r, c)
                if isinstance(tgt, MergedCell):
                    continue
                tgt._style = copy(src._style)
                tgt.number_format = src.number_format
                tgt.font = copy(src.font)
                tgt.border = copy(src.border)
                tgt.fill = copy(src.fill)
                tgt.alignment = copy(src.alignment)
                tgt.protection = copy(src.protection)
                tgt.comment = None
            ws.row_dimensions[tgt_r].height = ws.row_dimensions[example_row].height

    # 多余行：不删，只清空“我们会写的字段 + S列供应商”，避免公式/格式被破坏
    if existing > total_needed_rows:
        for r in range(start_row + total_needed_rows, start_row + existing):
            safe_set(ws.cell(r, SUPPLIER_COL_S), None)
            for h in WRITE_FIELDS:
                c = header_to_col.get(h)
                if c:
                    safe_set(ws.cell(r, c), None)

    def setv(r: int, header: str, value):
        if header not in WRITE_FIELDS:
            return
        c = header_to_col.get(header)
        if c:
            safe_set(ws.cell(r, c), value)

    # === 以“谈判表每一行”为主键填充 ===
    for i, (_, row) in enumerate(df_items.iterrows()):
        model_raw = row.get("型号")
        model_text = norm_text(model_raw)
        token = norm_text(row.get("__token__"))

        # 规格：按 token 从入库表映射取
        specs = specs_map.get(token, {})
        cpu, camera, screen_txt, battery_txt, net_txt = format_common_fields(specs)

        # 价格：严格用谈判表“当前行”
        price_buy = row.get("供应商报价（元/台）")
        price_retail = row.get("零售价")

        for j, supplier in enumerate(suppliers):
            r = start_row + i * per_item_rows + j

            # S列：供应商名称
            safe_set(ws.cell(r, SUPPLIER_COL_S), supplier)

            # 基本字段
            setv(r, "品牌", row.get("品牌"))
            setv(r, "型号", model_raw)
            setv(r, "品类", "手机")

            # 规格（识别到才填，避免串型号）
            if cpu: setv(r, "CPU型号", cpu)
            if net_txt: setv(r, "网络制式", net_txt)
            if camera: setv(r, "摄像头", camera)
            if screen_txt: setv(r, "屏幕", screen_txt)
            if battery_txt: setv(r, "电池", battery_txt)

            # 价格（本行）
            if pd.notna(price_buy):
                setv(r, "预计采购票面价（元）", float(price_buy))
            if pd.notna(price_retail):
                setv(r, "预计零售价（元）", float(price_retail))

            # 数量：谈判表“这一型号行” + “这一供应商列” => 合同预计数量
            qty = qty_map.get((model_text, supplier))
            setv(r, "合同预计数量（台）", qty if qty is not None else None)

    # debug sheet：便于你核对入库识别
    try:
        if "debug_入库识别" in wb.sheetnames:
            ws_dbg = wb["debug_入库识别"]
            wb.remove(ws_dbg)
        ws_dbg = wb.create_sheet("debug_入库识别")
        ws_dbg.append(["mode", "sheet", "token", "score", "cpu", "cam_main", "cam_sub", "screen", "battery", "net"])
        for d in debug_rows:
            ws_dbg.append([
                d.get("mode"), d.get("sheet"), d.get("token"), d.get("score"),
                d.get("cpu"), d.get("cam_main"), d.get("cam_sub"),
                d.get("screen"), d.get("battery"), d.get("net")
            ])
    except Exception:
        pass

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ======================
# 页面：上传两个文件（顺序随意）
# ======================
uploaded_files = st.file_uploader(
    "上传 2 个 Excel（谈判记录表 + 入库资料信息表），顺序随意",
    type=["xlsx"],
    accept_multiple_files=True
)

run_btn = st.button("🚀 一键生成", type="primary")

if run_btn:
    try:
        if not uploaded_files or len(uploaded_files) != 2:
            st.warning("请一次上传 2 个文件（谈判记录表 + 入库资料信息表）")
            st.stop()

        # 自动拆分
        neg_file, inbound_file = split_two_files(uploaded_files)
        st.info(f"识别结果：谈判表 = {neg_file.name} ｜ 入库表 = {inbound_file.name}")

        # 读取谈判表：每一行决定价格与型号变体
        neg_file.seek(0)
        df_items = read_negotiation_items(neg_file)

        # 读取数量：K-P 六供应商
        neg_file.seek(0)
        qty_map, suppliers = build_qty_map_and_suppliers(neg_file)

        # 读取入库规格：按 token 映射
        inbound_file.seek(0)
        specs_map, debug_rows = try_parse_inbound_as_table(inbound_file)
        if not specs_map:
            st.error("入库资料表未识别到结构化规格表头（型号/CP型号/电池容量/屏幕尺寸等）。")
            st.stop()

        # 模板（仓库内置）
        tpl_path = find_template_path()
        template_bytes = tpl_path.read_bytes()
        tpl_stream = io.BytesIO(template_bytes)

        # 生成
        out_bytes = fill_template(tpl_stream, df_items, specs_map, debug_rows, qty_map, suppliers)

        ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"【生成】产品引入详细信息及风险评估_{ts}.xlsx"

        st.success("✅ 生成成功！请下载：")
        st.download_button(
            "⬇️ 下载结果文件",
            data=out_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        with st.expander("识别概览（核对用）"):
            st.write("谈判表有效行数（即输出型号版本数）：", len(df_items))
            st.write("识别到的供应商数（K-P）：", len(suppliers))
            st.write("入库表识别到的型号 token 数：", len(specs_map))
            st.write("示例 token：", sorted(list(specs_map.keys()))[:30])

    except Exception as e:
        st.error("运行失败（请看详细报错）")
        st.exception(e)
