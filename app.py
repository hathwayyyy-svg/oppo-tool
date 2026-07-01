# -*- coding: utf-8 -*-
import io
import re
import datetime as dt
from copy import copy
from pathlib import Path
from typing import Dict, Tuple, List

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.formula.translate import Translator
from docx import Document

SHEET_OUT = "引入产品详细信息"
SHEET_NEG_DEFAULT = "产品谈判记录表"

WRITE_FIELDS = {
    "品牌", "型号", "品类", "CPU型号", "网络制式", "摄像头", "屏幕", "电池",
    "预计采购票面价（元）", "预计零售价（元）", "合同预计数量（台）",
}

SUPPLIER_COL_S = 19
MERGE_COL_Q = 17

st.set_page_config(page_title="OPPO 引入回填", layout="wide")
st.title("OPPO 引入回填（上传2个文件 → 一键生成 Excel + Word）")
st.caption("✅ 动态识别供应商公司列｜✅ 不误识别后续字段｜✅ Q列按相同3C合并｜✅ 自动生成Word")


def norm_text(x) -> str:
    if x is None:
        return ""
    return str(x).strip().replace("　", " ").replace("\u00a0", " ").strip()


def safe_set(cell, value):
    if isinstance(cell, MergedCell):
        return
    cell.value = value


def normalize_model_name(x) -> str:
    s = norm_text(x).upper()
    s = s.replace("（", "(").replace("）", ")")
    for bad in ["全网通", "移动", "联通", "电信", "分销公开版", "公开版", "定制", "TD-LTE", "LTE", "NR"]:
        s = s.replace(bad, "")
    return s.strip()


def extract_model_token(s: str) -> str:
    raw = norm_text(s).upper().replace("（", "(").replace("）", ")")
    bracket_matches = re.findall(r"\((.*?)\)", raw)
    for part in bracket_matches:
        m = re.search(r"[A-Z]{2,}\d+[A-Z0-9]*", part)
        if m:
            return m.group(0)

    s2 = normalize_model_name(s)
    matches = re.findall(r"[A-Z]{2,}\d+[A-Z0-9]*", s2)
    if matches:
        for m in matches:
            if m.startswith(("P", "O")):
                return m
        return matches[-1]
    return s2


def clean_model_for_output(x) -> str:
    s = norm_text(x)
    s = s.replace("分销公开版", "")
    s = re.sub(r"\s+", " ", s).strip()
    s = s.replace(" （", "（")
    return s


def find_excel_template_path() -> Path:
    for p in [Path("template.xlsx"), Path("template(1).xlsx"), Path("assets/template.xlsx")]:
        if p.exists():
            return p
    raise RuntimeError("仓库内未找到 Excel 模板：template.xlsx / template(1).xlsx / assets/template.xlsx")


def find_docx_template_path() -> Path:
    for p in [
        Path("关于2026年3月第二批产品引入的请示.docx"),
        Path("请示模板.docx"),
        Path("assets/关于2026年3月第二批产品引入的请示.docx"),
    ]:
        if p.exists():
            return p
    raise RuntimeError("仓库内未找到 Word 模板。")


def identify_excel_type(file_like) -> str:
    try:
        file_like.seek(0)
        wb = load_workbook(file_like, read_only=True, data_only=True)
        sheets = wb.sheetnames

        if SHEET_NEG_DEFAULT in sheets:
            return "negotiation"

        for s in sheets[:3]:
            ws = wb[s]
            found_quote = False
            found_model = False
            for row in ws.iter_rows(min_row=1, max_row=30, min_col=1, max_col=100, values_only=True):
                for v in row:
                    if isinstance(v, str):
                        t = v.strip()
                        if "供应商报价（元/台）" in t:
                            found_quote = True
                        if t == "型号":
                            found_model = True
            if found_quote and found_model:
                return "negotiation"

        inbound_keys = ["CP型号", "电池容量", "屏幕尺寸", "主摄像头物理像素", "次摄像头物理像素"]
        for s in sheets[:3]:
            ws = wb[s]
            for r in range(1, min(81, ws.max_row + 1)):
                vals = []
                for c in range(1, min(ws.max_column, 220) + 1):
                    v = ws.cell(r, c).value
                    if isinstance(v, str) and v.strip():
                        vals.append(v.strip())
                hits = sum(1 for k in inbound_keys if any(k in x for x in vals))
                if hits >= 2:
                    return "inbound"
    except Exception:
        pass

    return "unknown"


def split_two_files(files) -> Tuple:
    if len(files) != 2:
        raise RuntimeError("请一次上传 2 个文件：谈判记录表 + 入库资料信息表。")

    f1, f2 = files[0], files[1]
    t1, t2 = identify_excel_type(f1), identify_excel_type(f2)
    f1.seek(0)
    f2.seek(0)

    if t1 == "negotiation" and t2 == "inbound":
        return f1, f2
    if t2 == "negotiation" and t1 == "inbound":
        return f2, f1

    raise RuntimeError(f"识别失败：文件1={t1}, 文件2={t2}。请确认一个谈判表、一个入库表。")


def read_negotiation_with_rowid(neg_file_like) -> Tuple[pd.DataFrame, List[str], Dict[Tuple[int, str], int]]:
    neg_file_like.seek(0)
    wb = load_workbook(neg_file_like, data_only=True)
    ws = wb[SHEET_NEG_DEFAULT]

    header_row = None
    col_map = {}

    for r in range(1, 80):
        row_vals = [ws.cell(r, c).value for c in range(1, 180)]
        has_model = any(norm_text(v) == "型号" for v in row_vals)
        has_buy = any(isinstance(v, str) and "供应商报价（元/台）" in v for v in row_vals if isinstance(v, str))
        if has_model and has_buy:
            header_row = r
            for c in range(1, 180):
                v = ws.cell(r, c).value
                if isinstance(v, str) and v.strip():
                    col_map[v.strip()] = c
            break

    if not header_row:
        raise RuntimeError("谈判表找不到表头行（型号/供应商报价（元/台））。")

    c_brand = col_map.get("品牌")
    c_model = col_map.get("型号")
    c_buy = col_map.get("供应商报价（元/台）")
    c_retail = col_map.get("零售价") or col_map.get("建议零售价")

    if not (c_model and c_buy):
        raise RuntimeError("谈判表缺少必要列：型号 / 供应商报价（元/台）")

    # ✅ 动态识别供应商列：从 K 列开始，只识别公司名称列，遇到非公司字段停止
    supplier_cols = []
    suppliers = []

    STOP_WORDS = [
        "合计", "总计", "小计", "备注", "说明",
        "零售价", "供应商报价", "采购价", "价格",
        "数量合计", "金额", "毛利", "返利",
        "成本", "利润", "税率", "税额"
    ]

    c = 11  # K列开始

    while c <= ws.max_column:
        name = norm_text(ws.cell(header_row, c).value)

        if name == "":
            break

        if any(word in name for word in STOP_WORDS):
            break

        if "公司" in name or "有限公司" in name:
            supplier_cols.append(c)
            suppliers.append(name)
            c += 1
            continue

        break

    if not suppliers:
        raise RuntimeError("谈判表从 K 列开始未识别到供应商公司名称。")

    rows = []
    qty_by_row_supplier: Dict[Tuple[int, str], int] = {}

    for r in range(header_row + 1, ws.max_row + 1):
        model = ws.cell(r, c_model).value
        buy = ws.cell(r, c_buy).value

        if model is None or norm_text(model) == "":
            continue
        if buy is None or str(buy).strip() == "":
            continue

        brand = ws.cell(r, c_brand).value if c_brand else None
        retail = ws.cell(r, c_retail).value if c_retail else None

        rows.append({
            "品牌": brand,
            "型号": model,
            "供应商报价（元/台）": buy,
            "零售价": retail,
            "__row__": r,
            "__token__": extract_model_token(str(model)),
        })

        for idx, col in enumerate(supplier_cols):
            supplier = suppliers[idx]
            v = ws.cell(r, col).value
            if v is None or str(v).strip() == "":
                continue
            try:
                qty_by_row_supplier[(r, supplier)] = int(float(v))
            except Exception:
                continue

    df_items = pd.DataFrame(rows)
    if df_items.empty:
        raise RuntimeError("谈判表未读取到有效型号行。")

    return df_items, suppliers, qty_by_row_supplier


def try_parse_inbound_as_table(inbound_file_like) -> Tuple[Dict[str, dict], List[dict]]:
    inbound_file_like.seek(0)
    wb = load_workbook(inbound_file_like, read_only=True, data_only=True)

    model_headers = ["型号", "终端型号", "产品型号", "机型", "终端型号/机型"]
    col_variants = {
        "cpu": ["CP型号", "CPU型号", "CPU"],
        "cam_main": ["主摄像头物理像素（万像素）", "主摄像头物理像素", "主摄像头像素（万像素）"],
        "cam_sub": ["次摄像头物理像素（万像素）", "次摄像头物理像素", "副摄像头像素（万像素）"],
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

        for r in range(1, min(81, ws.max_row + 1)):
            row_texts = {}
            for c in range(1, min(ws.max_column, 240) + 1):
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
            hit = 0
            for key, names in col_variants.items():
                for nm in names:
                    if nm in row_texts:
                        found[key] = row_texts[nm]
                        hit += 1
                        break

            if hit >= 2:
                header_row = r
                header_values = found
                break

        if not header_row:
            continue

        mc = header_values["model"]
        for r in range(header_row + 1, ws.max_row + 1):
            mval = ws.cell(r, mc).value
            if mval is None or str(mval).strip() == "":
                continue

            token = extract_model_token(str(mval))

            sp = {}
            for key in ["cpu", "cam_main", "cam_sub", "screen", "battery", "net"]:
                col = header_values.get(key)
                sp[key] = ws.cell(r, col).value if col else None

            score = sum(1 for k in ["cpu", "screen", "battery", "cam_main", "cam_sub"] if sp.get(k) not in [None, ""])
            if score >= 2:
                old = specs_map.get(token)
                if not old:
                    specs_map[token] = sp
                else:
                    old_score = sum(1 for k in ["cpu", "screen", "battery", "cam_main", "cam_sub"] if old.get(k) not in [None, ""])
                    if score > old_score:
                        specs_map[token] = sp

                debug_rows.append({"sheet": sheet, "token": token, "score": score, **sp})

    return specs_map, debug_rows


def format_common_fields(specs: dict):
    cpu = norm_text(specs.get("cpu"))

    cam_main = specs.get("cam_main")
    cam_sub = specs.get("cam_sub")
    camera = ""
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
    screen_txt = f"{screen}英寸" if screen is not None and str(screen).strip() != "" else ""

    battery = specs.get("battery")
    battery_txt = str(battery).strip() if battery is not None and str(battery).strip() != "" else ""

    net_raw = specs.get("net")
    net_txt = ""
    if isinstance(net_raw, str):
        up = net_raw.upper()
        if "NR" in up or "5" in up:
            net_txt = "5G"
        elif "LTE" in up or "4" in up:
            net_txt = "4G"
        else:
            net_txt = net_raw.strip()

    return cpu, camera, screen_txt, battery_txt, net_txt


def derive_docx_product_names(df_items: pd.DataFrame) -> str:
    names = []
    seen = set()

    for _, row in df_items.iterrows():
        brand = norm_text(row.get("品牌"))
        model_raw = clean_model_for_output(row.get("型号"))

        main_name = model_raw.split("（")[0].split("(")[0].strip()
        main_name = re.sub(r"\s+", " ", main_name)

        display = main_name

        if brand.upper() == "OPPO":
            if main_name.upper().startswith("A"):
                display = f"OPPO {main_name}系列"
            else:
                display = main_name
        elif "一加" in brand or brand.upper() in ["ONEPLUS", "1+", "一加"]:
            display = main_name if main_name.startswith("一加") else f"一加{main_name}"
        else:
            if brand and not main_name.startswith(brand):
                display = f"{brand}{main_name}"

        if display not in seen:
            seen.add(display)
            names.append(display)

    return "、".join(names)


def derive_docx_price_range(df_items: pd.DataFrame) -> str:
    prices = pd.to_numeric(df_items["零售价"], errors="coerce").dropna()
    if prices.empty:
        return ""
    pmin = int(prices.min())
    pmax = int(prices.max())
    return f"{pmin}元" if pmin == pmax else f"{pmin}-{pmax}元"


def is_run_highlighted(run) -> bool:
    try:
        return run.font.highlight_color is not None
    except Exception:
        try:
            rPr = run._r.rPr
            return rPr is not None and rPr.highlight is not None
        except Exception:
            return False


def replace_highlight_groups_in_paragraph(paragraph, product_text: str, price_text: str, state: dict):
    runs = paragraph.runs
    groups = []
    cur = []

    for idx, run in enumerate(runs):
        txt = run.text or ""
        if txt and is_run_highlighted(run):
            cur.append(idx)
        else:
            if cur:
                groups.append(cur)
                cur = []
    if cur:
        groups.append(cur)

    for group in groups:
        group_text = "".join(runs[i].text for i in group).strip()
        if not group_text:
            continue

        replacement = price_text if re.search(r"\d+\s*-\s*\d+\s*元?", group_text) else product_text

        if replacement == product_text and state.get("product_done"):
            continue
        if replacement == price_text and state.get("price_done"):
            continue

        runs[group[0]].text = replacement
        for i in group[1:]:
            runs[i].text = ""

        if replacement == product_text:
            state["product_done"] = True
        if replacement == price_text:
            state["price_done"] = True


def fill_docx_template(docx_template_path: Path, df_items: pd.DataFrame) -> bytes:
    doc = Document(str(docx_template_path))
    product_names = derive_docx_product_names(df_items)
    price_range = derive_docx_price_range(df_items)
    state = {"product_done": False, "price_done": False}

    for p in doc.paragraphs:
        replace_highlight_groups_in_paragraph(p, product_names, price_range, state)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_highlight_groups_in_paragraph(p, product_names, price_range, state)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def fill_template(
    template_stream: io.BytesIO,
    df_items: pd.DataFrame,
    suppliers: List[str],
    qty_by_row_supplier: Dict[Tuple[int, str], int],
    specs_map: Dict[str, dict],
    debug_rows: List[dict],
) -> bytes:
    wb = load_workbook(template_stream)
    ws = wb[SHEET_OUT]

    header_row = None
    for r in range(1, 120):
        vals = [ws.cell(r, c).value for c in range(1, 240)]
        if "品牌" in vals and "型号" in vals and "CPU型号" in vals:
            header_row = r
            break
    if header_row is None:
        raise RuntimeError("模板找不到表头行。")

    header_to_col = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if isinstance(v, str) and v.strip():
            header_to_col[v.strip()] = c

    start_row = header_row + 1
    example_row = start_row
    model_col = header_to_col["型号"]

    total_needed_rows = len(df_items) * len(suppliers)

    end = start_row - 1
    for r in range(start_row, ws.max_row + 1):
        probe = [ws.cell(r, c).value for c in range(1, 11)]
        if all(v is None or str(v).strip() == "" for v in probe):
            break
        end = r
    existing = max(1, end - start_row + 1)

    if existing < total_needed_rows:
        ws.insert_rows(start_row + existing, amount=total_needed_rows - existing)

        write_cols = set()
        for h in WRITE_FIELDS:
            cc = header_to_col.get(h)
            if cc:
                write_cols.add(cc)
        write_cols.add(SUPPLIER_COL_S)

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

                if c not in write_cols:
                    v = src.value
                    if isinstance(v, str) and v.startswith("="):
                        try:
                            tgt.value = Translator(v, origin=src.coordinate).translate_formula(tgt.coordinate)
                        except Exception:
                            tgt.value = v
                    else:
                        tgt.value = v

            ws.row_dimensions[tgt_r].height = ws.row_dimensions[example_row].height

    if existing > total_needed_rows:
        ws.delete_rows(start_row + total_needed_rows, existing - total_needed_rows)

    def setv(r: int, header: str, value):
        if header not in WRITE_FIELDS:
            return
        cc = header_to_col.get(header)
        if cc:
            safe_set(ws.cell(r, cc), "" if value is None else value)

    for i, row in df_items.iterrows():
        model_raw = row.get("型号")
        model_output = clean_model_for_output(model_raw)
        brand = row.get("品牌")
        buy = row.get("供应商报价（元/台）")
        retail = row.get("零售价")
        token = norm_text(row.get("__token__"))
        row_id = int(row.get("__row__"))

        specs = specs_map.get(token, {})
        cpu, camera, screen_txt, battery_txt, net_txt = format_common_fields(specs)

        for j, supplier in enumerate(suppliers):
            r = start_row + i * len(suppliers) + j

            safe_set(ws.cell(r, SUPPLIER_COL_S), supplier)

            setv(r, "品牌", brand)
            setv(r, "型号", model_output)
            setv(r, "品类", "手机")
            setv(r, "CPU型号", cpu)
            setv(r, "网络制式", net_txt)
            setv(r, "摄像头", camera)
            setv(r, "屏幕", screen_txt)
            setv(r, "电池", battery_txt)
            setv(r, "预计采购票面价（元）", float(buy) if pd.notna(buy) else "")
            setv(r, "预计零售价（元）", float(retail) if pd.notna(retail) else "")

            qty = qty_by_row_supplier.get((row_id, supplier))
            setv(r, "合同预计数量（台）", qty if qty is not None else "")

    def merge_q_by_token():
        q_col = MERGE_COL_Q
        first = start_row
        last = start_row + total_needed_rows - 1

        to_remove = []
        for rng in list(ws.merged_cells.ranges):
            if rng.min_col == q_col and rng.max_col == q_col:
                if not (rng.max_row < first or rng.min_row > last):
                    to_remove.append(rng)
        for rng in to_remove:
            try:
                ws.unmerge_cells(str(rng))
            except Exception:
                pass

        def get_token(r: int) -> str:
            return extract_model_token(norm_text(ws.cell(r, model_col).value))

        r = first
        while r <= last:
            tk = get_token(r)
            if not tk:
                r += 1
                continue

            r2 = r
            while r2 + 1 <= last and get_token(r2 + 1) == tk:
                r2 += 1

            if r2 > r:
                top_val = ws.cell(r, q_col).value
                ws.merge_cells(start_row=r, start_column=q_col, end_row=r2, end_column=q_col)
                ws.cell(r, q_col).value = top_val
            r = r2 + 1

    merge_q_by_token()

    try:
        if "debug_入库识别" in wb.sheetnames:
            wb.remove(wb["debug_入库识别"])
        ws_dbg = wb.create_sheet("debug_入库识别")
        ws_dbg.append(["sheet", "token", "score", "cpu", "cam_main", "cam_sub", "screen", "battery", "net"])
        for d in debug_rows:
            ws_dbg.append([
                d.get("sheet"), d.get("token"), d.get("score"),
                d.get("cpu"), d.get("cam_main"), d.get("cam_sub"),
                d.get("screen"), d.get("battery"), d.get("net"),
            ])
    except Exception:
        pass

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


uploaded_files = st.file_uploader(
    "上传 2 个 Excel（谈判记录表 + 入库资料信息表），顺序随意",
    type=["xlsx"],
    accept_multiple_files=True
)

run_btn = st.button("🚀 一键生成 Excel + Word", type="primary")

if run_btn:
    try:
        if not uploaded_files or len(uploaded_files) != 2:
            st.warning("请一次上传 2 个文件")
            st.stop()

        neg_file, inbound_file = split_two_files(uploaded_files)
        st.info(f"识别结果：谈判表 = {neg_file.name} ｜ 入库表 = {inbound_file.name}")

        df_items, suppliers, qty_by_row_supplier = read_negotiation_with_rowid(neg_file)
        specs_map, debug_rows = try_parse_inbound_as_table(inbound_file)

        if not specs_map:
            st.error("入库资料表没有识别到规格表头。")
            st.stop()

        excel_tpl_path = find_excel_template_path()
        excel_bytes = fill_template(
            io.BytesIO(excel_tpl_path.read_bytes()),
            df_items,
            suppliers,
            qty_by_row_supplier,
            specs_map,
            debug_rows
        )

        docx_bytes = None
        try:
            docx_tpl_path = find_docx_template_path()
            docx_bytes = fill_docx_template(docx_tpl_path, df_items)
        except Exception:
            docx_bytes = None

        ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        st.success("✅ 生成成功！")

        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                "⬇️ 下载 Excel 结果文件",
                data=excel_bytes,
                file_name=f"【生成】产品引入详细信息及风险评估_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        with c2:
            if docx_bytes:
                st.download_button(
                    "⬇️ 下载 Word 请示文件",
                    data=docx_bytes,
                    file_name=f"【生成】关于产品引入的请示_{ts}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
            else:
                st.info("未检测到 Word 模板，仅生成 Excel。")

        with st.expander("核对信息"):
            st.write("供应商数量：", len(suppliers))
            st.write("供应商列表：", suppliers)
            st.write("谈判表有效行数：", len(df_items))
            st.write("入库表识别 token 数：", len(specs_map))

    except Exception as e:
        st.error("运行失败")
        st.exception(e)
