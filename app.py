import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.title("OPPO 产品引入自动生成系统")

neg_file = st.file_uploader("上传谈判记录表", type=["xlsx"])
inbound_file = st.file_uploader("上传入库资料表", type=["xlsx"])

if st.button("一键生成"):

    if not neg_file or not inbound_file:
        st.warning("请先上传两个文件")
        st.stop()

    # 读取谈判表
    df_neg = pd.read_excel(neg_file, sheet_name="产品谈判记录表", header=3)

    # 读取六家公司列（K-P）
    supplier_cols = df_neg.columns[10:16]

    df_unpivot = df_neg.melt(
        id_vars=["品牌","型号","供应商报价（元/台）","零售价"],
        value_vars=supplier_cols,
        var_name="供应商名称",
        value_name="合同预计数量（台）"
    )

    df_unpivot = df_unpivot[df_unpivot["合同预计数量（台）"].notna()]

    # 读取入库资料
    df_in = pd.read_excel(inbound_file)

    def get_value(df, keyword):
    df = df.astype(str)
    match = df[df.apply(lambda row: row.astype(str).str.contains(keyword).any(), axis=1)]
    if len(match) == 0:
        return ""
    return match.iloc[0,1]

    cpu = get_value(df_in, "CP型号")
    screen = get_value(df_in, "屏幕尺寸")
    battery = get_value(df_in, "电池容量")
    cam_main = get_value(df_in, "主摄")
    cam_sub = get_value(df_in, "次摄")

    df_unpivot["CPU型号"] = cpu
    df_unpivot["屏幕"] = str(screen) + "英寸"
    df_unpivot["电池"] = battery
    df_unpivot["摄像头"] = "主摄" + str(cam_main) + "，次摄" + str(cam_sub)

    output = BytesIO()
    df_unpivot.to_excel(output, index=False)

    st.download_button(
        label="下载生成文件",
        data=output.getvalue(),
        file_name=f"产品引入结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
