import streamlit as st
import pandas as pd
from datetime import datetime

# Streamlit title
st.title("醫療單項成本分析 - 資料處理平台")

# Streamlit 描述
st.write("""
這是一個用於進行 B1 單項成本分析的應用程式，您可以上傳相關 Excel 檔案，進行數據處理，並下載結果。
請依照順序上傳以下檔案：
1. 分項成本.xlsx  
2. 開刀房時間.xlsx  
3. 醫師薪資.xlsx  
4. 手術檔.xlsx  
5. 醫師抽成費.xlsx  
6. 設備使用檔.xlsx  
7. 手術材料檔.xlsx  
""")

# 文件上傳
uploaded_files = {}
file_names = [
    "1.分項成本.xlsx",
    "2.開刀房時間.xlsx",
    "3.醫師薪資.xlsx",
    "4.手術檔.xlsx",
    "5.醫師抽成費.xlsx",
    "6.設備使用檔.xlsx",
    "7.手術材料檔.xlsx"
]

for name in file_names:
    uploaded_files[name] = st.file_uploader(f"請上傳檔案: {name}", type=["xlsx"])

if st.button("開始處理數據"):
    # 確保所有檔案已上傳
    if None in uploaded_files.values():
        st.error("請確保所有檔案均已上傳！")
    else:
        # 讀取上傳的資料
        st.info("正在讀取資料...")
        df1 = pd.read_excel(uploaded_files[file_names[0]], sheet_name='成本1-52')
        df2 = pd.read_excel(uploaded_files[file_names[1]])
        df3 = pd.read_excel(uploaded_files[file_names[2]])
        df4 = pd.read_excel(uploaded_files[file_names[3]], sheet_name='總合併')
        df5 = pd.read_excel(uploaded_files[file_names[4]])
        df6 = pd.read_excel(uploaded_files[file_names[5]])
        df7 = pd.read_excel(uploaded_files[file_names[6]], sheet_name='材料利潤')

        # 清理數據（去除欄位名稱空格）
        for df in [df1, df2, df3, df4, df5, df6, df7]:
            df.columns = df.columns.str.strip()

        df1["病患姓名"] = df1["病患姓名"].str.strip()
        df2["病患姓名"] = df2["病患姓名"].str.strip()
        df4["病歷號"] = df4["病歷號"].str.strip()
        df7["病歷號"] = df7["病歷號"].str.strip()

        df1["手術院碼"] = pd.to_numeric(df1["手術院碼"], errors="coerce").fillna(0).astype(int)
        df5["手術院碼"] = df5["手術院碼"].fillna(0).astype(int)
        df6["手術院碼"] = df6["手術院碼"].fillna(0).astype(int)

        # 合併數據
        st.info("正在清理與合併資料...")
        merged_data1 = pd.merge(df1, df2, on="病患姓名")
        merged_data1 = pd.merge(merged_data1, df5, on="手術院碼")
        merged_data1 = pd.merge(merged_data1, df6, on="手術院碼")
        merged_data1 = pd.merge(merged_data1, df4, on=["病歷號", "醫師", "手術院碼"])
        merged_data1 = pd.merge(merged_data1, df7, on=["病歷號", "手術院碼"], how="left")
        merged_data1 = pd.merge(merged_data1, df3, on="醫師")

        # 計算成本
        st.info("正在計算成本...")
        merged_data1["醫師固定薪成本"] = merged_data1["醫師時間2"] * merged_data1["醫師每分鐘人力成本"]
        merged_data1["刷手及流動護士成本"] = merged_data1["刷手及流動護士"] * 13.31
        merged_data1["外科助手成本"] = merged_data1["外科助手"] * 12.88
        merged_data1["恢復室成本"] = merged_data1["恢復室"] * 10.48
        merged_data1["行政人員"] = merged_data1["參數"] * 100
        merged_data1["用人成本合計"] = merged_data1["醫師抽成費"] + merged_data1["醫師固定薪成本"] + \
                                 merged_data1["刷手及流動護士成本"] + merged_data1["外科助手成本"] + \
                                 merged_data1["恢復室成本"] + merged_data1["行政人員"]
        merged_data1["設備折舊成本"] = merged_data1["折舊時間2"] * merged_data1["設備折舊"]
        merged_data1["房屋折舊成本"] = merged_data1["折舊時間2"] * 0.71
        merged_data1["總折舊成本"] = merged_data1["設備折舊成本"] + merged_data1["房屋折舊成本"]
        merged_data1["維修費用"] = merged_data1["總折舊成本"] * 0.18
        merged_data1["設施設備費用合計"] = merged_data1["總折舊成本"] + merged_data1["維修費用"]
        merged_data1["直接成本合計"] = merged_data1["用人成本合計"] + merged_data1["藥品醫材成本合計"] + merged_data1["設施設備費用合計"]
        merged_data1["作業成本"] = merged_data1["直接成本合計"] * 0.15
        merged_data1["行政管理成本"] = merged_data1["直接成本合計"] * 0.05
        merged_data1["成本總計"] = merged_data1["直接成本合計"] + merged_data1["作業成本"] + merged_data1["行政管理成本"]

        # 重新排列列順序（已修正拼寫錯誤）
        columns_order = [ '手術院碼','病患姓名','人數','病歷號','日期','健保收入','健保點值(6%)', '健保收入淨額', '醫師', '醫師抽成費', '醫師固定薪成本',
                         '刷手及流動護士成本','外科助手成本', '恢復室成本', '行政人員', '用人成本合計', '特材費', '藥費', '藥品醫材成本合計',
                         '設備折舊成本', '房屋折舊成本', '維修費用', '設施設備費用合計', '直接成本合計', '作業成本',
                         '行政管理成本', '成本總計','健保材料收入','自費材料收入','手術材料收入合計','材料成本','健保材料點值(6%)','材料淨利潤']
        missing_columns = [col for col in columns_order if col not in merged_data1.columns]
        if missing_columns:
            st.warning(f"以下欄位缺失: {missing_columns}")
        else:
            merged_data1 = merged_data1[columns_order]

        # 填充空值並四捨五入
        numeric_columns = merged_data1.select_dtypes(include=['float64', 'int64']).columns
        merged_data1[numeric_columns] = merged_data1[numeric_columns].fillna(0).round(0).astype(int)
        merged_data1['日期'] = pd.to_datetime(merged_data1['日期']).dt.strftime('%y%m%d')

        # 產生 Excel 檔案
        st.info("正在生成 Excel 檔案...")
        output_file = "產出報表1-報表版(全部院碼).xlsx"
        merged_data1.to_excel(output_file, index=False, sheet_name="報表結果")

        # 提供下載連結
        with open(output_file, "rb") as file:
            st.download_button(
                label="下載結果報表",
                data=file,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
