import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title = "Excel 重複資料檢查工具",
    layout = "wide"
)
st.title("Excel 重複資料檢查工具")

st.info("您好我是Excel Duplicate Data Checker")

st.subheader("上傳Excel檔案")
upload_file = st.file_uploader(
    "選擇檔案!",
    type=["xlsx"]
)
if upload_file is not None:
    try:
        df = pd.read_excel(upload_file, dtype=str) #這個要記得不然0會消失
        st.success("上傳成功!")
        st.subheader("原始資料預覽")
        st.dataframe(df, use_container_width=True)

        st.write(f"資料總筆數{len(df)}")
        st.write(f"欄位數量{len(df.columns)}")

        st.divider()
        st.subheader("原始資料重複檢查結果")
        selected_columns = st.multiselect(
            "請選擇要比對的欄位",
            options=df.columns.tolist(),
            default=df.columns.tolist()
        )

        if selected_columns:
            duplicate_mask = df.duplicated(
                subset = selected_columns,
                keep=False
            )
            duplicate_df = df[duplicate_mask]
            duplicate_count = len(duplicate_df)
        else:
            duplicate_df = pd.DataFrame()
            duplicate_count = 0

        if not selected_columns:
            st.info("請至少選擇一個比對欄位")
        elif duplicate_count>0:
            st.warning(f"!!發現{duplicate_count}筆重複資料!!")
            st.dataframe(
                duplicate_df,
                use_container_width = True
            )
        else:
            st.success("沒有發現重複~")
              
    except Exception as e:
        st.error(f"讀取發生錯誤~{e}")

    st.divider()
    st.subheader("清除重複資料")

    if selected_columns:
        cleaned_df = df.drop_duplicates(
            subset = selected_columns,
            keep="first"
        )
        orginal_count = len(df)
        cleaned_count = len(cleaned_df)
        remove_count = orginal_count - cleaned_count
        c1, c2, c3 = st.columns(3)

        c1.metric("原始資料筆數", orginal_count)
        c2.metric("總共刪除重複筆數", remove_count)
        c3.metric("清除後剩餘筆數", cleaned_count)

        st.subheader("完成清理!以下資料預覽")
        st.dataframe(
            cleaned_df,
            use_container_width=True
        )
        output = BytesIO()           
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            cleaned_df.to_excel(
                writer,
                index=False,
                sheet_name="Cheaned Data"
            )
        excel_data = output.getvalue()

        st.download_button(
            label="清理完成，點我下載清理好的檔案!!",
            data = excel_data,
            file_name="cheaned_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("請先選擇需要比對的欄位，才能清理喔")