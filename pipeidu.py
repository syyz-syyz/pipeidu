import streamlit as st
import pandas as pd
from thefuzz import fuzz

def main():
    st.title("文件数据比较与匹配")

    # 上传两个文件
    file1 = st.file_uploader("上传第一个文件", type=["xlsx", "xls"])
    file2 = st.file_uploader("上传第二个文件", type=["xlsx", "xls"])

    if file1 and file2:
        # 读取文件
        df1 = pd.ExcelFile(file1).parse()
        df2 = pd.ExcelFile(file2).parse()

        # 获取所有表名
        sheet_names1 = pd.ExcelFile(file1).sheet_names
        sheet_names2 = pd.ExcelFile(file2).sheet_names

        # 选择 sheet
        sheet1 = st.selectbox("选择第一个文件的 sheet", sheet_names1)
        sheet2 = st.selectbox("选择第二个文件的 sheet", sheet_names2)

        # 读取对应 sheet 的数据
        df1 = pd.ExcelFile(file1).parse(sheet1)
        df2 = pd.ExcelFile(file2).parse(sheet2)

        # 选择列
        column1 = st.selectbox("选择第一个文件的列", df1.columns)
        column2 = st.selectbox("选择第二个文件用于匹配的列", df2.columns)
        mapping_column = st.selectbox("选择第二个文件的映射列", df2.columns)

        if st.button("比较数据"):
            final_matches = []
            for value1 in df1[column1]:
                best_match = None
                best_score = 0
                for idx, value2 in df2[column2].items():
                    score = fuzz.ratio(str(value1), str(value2))
                    if score > best_score:
                        best_score = score
                        best_match = {
                            f"{column1}": value1,
                            f"{column2}": value2,
                            f"{mapping_column}": df2.loc[idx, mapping_column],
                            "匹配度": score
                        }
                final_matches.append(best_match)

            # 创建结果 DataFrame
            result_df = pd.DataFrame(final_matches)

            # 显示结果，不显示索引
            st.dataframe(result_df.reset_index(drop=True))

            # 提供导出功能
            csv = result_df.to_csv(sep='\t', na_rep='nan')
            st.download_button(
                label="导出结果为 CSV",
                data=csv,
                file_name='comparison_result.csv',
                mime='text/csv',
            )

if __name__ == "__main__":
    main()
    
