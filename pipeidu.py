import streamlit as st
import pandas as pd
from thefuzz import fuzz
import io


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
            
            # 遍历第一个文件的每一行，确保结果与原数据一一对应
            for value1 in df1[column1]:
                best_match = {
                    f"文件1_{column1}": value1,
                    f"文件2_{column2}": None,
                    f"文件2_{mapping_column}": None,
                    "匹配度": 0  # 默认匹配度为0
                }
                best_score = 0
                
                # 在第二个文件中查找最佳匹配
                for idx, value2 in df2[column2].items():
                    clean_value1 = str(value1).replace(" ", "")
                    clean_value2 = str(value2).replace(" ", "")
                    score = fuzz.ratio(clean_value1, clean_value2)
                    
                    if score > best_score:
                        best_score = score
                        best_match = {
                            f"文件1_{column1}": value1,
                            f"文件2_{column2}": value2,
                            f"文件2_{mapping_column}": df2.loc[idx, mapping_column],
                            "匹配度": score
                        }
                
                final_matches.append(best_match)
            
            # 创建结果 DataFrame（行数与文件1的所选列完全一致）
            result_df = pd.DataFrame(final_matches)

            # 显示结果，不显示索引
            st.dataframe(result_df.reset_index(drop=True))

            # 提供导出功能
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False)
            output.seek(0)

            st.download_button(
                label="导出结果为 Excel",
                data=output,
                file_name='comparison_result.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )


if __name__ == "__main__":
    main()
