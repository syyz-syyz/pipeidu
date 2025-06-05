import streamlit as st
import pandas as pd
from thefuzz import fuzz
import io


def main():
    st.title("文件数据比较与匹配")
    
    # 新增一个选项，让用户选择模式
    mode = st.radio("选择操作模式", ["标准模式", "高级模式"])
    
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
        
        # 初始化高级模式变量
        advanced_column1 = None
        advanced_column2 = None
        
        # 高级模式下的额外选项
        if mode == "高级模式":
            st.subheader("高级筛选设置")
            advanced_column1 = st.selectbox("选择第一个文件的筛选列", df1.columns, key="adv_col1")
            advanced_column2 = st.selectbox("选择第二个文件的筛选列", df2.columns, key="adv_col2")
            st.info("只有这两列内容相同的行才会进行匹配度分析")

        if st.button("比较数据"):
            final_matches = []
            
            # 创建进度条
            progress_bar = st.progress(0)
            total = len(df1[column1])
            
            for i, value1 in enumerate(df1[column1]):
                best_match = None
                best_score = 0
                common_chars = ""
                
                # 获取当前行的筛选列值（高级模式）
                current_advanced_value1 = df1.loc[i, advanced_column1] if (mode == "高级模式" and advanced_column1 is not None) else None
                
                for idx, value2 in df2[column2].items():
                    # 高级模式下的筛选条件
                    if mode == "高级模式" and advanced_column1 is not None and advanced_column2 is not None:
                        current_advanced_value2 = df2.loc[idx, advanced_column2]
                        # 如果筛选列的值不匹配，则跳过此行
                        if current_advanced_value1 != current_advanced_value2:
                            continue
                    
                    # 去除字符串中的空格
                    clean_value1 = str(value1).replace(" ", "")
                    clean_value2 = str(value2).replace(" ", "")
                    
                    # 计算匹配度
                    score = fuzz.ratio(clean_value1, clean_value2)
                    
                    # 找出相同的字符
                    current_common = ''.join(sorted(set(clean_value1) & set(clean_value2)))
                    
                    if score > best_score:
                        best_score = score
                        best_match = {
                            f"文件1_{column1}": value1,
                            f"文件2_{column2}": value2,
                            f"文件2_{mapping_column}": df2.loc[idx, mapping_column],
                            "匹配度": score,
                            "相同字符": current_common,
                            f"筛选_{advanced_column1}" if advanced_column1 is not None else "筛选_未选择列": current_advanced_value1 if mode == "高级模式" else "",
                            f"筛选_{advanced_column2}" if advanced_column2 is not None else "筛选_未选择列": current_advanced_value2 if (mode == "高级模式" and current_advanced_value1 is not None) else ""
                        }
                
                # 只添加非空的匹配结果
                if best_match is not None:
                    final_matches.append(best_match)
                else:
                    # 添加默认值，避免空结果
                    final_matches.append({
                        f"文件1_{column1}": value1,
                        f"文件2_{column2}": None,
                        f"文件2_{mapping_column}": None,
                        "匹配度": 0,
                        "相同字符": "",
                        f"筛选_{advanced_column1}" if advanced_column1 is not None else "筛选_未选择列": current_advanced_value1 if mode == "高级模式" else "",
                        f"筛选_{advanced_column2}" if advanced_column2 is not None else "筛选_未选择列": "" if mode == "高级模式" else ""
                    })
                
                # 更新进度条
                progress_bar.progress(int((i + 1) / total * 100))
            
            # 创建结果 DataFrame
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
