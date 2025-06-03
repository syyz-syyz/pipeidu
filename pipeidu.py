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
            # 创建与df1等长的结果列表
            final_matches = []
            
            # 确保df2的匹配列非空
            df2_valid = df2[df2[column2].notna()].copy()
            
            # 遍历df1的每一行
            for idx, row in df1.iterrows():
                value1 = row[column1]
                clean_value1 = str(value1).replace(" ", "")
                
                # 初始化匹配结果（默认匹配度0）
                match_result = {
                    f"文件1_{column1}": value1,
                    f"文件2_{column2}": None,
                    f"文件2_{mapping_column}": None,
                    "匹配度": 0
                }
                
                # 如果df2_valid为空，直接添加默认结果
                if df2_valid.empty:
                    final_matches.append(match_result)
                    continue
                
                # 计算与df2中所有值的匹配度
                scores = []
                for _, row2 in df2_valid.iterrows():
                    value2 = row2[column2]
                    clean_value2 = str(value2).replace(" ", "")
                    
                    # 计算相似度
                    score = fuzz.ratio(clean_value1, clean_value2)
                    scores.append((score, row2[column2], row2[mapping_column]))
                
                # 获取最高匹配度
                if scores:
                    best_score, best_value2, best_mapping = max(scores, key=lambda x: x[0])
                    
                    # 如果有匹配（分数大于0），更新结果
                    if best_score > 0:
                        match_result.update({
                            f"文件2_{column2}": best_value2,
                            f"文件2_{mapping_column}": best_mapping,
                            "匹配度": best_score
                        })
                
                # 无论是否匹配成功，都添加到结果列表
                final_matches.append(match_result)
            
            # 创建结果DataFrame
            result_df = pd.DataFrame(final_matches)
            
            # 显示统计信息
            total_records = len(result_df)
            matched_records = len(result_df[result_df['匹配度'] > 0])
            unmatched_records = total_records - matched_records
            
            st.write(f"总记录数: {total_records}")
            st.write(f"匹配成功: {matched_records}")
            st.write(f"匹配失败: {unmatched_records}")
            
            # 确保所有列显示（特别是匹配度为0的行）
            st.write("### 匹配结果详情")
            
            # 使用st.dataframe并设置height参数，确保所有行可见
            st.dataframe(
                result_df.reset_index(drop=True),
                height=min(600, total_records * 35)  # 根据记录数动态调整高度
            )
            
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
