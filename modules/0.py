import pandas as pd
import os
import re


def process_excel(input_file, schedule_file, output_file, month_column="班次"):
    """
    处理Excel文件：
    1. 删除前四行
    2. 保留第一列和第27列开始的列
    3. 第一列表头改为"姓名"
    4. 在第一列后插入三列空白列，表头为"员工ID"、"部门"和"班次"
    5. 从班次.xlsx中根据姓名匹配并填充上述三列数据
    6. 按顺序将指定字段内容替换为空（支持*模糊匹配，非整单元格匹配）

    参数:
        input_file: 输入Excel文件路径
        schedule_file: 班次Excel文件路径
        output_file: 输出Excel文件路径，默认为在输入文件名后加"_processed"
    """
    try:
        # 自动生成输出文件名
        output_file = output_file

        # 读取班次信息
        try:
            schedule_df = pd.read_excel(schedule_file)
            required_columns = ['姓名', '员工ID', '部门', month_column]
            if not set(required_columns).issubset(schedule_df.columns):
                missing = [col for col in required_columns if col not in schedule_df.columns]
                raise ValueError(f"班次文件缺少必要的列: {', '.join(missing)}")
        except Exception as e:
            print(f"读取班次文件出错: {str(e)}")
            return None

        # 读取主Excel文件
        excel_file = pd.ExcelFile(input_file)
        sheet_names = excel_file.sheet_names

        # 替换模式列表（全部为非单元格匹配，按优先级排序）
        patterns_to_replace = [
            # 1. 带分号的缺卡格式（增强匹配）
            r'缺卡\([^)]*\);',  # 匹配"缺卡(任意内容);"
            r'缺卡\(.*?\);',  # 备用模式，确保匹配
            # 2. 不带分号的缺卡格式
            r'缺卡\([^)]*\)',
            r'缺卡\(.*?\)',
            # 3. 补卡申请格式
            r'补卡申请（[^）]*）',
            r'补卡申请（.*?）',
            # 4. 正常(补卡)格式
            r'正常\(补卡\)-',
            # 5. 正常格式
            r'正常-',
            # 6. 双横线格式（多种可能的横线）
            r'--',
            r'— —',  # 全角横线
            r'——',  # 破折号
            # 7. 单独的缺卡
            r'缺卡',
            # 8. 各种换行符和空白字符
            r'\r\n|\r|\n|\t',
            # 9. 空格（多个连续空格）
            r' +'
        ]

        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for sheet_name in sheet_names:
                # 读取数据，不设表头
                df = excel_file.parse(sheet_name, header=None)

                # 删除前四行
                if len(df) > 4:
                    df = df[4:].reset_index(drop=True)

                    # 保留第一列和第27列及以后
                    if len(df.columns) >= 27:
                        columns_to_keep = [0] + list(range(26, len(df.columns)))
                        df = df.iloc[:, columns_to_keep]

                        # 设置表头
                        original_headers = list(range(0, len(df.columns)))
                        original_headers[0] = "姓名"
                        df.columns = original_headers

                        # 插入新列
                        df.insert(1, "员工ID", "")
                        df.insert(2, "部门", "")
                        df.insert(3, "班次", "")

                        # 创建姓名映射
                        name_mapping = {}
                        for _, row in schedule_df.iterrows():
                            name = row['姓名']
                            name_mapping[name] = {
                                '员工ID': row['员工ID'],
                                '部门': row['部门'],
                                '班次': row[month_column]
                            }

                        # 填充数据
                        matched_count = 0
                        for idx, row in df.iterrows():
                            name = row['姓名']
                            if pd.notna(name) and name in name_mapping:
                                df.at[idx, '员工ID'] = name_mapping[name]['员工ID']
                                df.at[idx, '部门'] = name_mapping[name]['部门']
                                df.at[idx, '班次'] = name_mapping[name]['班次']
                                matched_count += 1

                        print(f"工作表 {sheet_name} 已匹配并填充 {matched_count} 条记录")

                        # 替换处理函数（确保非单元格匹配）
                        def replace_in_order(cell_value):
                            if pd.isna(cell_value):
                                return cell_value

                            # 强制转换为字符串
                            cell_str = str(cell_value)

                            # 逐个模式进行替换（仅替换匹配的部分）
                            for pattern in patterns_to_replace:
                                # 全局替换，只移除匹配的部分，保留其他内容
                                cell_str = re.sub(pattern, '', cell_str)

                            # 处理替换后可能产生的空白
                            cleaned_str = cell_str.strip()
                            return cleaned_str if cleaned_str else cell_value

                        # 应用替换到相关列
                        for col in df.columns:
                            if col not in ['姓名', '员工ID', '部门', '班次']:
                                # 先转换为字符串再处理，确保所有类型都能被正确匹配
                                df[col] = df[col].apply(lambda x: replace_in_order(str(x) if x is not None else ''))

                        # 保存处理后的工作表
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        print(f"已处理工作表: {sheet_name}")
                    else:
                        print(f"工作表 {sheet_name} 列数不足27列，已跳过")
                else:
                    print(f"工作表 {sheet_name} 行数不足，已跳过")

        print(f"文件处理完成，已保存至: {output_file}")
        return output_file

    except Exception as e:
        print(f"处理文件时出错: {str(e)}")
        return None


if __name__ == "__main__":
    print("用法: python 0.py <输入月报文件1路径> <输入班次文件2路径> <月份列名>")
    import sys
    if len(sys.argv) >= 4:
        input_file_path = sys.argv[1]
        schedule_file_path = sys.argv[2]
        month_column = sys.argv[3]
        output_file_path = sys.argv[4] if len(sys.argv) > 3 else None
        process_excel(input_file_path, schedule_file_path,output_file=output_file_path, month_column=month_column)
        output_file_path = sys.argv[4] if len(sys.argv) > 4 else None
        process_excel(
            input_file_path, 
            schedule_file_path, 
            output_file=output_file_path,  # 传递输出路径
            month_column=month_column
        )
    else:
        print("用法: python 0.py <输入文件> <班次文件> <月份列名> [输出文件]")