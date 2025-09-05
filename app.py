import streamlit as st
import os
import subprocess
import uuid
import time
from io import BytesIO
import json

# 确保临时目录存在
os.makedirs('temp_files', exist_ok=True)
TEMP_DIR = 'temp_files'
processed_files = {}  # 存储处理后的文件映射


def clean_temp_files(max_age=3600):
    """清理过期临时文件（默认1小时）"""
    now = time.time()
    for filename in os.listdir(TEMP_DIR):
        file_path = os.path.join(TEMP_DIR, filename)
        if os.path.isfile(file_path) and now - os.path.getmtime(file_path) > max_age:
            try:
                os.remove(file_path)
            except:
                pass


def process_file(uploaded_file1, uploaded_file2, month_column):
    """处理上传的文件，按顺序执行两个处理脚本"""
    try:
        # 保存第一个上传的文件（使用唯一文件名避免覆盖）
        original_path = os.path.join(TEMP_DIR, f"月报_{uuid.uuid4()}.xlsx")
        with open(original_path, "wb") as f:
            f.write(uploaded_file1.getbuffer())
        
        # 保存第二个上传的文件（班次文件）
        schedule_path = os.path.join(TEMP_DIR, f"班次_{uuid.uuid4()}.xlsx")
        with open(schedule_path, "wb") as f:
            f.write(uploaded_file2.getbuffer())

        # 定义要执行的处理脚本
        scripts = [
            "0.py",  # 第一个处理脚本
            "1.py"  # 第二个处理脚本
        ]

        # 执行第一个脚本（修复参数顺序，与0.py的参数要求匹配）
        intermediate_path = os.path.join(TEMP_DIR, f"处理月报_{uuid.uuid4()}.xlsx")
        result1 = subprocess.run(
            ["python", os.path.join("modules", scripts[0]), 
             original_path, schedule_path, month_column, intermediate_path],  # 调整参数顺序
            capture_output=True,
            text=True
        )
        if result1.returncode != 0:
            return {"status": "error", "error": 
                    f"0.py执行失败（返回码：{result1.returncode}）\n"
                    f"标准输出：{result1.stdout}\n"
                    f"错误输出：{result1.stderr}"}
        
        # 检查中间文件是否生成
        if not os.path.exists(intermediate_path):
            return {"status": "error", "error": f"0.py未生成中间文件: {intermediate_path}"}

        # 执行第二个脚本
        final_path = os.path.join(TEMP_DIR, f"{month_column}原始数据.xlsx")
        result2 = subprocess.run(
            ["python", os.path.join("modules", scripts[1]), intermediate_path, final_path],
            capture_output=True,
            text=True
        )
        if result2.returncode != 0:
            return {"status": "error", "error": 
                    f"1.py执行失败（返回码：{result2.returncode}）\n"
                    f"标准输出：{result2.stdout}\n"
                    f"错误输出：{result2.stderr}"}

        # 检查最终文件是否生成
        if not os.path.exists(final_path):
            raise FileNotFoundError(f"1.py未生成最终文件: {final_path}")

        # 生成文件ID并存储路径
        file_id = str(uuid.uuid4())
        processed_files[file_id] = final_path

        print(f"生成的文件路径: {final_path}")
        print(f"文件是否存在: {os.path.exists(final_path)}")
        return {"status": "success", "file_id": file_id}
       
    except Exception as e:
        return {"status": "error", "error": str(e)}


def get_processed_file(file_id):
    """获取处理后的文件数据"""
    if file_id not in processed_files:
        return None
    file_path = processed_files[file_id]
    if not os.path.exists(file_path):
        return None

    with open(file_path, "rb") as f:
        return BytesIO(f.read())


def main():
    st.set_page_config(
        page_title="文件预处理工具",
        layout="wide",
        initial_sidebar_state="collapsed"
    )

    # 初始化session_state
    if "uploaded_file1" not in st.session_state:
        st.session_state["uploaded_file1"] = None
    if "uploaded_file2" not in st.session_state:
        st.session_state["uploaded_file2"] = None
    if "processing" not in st.session_state:
        st.session_state["processing"] = False
    if "processed_file_id" not in st.session_state:
        st.session_state["processed_file_id"] = None
    if "process_result" not in st.session_state:
        st.session_state["process_result"] = None

    # 页面标题
    st.title("文件预处理工具")
    st.write("上传文件后将自动按顺序执行预处理步骤，完成后可下载结果文件")

    st.subheader("选择班次月份")
    month = st.selectbox(
        "请选择需要匹配的班次月份",
        ["1月", "2月", "3月", "4月", "5月", "6月", 
        "7月", "8月", "9月", "10月", "11月", "12月"]
    )
    selected_month = f"{month}班次"  # 生成"8月班次"这样的列名格式

    # 第一个文件上传区域
    uploaded_file1 = st.file_uploader(
        "选择第一个需要处理的文件",
        type=["xlsx", "xls", "csv"],
        key="file_uploader1"
    )

    # 第二个文件上传区域
    uploaded_file2 = st.file_uploader(
        "选择第二个需要处理的文件（班次文件）",
        type=["xlsx", "xls", "csv"],
        key="file_uploader2"
    )

    # 检查两个文件是否都已上传
    if uploaded_file1 and uploaded_file2:
        st.session_state["uploaded_file1"] = uploaded_file1
        st.session_state["uploaded_file2"] = uploaded_file2
        st.success(f"文件上传成功: {uploaded_file1.name} 和 {uploaded_file2.name}")

        # 处理按钮（只有两个文件都上传后才启用）
        if st.button(
               "开始处理文件",
                disabled=st.session_state["processing"] or (st.session_state["processed_file_id"] is not None)
            ):
            st.session_state["processing"] = True
            st.session_state["processed_file_id"] = None

            # 显示处理状态
            with st.spinner("正在处理文件，请稍候..."):
                # 传入两个文件进行处理
                result = process_file(uploaded_file1, uploaded_file2, selected_month)
                st.session_state["process_result"] = result
                st.session_state["processing"] = False

                if result["status"] == "success":
                    st.session_state["processed_file_id"] = result["file_id"]
                    st.success("文件处理完成！")

                    # 显示下载按钮
                    excel_data = get_processed_file(result["file_id"])
                    if excel_data:
                        st.download_button(
                            label="下载处理结果",
                            data=excel_data,
                            file_name=f"{selected_month}处理结果.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else:
                    st.error(f"处理失败: {result['error']}")

    # 已处理文件下载区（如果有）
    if st.session_state["processed_file_id"] and not st.session_state["processing"]:
        st.subheader("处理结果")
        excel_data = get_processed_file(st.session_state["processed_file_id"])
        if excel_data and st.session_state.get("uploaded_file1"):
            st.download_button(
                label="重新下载处理结果",
                data=excel_data,
                file_name=f"{selected_month}处理结果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="redownload_btn"
            )
        elif not excel_data:
            st.warning("处理后的文件不存在或已过期")

    # 定期清理临时文件
    clean_temp_files()

    # 隐藏Streamlit默认元素（可选）
    # st.markdown("""
    #     <style>
    #         #MainMenu {display: none !important;}
    #         footer {display: none !important;}
    #         header {display: none !important;}
    #     </style>
    # """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()