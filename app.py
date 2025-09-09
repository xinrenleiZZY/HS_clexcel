import streamlit as st
import os
import subprocess
import uuid
import time
from io import BytesIO
import sys

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


def process_file(uploaded_file1, uploaded_file2):
    """处理上传的文件，按顺序执行两个处理脚本"""
    try:
        # 保存第一个上传的文件（使用唯一文件名避免覆盖）
        original_path = os.path.join(TEMP_DIR, f"月报_xin01_1.xlsx")
        with open(original_path, "wb") as f:
            f.write(uploaded_file1.getbuffer())
        
        # 保存第二个上传的文件（员工信息文件）
        employee_info_path = os.path.join(TEMP_DIR, f"员工信息_xin01_2.xlsx")
        with open(employee_info_path, "wb") as f:
            f.write(uploaded_file2.getbuffer())

        # 定义脚本及对应的参数，已更新为新脚本名称
        scripts = [
            {
                "path": "新01.py",
                "args": [original_path, employee_info_path]  # 新01.py需要的参数
            },
            {
                "path": "新02.py",
                "args": []  # 新02.py的参数将在第一个脚本执行后生成
            }
        ]

        # 执行第一个脚本（新01.py）
        script0_path = os.path.join("modules", scripts[0]["path"])
        intermediate_path = os.path.join(TEMP_DIR, f"处理月报_xin01_3.xlsx")
        # 完整参数：[Python解释器, 脚本路径, 输入文件, 员工信息文件, 中间输出文件]
        script0_args = [sys.executable, script0_path] + scripts[0]["args"] + ["",intermediate_path]
        
        result1 = subprocess.run(
            script0_args,
            capture_output=True,
            text=True,
            check=True  # 启用检查，返回非0状态码时抛出异常
        )

        # 检查中间文件是否生成
        if not os.path.exists(intermediate_path):
            return {"status": "error", "error": f"新01.py未生成中间文件: {intermediate_path}"}

        # 执行第二个脚本（新02.py）
        script1_path = os.path.join("modules", scripts[1]["path"])
        final_path = os.path.join(TEMP_DIR, f"原始数据.xlsx")
        # 完整参数：[Python解释器, 脚本路径, 中间文件, 最终输出文件]
        script1_args = [sys.executable, script1_path] + [intermediate_path] + [final_path]
        
        result2 = subprocess.run(
            script1_args,
            capture_output=True,
            text=True,
            check=True  # 启用检查，返回非0状态码时抛出异常
        )

        # 检查最终文件是否生成
        if not os.path.exists(final_path):
            raise FileNotFoundError(f"新02.py未生成最终文件: {final_path}")
        
        print(f"生成的文件路径: {final_path}")
        print(f"文件是否存在: {os.path.exists(final_path)}")
        # 存储结果并返回
        file_id = str(uuid.uuid4())
        processed_files[file_id] = final_path
        return {"status": "success", "file_id": file_id}
       


    except subprocess.CalledProcessError as e:
        # 捕获脚本执行失败的异常（返回码非0）
        return {"status": "error", "error": 
                f"脚本执行失败（返回码：{e.returncode}）\n"
                f"命令：{' '.join(e.cmd)}\n"
                f"错误输出：{e.stderr}"}
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

    # 第一个文件上传区域
    uploaded_file1 = st.file_uploader(
        "选择需要处理的月报文件",
        type=["xlsx", "xls", "csv"],
        key="file_uploader1"
    )

    # 第二个文件上传区域
    uploaded_file2 = st.file_uploader(
        "选择员工信息文件（包含姓名、员工ID、部门）",
        type=["xlsx", "xls", "csv"],
        key="file_uploader2"
    )

    # 检查两个文件是否都已上传
    if uploaded_file1 and uploaded_file2:
        st.session_state["uploaded_file1"] = uploaded_file1
        st.session_state["uploaded_file2"] = uploaded_file2
        st.success(f"文件上传成功: {uploaded_file1.name} 和 {uploaded_file2.name}")

        # 处理按钮
        if st.button(
               "开始处理文件",
                disabled=st.session_state["processing"] or (st.session_state["processed_file_id"] is not None)
            ):
            st.session_state["processing"] = True
            st.session_state["processed_file_id"] = None

            # 显示处理状态
            with st.spinner("正在处理文件，请稍候..."):
                # 传入两个文件进行处理
                result = process_file(uploaded_file1, uploaded_file2)
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
                            file_name=f"处理结果.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else:
                    st.error(f"处理失败: {result['error']}")

    # 已处理文件下载区
    if st.session_state["processed_file_id"] and not st.session_state["processing"]:
        st.subheader("处理结果")
        excel_data = get_processed_file(st.session_state["processed_file_id"])
        if excel_data and st.session_state.get("uploaded_file1"):
            st.download_button(
                label="重新下载处理结果",
                data=excel_data,
                file_name=f"处理结果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="redownload_btn"
            )
        elif not excel_data:
            st.warning("处理后的文件不存在或已过期")

    # 定期清理临时文件
    clean_temp_files()


if __name__ == "__main__":
    main()
    