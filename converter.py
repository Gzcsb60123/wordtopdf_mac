import subprocess
import os
from config import LIBREOFFICE_PATH, DEFAULT_INPUT_FOLDER, DEFAULT_OUTPUT_FOLDER

def convert_word_to_pdf(input_folder=DEFAULT_INPUT_FOLDER, output_folder=DEFAULT_OUTPUT_FOLDER):
    # 扩展输入和输出路径中的 ~
    full_input_path = os.path.expanduser(input_folder)
    full_output_path = os.path.expanduser(output_folder)
    
    # 自动创建输入目录（如果不存在）
    if not os.path.exists(full_input_path):
        os.makedirs(full_input_path, exist_ok=True)
        log_messages = [f"ℹ️ 已创建输入文件夹: {full_input_path}"]
    else:
        log_messages = []
    
    # 自动创建输出目录（如果不存在）
    if not os.path.exists(full_output_path):
        os.makedirs(full_output_path, exist_ok=True)
        log_messages.append(f"ℹ️ 已创建输出文件夹: {full_output_path}")

    # 如果输入目录为空，提示用户放入文件
    if not os.listdir(full_input_path):
        log_messages.append(f"⚠️ 输入文件夹 {full_input_path} 为空，请放入 .doc 或 .docx 文件")
        return log_messages

    # 遍历输入文件夹中的文件
    for filename in os.listdir(full_input_path):
        if filename.lower().endswith((".doc", ".docx")) and not filename.startswith('~$'):
            input_path = os.path.join(full_input_path, filename)
            base_name = os.path.splitext(filename)[0]
            output_path = os.path.join(full_output_path, f"{base_name}.pdf")

            cmd = [
                LIBREOFFICE_PATH,
                "--headless",
                "--invisible",
                "--norestore",
                "--nologo",
                "--convert-to", "pdf",
                input_path,
                "--outdir", full_output_path
            ]

            # 设置环境变量以禁用 GUI
            env = os.environ.copy()
            env["DISPLAY"] = ""

            try:
                subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, env=env)
                log_messages.append(f"✅ 成功转换: {filename}")
            except subprocess.CalledProcessError as e:
                log_messages.append(f"❌ 转换失败: {filename} | 错误: {e.stderr.decode('utf-8')}")
            except Exception as e:
                log_messages.append(f"⚠️ 异常: {filename} | {str(e)}")

    # 清理残留的 LibreOffice 进程
    subprocess.run(["pkill", "-f", "soffice"], stderr=subprocess.DEVNULL)
    
    return log_messages