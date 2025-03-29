import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from converter import convert_word_to_pdf
from config import DEFAULT_INPUT_FOLDER, DEFAULT_OUTPUT_FOLDER

class WordToPDFApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Word 转 PDF 工具")
        self.root.geometry("700x400")  # 放宽窗口宽度，从 600 到 700
        self.root.minsize(700, 400)    # 设置最小宽度，防止缩小时遮挡
        
        # 输入输出路径
        self.input_folder = tk.StringVar(value=DEFAULT_INPUT_FOLDER)
        self.output_folder = tk.StringVar(value=DEFAULT_OUTPUT_FOLDER)
        
        # 创建控件
        self.create_widgets()
    
    def create_widgets(self):
        # 输入路径选择
        ttk.Label(self.root, text="输入文件夹:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        ttk.Entry(self.root, textvariable=self.input_folder, width=50).grid(row=0, column=1, padx=10, pady=5)
        ttk.Button(self.root, text="浏览", command=self.select_input_folder).grid(row=0, column=2, padx=10, pady=5)
        
        # 输出路径选择
        ttk.Label(self.root, text="输出文件夹:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        ttk.Entry(self.root, textvariable=self.output_folder, width=50).grid(row=1, column=1, padx=10, pady=5)
        ttk.Button(self.root, text="浏览", command=self.select_output_folder).grid(row=1, column=2, padx=10, pady=5)
        
        # 转换按钮
        ttk.Button(self.root, text="开始转换", command=self.start_conversion).grid(row=2, column=1, pady=20)
        
        # 日志区域
        self.log_text = tk.Text(self.root, height=15, width=80)  # 增加宽度以匹配窗口
        self.log_text.grid(row=3, column=0, columnspan=3, padx=10, pady=5)
    
    def select_input_folder(self):
        folder = filedialog.askdirectory(initialdir=self.input_folder.get())
        if folder:
            self.input_folder.set(folder)
    
    def select_output_folder(self):
        folder = filedialog.askdirectory(initialdir=self.output_folder.get())
        if folder:
            self.output_folder.set(folder)
    
    def start_conversion(self):
        self.log_text.delete(1.0, tk.END)
        logs = convert_word_to_pdf(self.input_folder.get(), self.output_folder.get())
        for log in logs:
            self.log_text.insert(tk.END, log + "\n")
        messagebox.showinfo("完成", "转换任务已完成！")

if __name__ == "__main__":
    root = tk.Tk()
    app = WordToPDFApp(root)
    root.mainloop()