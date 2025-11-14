import tkinter as tk
from tkinter import filedialog, messagebox
from . import _run_conversion


class LessonPlanGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("教案生成器")
        self.root.geometry("600x250")

        # Template file path
        self.template_path = tk.StringVar()
        # Markdown files paths
        self.markdown_paths = []
        # Output directory
        self.output_dir = tk.StringVar()

        # UI setup
        self.setup_ui()

    def setup_ui(self):
        # Template selection
        tk.Label(self.root, text="模板:").grid(row=0, column=0, padx=10, pady=10, sticky='w')
        tk.Entry(self.root, textvariable=self.template_path, width=50).grid(row=0, column=1, padx=5, pady=10)
        tk.Button(self.root, text="选择", command=self.select_template).grid(row=0, column=2, padx=5, pady=10)

        # Markdown files selection
        tk.Label(self.root, text="markdown教案:").grid(row=1, column=0, padx=10, pady=10, sticky='w')
        self.markdown_entry = tk.Entry(self.root, width=50)
        self.markdown_entry.grid(row=1, column=1, padx=5, pady=10)
        tk.Button(self.root, text="选择", command=self.select_markdown_files).grid(row=1, column=2, padx=5, pady=10)

        # Output directory selection
        tk.Label(self.root, text="输出目录:").grid(row=2, column=0, padx=10, pady=10, sticky='w')
        tk.Entry(self.root, textvariable=self.output_dir, width=50).grid(row=2, column=1, padx=5, pady=10)
        tk.Button(self.root, text="选择", command=self.select_output_dir).grid(row=2, column=2, padx=5, pady=10)

        # Convert button
        tk.Button(self.root, text="开始", command=self.convert_files).grid(row=3, column=1, pady=20)

    def select_template(self):
        file_path = filedialog.askopenfilename(
            title="选择模板文件",
            filetypes=[("Word 文档", "*.docx")]
        )
        if file_path:
            self.template_path.set(file_path)

    def select_markdown_files(self):
        files = filedialog.askopenfilenames(
            title="选择 Markdown 教案文件",
            filetypes=[("Markdown 文件", "*.md")]
        )
        if files:
            self.markdown_paths = list(files)
            self.markdown_entry.delete(0, tk.END)
            self.markdown_entry.insert(0, f"{len(files)} 个文件已选择")

    def select_output_dir(self):
        directory = filedialog.askdirectory(title="选择输出目录")
        if directory:
            self.output_dir.set(directory)

    def convert_files(self):
        template = self.template_path.get()
        if not template:
            messagebox.showerror("错误", "请选择模板文件")
            return
        if not self.markdown_paths:
            messagebox.showerror("错误", "请选择至少一个 Markdown 教案文件")
            return
        output_dir = self.output_dir.get()
        if not output_dir:
            messagebox.showerror("错误", "请选择输出目录")
            return

        try:
            for md_file in self.markdown_paths:
                _run_conversion(template, md_file, output_dir)
            messagebox.showinfo("完成", "已经生成所有docx教案文件")
        except Exception as e:
            messagebox.showerror("错误", f"转换过程中出错:\n{str(e)}")


# def main():
#     root = tk.Tk()
#     app = LessonPlanGUI(root)
#     root.mainloop()


# if __name__ == "__main__":
#     main()