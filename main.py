import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import openpyxl
import os
from datetime import datetime

class ExcelRenamerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("批量修改表格名称")
        
        self.file_path = None
        self.workbook = None
        self.sheet_names = []
        self.is_topmost = False  # 用于跟踪窗口是否置顶
        
        # UI Components
        self.create_widgets()
        
    def create_widgets(self):
        # 显示当前打开的文件名称
        self.file_label = tk.Label(self.root, text="当前文件: 未打开")
        self.file_label.pack(side=tk.RIGHT)
        
        # 显示当前文件的修改时间
        self.modified_time_label = tk.Label(self.root, text="修改时间: 未知")
        self.modified_time_label.pack(side=tk.TOP)
        
        # 显示当前时间（左下角）
        self.time_label = tk.Label(self.root, text=self.get_current_time())
        self.time_label.pack(side=tk.LEFT, anchor='sw')
        
        # 显示表格名称
        self.sheet_listbox = tk.Listbox(self.root, selectmode=tk.MULTIPLE)
        self.sheet_listbox.pack(fill=tk.BOTH, expand=True)
        
        # 复选框和按钮
        self.rename_button = tk.Button(self.root, text="修改名称", command=self.rename_sheet)
        self.rename_button.pack(side=tk.LEFT)
        
        self.delete_button = tk.Button(self.root, text="删除所选行", command=self.delete_selected)
        self.delete_button.pack(side=tk.LEFT)
        
        self.exit_button = tk.Button(self.root, text="退出程序", command=self.root.quit)
        self.exit_button.pack(side=tk.LEFT)
        
        # 置顶按钮
        self.topmost_button = tk.Button(self.root, text="置顶", command=self.toggle_topmost)
        self.topmost_button.pack(side=tk.TOP, anchor='ne')  # 右上角
        
        # 打开文件按钮
        self.open_button = tk.Button(self.root, text="打开文件", command=self.open_file)
        self.open_button.pack(side=tk.BOTTOM)

    def toggle_topmost(self):
        self.is_topmost = not self.is_topmost
        self.root.attributes('-topmost', self.is_topmost)
        self.topmost_button.config(text="取消置顶" if self.is_topmost else "置顶")

    def get_current_time(self):
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def open_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if self.file_path:
            self.workbook = openpyxl.load_workbook(self.file_path)
            self.sheet_names = self.workbook.sheetnames
            self.update_ui()

    def update_ui(self):
        self.file_label.config(text=f"当前文件: {os.path.basename(self.file_path)}")
        self.modified_time_label.config(text=f"修改时间: {datetime.fromtimestamp(os.path.getmtime(self.file_path)).strftime('%Y-%m-%d %H:%M:%S')}")
        self.sheet_listbox.delete(0, tk.END)
        for sheet in self.sheet_names:
            self.sheet_listbox.insert(tk.END, sheet)

    def rename_sheet(self):
        selected_indices = self.sheet_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("警告", "请先选择要修改的表格名称")
            return
        
        new_name = simpledialog.askstring("输入新名称", "请输入新名称:")
        if new_name:
            for index in selected_indices:
                old_name = self.sheet_names[index]
                # 检查新名称是否已存在
                if new_name in self.sheet_names and new_name != old_name:
                    messagebox.showwarning("警告", f"工作表名称 '{new_name}' 已存在，请使用其他名称。")
                    return
                # 修改工作表名称
                self.workbook[old_name].title = new_name
                # 更新 sheet_names 列表
                self.sheet_names[index] = new_name
            self.workbook.save(self.file_path)
            self.update_ui()

    def delete_selected(self):
        selected_indices = self.sheet_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("警告", "请先选择要删除的表格名称")
            return
        
        for index in reversed(selected_indices):
            del self.sheet_names[index]
            self.sheet_listbox.delete(index)
        self.workbook.save(self.file_path)
        self.update_ui()  # 更新 UI，清空表格名称

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelRenamerApp(root)
    root.mainloop()