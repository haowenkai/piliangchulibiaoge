import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from app.ui import create_widgets
from app.file_operations import load_workbook, save_workbook
from datetime import datetime
import os

class ExcelRenamerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("批量修改表格名称")
        
        self.file_path = None  # 当前文件路径
        self.workbook = None  # 当前工作簿
        self.sheet_names = []  # 工作表名称列表
        self.is_topmost = False  # 用于跟踪窗口是否置顶
        
        # 设置背景颜色
        self.root.configure(bg="#f0f0f0")
        
        # 创建 UI 组件
        create_widgets(self)

        # 显示当前时间
        self.update_time_labels()

    def toggle_topmost(self):
        """切换窗口置顶状态"""
        self.is_topmost = not self.is_topmost
        self.root.attributes('-topmost', self.is_topmost)

    def open_file(self):
        """打开 Excel 文件并加载工作表"""
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if self.file_path:
            self.workbook, self.sheet_names = load_workbook(self.file_path)
            self.update_ui()

    def update_ui(self):
        """更新 UI 显示内容"""
        self.file_label.config(text=f"当前文件: {os.path.basename(self.file_path)}")
        self.modified_time_label.config(text=f"修改时间: {datetime.fromtimestamp(os.path.getmtime(self.file_path)).strftime('%Y-%m-%d %H:%M:%S')}")
        self.sheet_listbox.delete(0, tk.END)
        for sheet in self.sheet_names:
            self.sheet_listbox.insert(tk.END, sheet)

    def update_time_labels(self):
        """更新当前时间标签"""
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.time_label.config(text=current_time)
        self.root.after(1000, self.update_time_labels)  # 每秒更新一次

    def rename_sheet(self):
        """重命名选中的工作表"""
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
        """删除选中的工作表"""
        selected_indices = self.sheet_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("警告", "请先选择要删除的表格名称")
            return
        
        for index in reversed(selected_indices):
            del self.sheet_names[index]
            self.sheet_listbox.delete(index)
        
        # 清空当前文件名和修改时间
        self.file_label.config(text="当前文件: 未打开")
        self.modified_time_label.config(text="修改时间: 未知")
        
        self.update_ui()  # 更新 UI，清空表格名称

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelRenamerApp(root)
    root.mainloop() 