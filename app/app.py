import os
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from app.ui import create_widgets
from app.core import ExcelManager
from datetime import datetime
from dotenv import load_dotenv
from . import logger

class ExcelRenamerApp:
    def __init__(self, root):
        # 加载环境变量
        load_dotenv()
        
        # 初始化UI
        self.root = root
        self.root.title(os.getenv("APP_TITLE", "批量修改表格名称"))
        self.root.geometry(f"{os.getenv('APP_WIDTH', '800')}x{os.getenv('APP_HEIGHT', '600')}")
        
        # 初始化业务逻辑
        self.excel_manager = ExcelManager()
        
        # 创建UI组件
        create_widgets(self)
        
        # 初始化状态
        self.is_topmost = False
        self.update_time_labels()

    def toggle_topmost(self):
        """切换窗口置顶状态"""
        self.is_topmost = not self.is_topmost
        self.root.attributes('-topmost', self.is_topmost)
        logger.info(f"窗口置顶状态: {self.is_topmost}")

    def open_file(self):
        """打开Excel文件"""
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        )
        if file_path:
            try:
                self.excel_manager.load_workbook(file_path)
                self.update_ui()
            except Exception as e:
                messagebox.showerror("错误", str(e))
                logger.error(f"打开文件失败: {str(e)}")

    def update_ui(self):
        """更新UI显示"""
        filename, modified_time = self.excel_manager.get_file_info()
        self.file_label.config(text=f"当前文件: {filename}")
        self.modified_time_label.config(text=f"修改时间: {modified_time}")
        
        # 更新工作表列表
        self.sheet_listbox.delete(0, tk.END)
        for sheet in self.excel_manager.sheet_names:
            self.sheet_listbox.insert(tk.END, sheet)

    def update_time_labels(self):
        """更新时间显示"""
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.time_label.config(text=current_time)
        self.root.after(1000, self.update_time_labels)

    def rename_sheet(self):
        """重命名工作表"""
        selected_indices = self.sheet_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("警告", "请先选择要修改的表格名称")
            return
        
        new_name = simpledialog.askstring("输入新名称", "请输入新名称:")
        if new_name:
            try:
                for index in selected_indices:
                    old_name = self.excel_manager.sheet_names[index]
                    self.excel_manager.rename_sheet(old_name, new_name)
                self.excel_manager.save_workbook()
                self.update_ui()
            except Exception as e:
                messagebox.showerror("错误", str(e))
                logger.error(f"重命名失败: {str(e)}")

    def delete_selected(self):
        """删除选中的工作表"""
        selected_indices = self.sheet_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("警告", "请先选择要删除的表格名称")
            return
        
        try:
            for index in reversed(selected_indices):
                sheet_name = self.excel_manager.sheet_names[index]
                self.excel_manager.delete_sheet(sheet_name)
            self.excel_manager.save_workbook()
            self.update_ui()
        except Exception as e:
            messagebox.showerror("错误", str(e))
            logger.error(f"删除失败: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelRenamerApp(root)
    root.mainloop()
