import tkinter as tk
from tkinter import simpledialog, messagebox, filedialog
import os
from datetime import datetime

class UI:
    def __init__(self, root, excel_manager):
        self.root = root
        self.excel_manager = excel_manager
        self.file_label = tk.Label(root, text="当前文件: 未打开", bg="#f0f0f0", font=("Arial", 12))
        self.file_label.pack(side=tk.TOP, pady=(10, 0))
        self.modified_time_label = tk.Label(root, text="修改时间: 未知", bg="#f0f0f0", font=("Arial", 10))
        self.modified_time_label.pack(side=tk.TOP, pady=(5, 10))
        self.sheet_listbox = tk.Listbox(root, selectmode=tk.SINGLE)
        self.sheet_listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        button_frame = tk.Frame(root, bg="#f0f0f0")
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(5, 10))

        self.rename_button = tk.Button(button_frame, text="修改名称", command=self.rename_selected_sheet)
        self.rename_button.pack(side=tk.LEFT, padx=5, pady=5)
        self.delete_button = tk.Button(button_frame, text="删除所选行", command=self.delete_selected_sheet)
        self.delete_button.pack(side=tk.LEFT, padx=5, pady=5)

        # 读取数据按钮
        self.read_data_button = tk.Button(button_frame, text="读取数据", command=self.read_selected_data)
        self.read_data_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.exit_button = tk.Button(button_frame, text="退出程序", command=root.quit)
        self.exit_button.pack(side=tk.LEFT, padx=5, pady=5)
        self.open_file_button = tk.Button(button_frame, text="打开文件", command=self.open_file)
        self.open_file_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.time_label = tk.Label(root, text="", bg="#f0f0f0", font=("Arial", 10))
        self.time_label.pack(side=tk.BOTTOM, anchor='se', padx=10, pady=(5, 10))

        root.bind("<Delete>", lambda event: self.delete_selected_sheet())

    def open_file(self):
        filepath = filedialog.askopenfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx;*.xlsm;*.xltx;*.xltm")]
        )
        if filepath:
            try:
                self.excel_manager.load_workbook(filepath)
                filename = os.path.basename(filepath)
                modified_time = datetime.fromtimestamp(os.path.getmtime(filepath)).strftime("%Y-%m-%d %H:%M:%S")
                self.update_file_label(filename)
                self.update_modified_time_label(modified_time)
                self.populate_sheet_list(self.excel_manager.workbook.sheetnames)
            except Exception as e:
                messagebox.showerror("错误", f"打开文件错误: {e}")

    def read_selected_data(self):
        selected_sheet = self.sheet_listbox.get(tk.ANCHOR)
        if selected_sheet:
            data = self.excel_manager.read_sheet_data(selected_sheet)
            # 显示数据在一个新的窗口
            data_window = tk.Toplevel(self.root)
            data_window.title(f"{selected_sheet} 数据")
            text_area = tk.Text(data_window, wrap=tk.NONE)
            text_area.pack(expand=True, fill='both')
            for row in data:
                text_area.insert(tk.END, "\t".join(map(str, row)) + "\n")
            text_area.config(state=tk.DISABLED)
        else:
            messagebox.showinfo("提示", "请选择要读取的工作表")

    def update_file_label(self, filename):
        self.file_label.config(text=f"当前文件: {filename}")

    def update_modified_time_label(self, modified_time):
        self.modified_time_label.config(text=f"修改时间: {modified_time}")

    def populate_sheet_list(self, sheet_names):
        self.sheet_listbox.delete(0, tk.END)
        for name in sheet_names:
            self.sheet_listbox.insert(tk.END, name)

    def rename_selected_sheet(self):
        selected_sheet = self.sheet_listbox.get(tk.ANCHOR)
        if not selected_sheet:
            messagebox.showinfo("提示", "请选择要重命名的工作表")
            return

        new_name = simpledialog.askstring("重命名", f"将 '{selected_sheet}' 重命名为:")
        if new_name:
            try:
                self.excel_manager.rename_sheet(selected_sheet, new_name)
                self.populate_sheet_list(self.excel_manager.workbook.sheetnames)
                messagebox.showinfo("成功", f"工作表 '{selected_sheet}' 已重命名为 '{new_name}'")
            except Exception as e:
                messagebox.showerror("错误", f"重命名工作表错误: {e}")

    def delete_selected_sheet(self):
        selected_sheet = self.sheet_listbox.get(tk.ANCHOR)
        if selected_sheet:
            if messagebox.askyesno("确认", f"确定要删除工作表 '{selected_sheet}' 吗?"):
                try:
                    self.excel_manager.delete_sheet(selected_sheet)
                    self.populate_sheet_list(self.excel_manager.workbook.sheetnames)
                    messagebox.showinfo("成功", f"工作表 '{selected_sheet}' 已删除")
                except Exception as e:
                    messagebox.showerror("错误", f"删除工作表错误: {e}")
