import tkinter as tk

def create_widgets(app):
    """创建 UI 组件"""
    # 显示当前打开的文件名称
    app.file_label = tk.Label(app.root, text="当前文件: 未打开", bg="#f0f0f0", font=("Arial", 12))
    app.file_label.pack(side=tk.TOP, pady=(10, 0))  # 上方显示
    
    # 显示当前文件的修改时间
    app.modified_time_label = tk.Label(app.root, text="修改时间: 未知", bg="#f0f0f0", font=("Arial", 10))
    app.modified_time_label.pack(side=tk.TOP, pady=(5, 10))  # 当前文件下方显示
    
    # 显示表格名称
    app.sheet_listbox = tk.Listbox(app.root, selectmode=tk.MULTIPLE)
    app.sheet_listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # 创建一个框架来放置按钮
    button_frame = tk.Frame(app.root, bg="#f0f0f0")
    button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(5, 10))  # 固定在底部
    
    # 复选框和按钮
    tk.Button(button_frame, text="修改名称", command=app.rename_sheet).pack(side=tk.LEFT, padx=5, pady=5)
    tk.Button(button_frame, text="删除所选行", command=app.delete_selected).pack(side=tk.LEFT, padx=5, pady=5)
    tk.Button(button_frame, text="退出程序", command=app.root.quit).pack(side=tk.LEFT, padx=5, pady=5)
    tk.Button(button_frame, text="打开文件", command=app.open_file).pack(side=tk.LEFT, padx=5, pady=5)

    # 显示当前时间（左下角）
    app.time_label = tk.Label(app.root, text="", bg="#f0f0f0", font=("Arial", 10))
    app.time_label.pack(side=tk.BOTTOM, anchor='se', padx=10, pady=(5, 10))  # 右下角
