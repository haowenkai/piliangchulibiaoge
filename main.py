from tkinter import Tk
from app.app import ExcelRenamerApp
from app.core import ExcelManager

# 主程序入口
if __name__ == "__main__":
    root = Tk()  # 创建主窗口
    excel_manager = ExcelManager() # 初始化 ExcelManager
    app = ExcelRenamerApp(root, excel_manager)  # 实例化应用程序，传递 excel_manager
    root.mainloop()  # 运行主循环
