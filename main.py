from tkinter import Tk
from app.app import ExcelRenamerApp

# 主程序入口
if __name__ == "__main__":
    root = Tk()  # 创建主窗口
    app = ExcelRenamerApp(root)  # 实例化应用程序
    root.mainloop()  # 运行主循环