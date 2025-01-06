from openpyxl import Workbook
from pathlib import Path
from datetime import datetime
import os
from typing import Tuple, List, Optional
from . import logger
import openpyxl

class ExcelManager:
    def __init__(self):
        self.file_path: Optional[Path] = None
        self.workbook: Optional[Workbook] = None
        self.sheet_names: List[str] = []

    def load_workbook(self, file_path: str) -> Tuple[Workbook, List[str]]:
        """加载Excel文件"""
        try:
            self.file_path = Path(file_path)
            self.workbook = openpyxl.load_workbook(file_path)
            self.sheet_names = self.workbook.sheetnames
            logger.info(f"成功加载文件: {self.file_path}")
            return self.workbook, self.sheet_names
        except Exception as e:
            logger.error(f"加载文件失败: {str(e)}")
            raise

    def save_workbook(self) -> None:
        """保存Excel文件"""
        if self.workbook and self.file_path:
            try:
                self.workbook.save(self.file_path)
                logger.info(f"成功保存文件: {self.file_path}")
            except Exception as e:
                logger.error(f"保存文件失败: {str(e)}")
                raise

    def rename_sheet(self, old_name: str, new_name: str) -> None:
        """重命名工作表"""
        if self.workbook:
            try:
                if new_name in self.sheet_names and new_name != old_name:
                    raise ValueError(f"工作表名称 '{new_name}' 已存在")

                self.workbook[old_name].title = new_name
                index = self.sheet_names.index(old_name)
                self.sheet_names[index] = new_name
                logger.info(f"成功重命名工作表: {old_name} -> {new_name}")
            except Exception as e:
                logger.error(f"重命名工作表失败: {str(e)}")
                raise

    def delete_sheet(self, sheet_name: str) -> None:
        """删除工作表"""
        if self.workbook:
            try:
                del self.workbook[sheet_name]
                self.sheet_names.remove(sheet_name)
                logger.info(f"成功删除工作表: {sheet_name}")
            except Exception as e:
                logger.error(f"删除工作表失败: {str(e)}")
                raise

    def get_file_info(self) -> Tuple[str, str]:
        """获取文件信息"""
        if self.file_path and self.file_path.exists():
            filename = self.file_path.name
            modified_time = datetime.fromtimestamp(
                self.file_path.stat().st_mtime
            ).strftime('%Y-%m-%d %H:%M:%S')
            return filename, modified_time
        return "未打开", "未知"
