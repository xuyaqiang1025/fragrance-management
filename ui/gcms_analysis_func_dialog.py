from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView, QMessageBox)
from PyQt6.QtCore import Qt
from sqlalchemy.orm import Session
from models import GCMSAnalysis, GCMSCompound
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import seaborn as sns

class GCMSAnalysisFuncDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("GC-MS 分析功能")
        self.setMinimumSize(800, 600)
        self.setup_ui()
        self.load_data()

    def setup_ui(self):
        layout = QVBoxLayout()

        # 添加样品列表表格
        self.sample_table = QTableWidget()
        self.sample_table.setColumnCount(4)
        self.sample_table.setHorizontalHeaderLabels(["编号", "名称", "供应商信息", "仪器参数"])
        self.sample_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.sample_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.sample_table.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        layout.addWidget(self.sample_table)

        # 添加分析按钮
        analyze_layout = QHBoxLayout()
        self.analyze_btn = QPushButton("分析选定样品")
        self.analyze_btn.clicked.connect(self.analyze_selected_samples)
        analyze_layout.addWidget(self.analyze_btn)
        layout.addLayout(analyze_layout)

        self.setLayout(layout)

    def load_data(self):
        with Session(self.parent().engine) as session:
            analyses = session.query(GCMSAnalysis).all()
            self.sample_table.setRowCount(len(analyses))
            for i, analysis in enumerate(analyses):
                self.sample_table.setItem(i, 0, QTableWidgetItem(analysis.number))
                self.sample_table.setItem(i, 1, QTableWidgetItem(analysis.name))
                self.sample_table.setItem(i, 2, QTableWidgetItem(analysis.supplier_info))
                self.sample_table.setItem(i, 3, QTableWidgetItem(analysis.instrument_params))

    def get_selected_sample_ids(self):
        selected_rows = self.sample_table.selectedItems()
        if not selected_rows:
            QMessageBox.warning(self, "警告", "请至少选择一个样品")
            return []
        return [self.sample_table.item(row.row(), 0).text() for row in selected_rows]

    def analyze_selected_samples(self):
        sample_ids = self.get_selected_sample_ids()
        if not sample_ids:
            return
        self.analyze_characteristic_compounds(sample_ids)
        self.analyze_difference_compounds(sample_ids)

    def analyze_characteristic_compounds(self, sample_ids):
        with Session(self.parent().engine) as session:
            compounds = session.query(GCMSCompound).filter(GCMSCompound.analysis_id.in_(sample_ids)).all()
            if not compounds:
                QMessageBox.warning(self, "警告", "未找到相关化合物数据")
                return
            # 处理化合物数据，生成特征物质报告
            # ... 省略具体分析逻辑 ...

    def analyze_difference_compounds(self, sample_ids):
        with Session(self.parent().engine) as session:
            compounds = session.query(GCMSCompound).filter(GCMSCompound.analysis_id.in_(sample_ids)).all()
            if not compounds:
                QMessageBox.warning(self, "警告", "未找到相关化合物数据")
                return
            # 处理化合物数据，生成差异物质报告
            # ... 省略具体分析逻辑 ... 