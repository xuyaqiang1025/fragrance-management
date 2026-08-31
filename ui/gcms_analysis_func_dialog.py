from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                             QTableWidget, QTableWidgetItem, QHeaderView,
                             QAbstractItemView, QMessageBox, QTabWidget, QWidget)
from PyQt6.QtCore import Qt
from sqlalchemy.orm import Session
from models import GCMSAnalysis, GCMSCompound

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure


def _cname(c):
    """化合物显示名：中文优先，其次英文，否则标记未知"""
    return (c.name_cn or c.name_en or '').strip() or '未知化合物'


class GCMSAnalysisFuncDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("GC-MS 分析功能")
        self.setMinimumSize(920, 640)
        self.setup_ui()
        self.load_data()

    def setup_ui(self):
        layout = QVBoxLayout()

        # 样品列表
        self.sample_table = QTableWidget()
        self.sample_table.setColumnCount(4)
        self.sample_table.setHorizontalHeaderLabels(
            ["编号", "名称", "供应商", "仪器参数"])
        self.sample_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch)
        self.sample_table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows)
        self.sample_table.setSelectionMode(
            QAbstractItemView.SelectionMode.MultiSelection)
        layout.addWidget(self.sample_table, 1)

        # 分析按钮
        analyze_layout = QHBoxLayout()
        self.analyze_btn = QPushButton("分析选定样品")
        self.analyze_btn.clicked.connect(self.analyze_selected_samples)
        analyze_layout.addWidget(self.analyze_btn)
        analyze_layout.addStretch()
        layout.addLayout(analyze_layout)

        # 结果区（选项卡）
        self.result_tabs = QTabWidget()

        self.tab_char = QWidget()
        self.tab_char_layout = QVBoxLayout()
        self.char_table = QTableWidget()
        self.char_table.setColumnCount(6)
        self.char_table.setHorizontalHeaderLabels(
            ["样品编号", "样品名称", "排名", "化合物", "相对含量", "CAS"])
        self.char_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch)
        self.tab_char_layout.addWidget(self.char_table)
        self.tab_char.setLayout(self.tab_char_layout)
        self.result_tabs.addTab(self.tab_char, "特征物质")

        self.tab_diff = QWidget()
        self.tab_diff_layout = QVBoxLayout()
        self.diff_table = QTableWidget()
        self.diff_table.setColumnCount(4)
        self.diff_table.setHorizontalHeaderLabels(
            ["化合物", "出现样品数", "类型", "各样品相对含量"])
        self.diff_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch)
        self.tab_diff_layout.addWidget(self.diff_table)
        self.tab_diff.setLayout(self.tab_diff_layout)
        self.result_tabs.addTab(self.tab_diff, "差异物质")

        self.tab_chart = QWidget()
        self.tab_chart_layout = QVBoxLayout()
        self.chart_box = QWidget()
        self.chart_box_layout = QVBoxLayout(self.chart_box)
        self.chart_placeholder = QLabel(
            "点击「分析选定样品」后展示特征化合物柱状图")
        self.chart_placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.chart_placeholder.setMinimumHeight(300)
        self.chart_box_layout.addWidget(self.chart_placeholder)
        self.tab_chart_layout.addWidget(self.chart_box)
        self.tab_chart.setLayout(self.tab_chart_layout)
        self.result_tabs.addTab(self.tab_chart, "图表")

        layout.addWidget(self.result_tabs, 2)
        self.setLayout(layout)

    def load_data(self):
        with Session(self.parent().engine) as session:
            analyses = session.query(GCMSAnalysis).all()
            self.sample_table.setRowCount(len(analyses))
            for i, analysis in enumerate(analyses):
                # 注意：模型字段为 supplier，原代码误用 supplier_info 会崩溃
                self.sample_table.setItem(
                    i, 0, QTableWidgetItem(analysis.number or ''))
                self.sample_table.setItem(
                    i, 1, QTableWidgetItem(analysis.name or ''))
                self.sample_table.setItem(
                    i, 2, QTableWidgetItem(analysis.supplier or ''))
                self.sample_table.setItem(
                    i, 3, QTableWidgetItem(analysis.instrument_params or ''))

    def get_selected_analyses(self):
        selected = self.sample_table.selectionModel().selectedRows()
        if not selected:
            QMessageBox.warning(self, "警告", "请至少选择一个样品")
            return []
        # 选中的是样品编号（字符串），需按编号回查到 analysis.id
        numbers = [self.sample_table.item(r.row(), 0).text() for r in selected]
        with Session(self.parent().engine) as session:
            analyses = session.query(GCMSAnalysis).filter(
                GCMSAnalysis.number.in_(numbers)).all()
        if not analyses:
            QMessageBox.warning(self, "警告", "未找到相关样品数据")
            return []
        return analyses

    def analyze_selected_samples(self):
        analyses = self.get_selected_analyses()
        if not analyses:
            return
        self.analyze_characteristic_compounds(analyses)
        self.analyze_difference_compounds(analyses)
        self.result_tabs.setCurrentIndex(0)

    def analyze_characteristic_compounds(self, analyses):
        """特征物质：每个样品按相对含量降序取 Top10 化合物"""
        rows = []
        for a in analyses:
            with Session(self.parent().engine) as session:
                comps = session.query(GCMSCompound).filter(
                    GCMSCompound.analysis_id == a.id
                ).order_by(GCMSCompound.relative_content.desc()).limit(10).all()
            for rank, c in enumerate(comps, 1):
                rows.append([
                    a.number or '', a.name or '', str(rank),
                    _cname(c),
                    f"{(c.relative_content or 0):.2f}",
                    c.cas or ''
                ])
        self.char_table.setRowCount(len(rows))
        self.char_table.setColumnCount(6)
        for r, row in enumerate(rows):
            for col, val in enumerate(row):
                self.char_table.setItem(r, col, QTableWidgetItem(val))
        if not rows:
            QMessageBox.information(
                self, "提示", "所选样品暂无化合物数据，无法提取特征物质。")

    def analyze_difference_compounds(self, analyses):
        """差异物质：跨样品汇总化合物，标记共有/独有/部分共有"""
        ids = [a.id for a in analyses]
        names = {a.id: (a.number or a.name or '') for a in analyses}
        with Session(self.parent().engine) as session:
            comps = session.query(GCMSCompound).filter(
                GCMSCompound.analysis_id.in_(ids)).all()

        # 化合物 -> {analysis_id: relative_content}
        matrix = {}
        for c in comps:
            matrix.setdefault(_cname(c), {})[c.analysis_id] = c.relative_content or 0

        rows = []
        for name, per in matrix.items():
            n = len(per)
            if n == len(analyses):
                kind = '共有'
            elif n == 1:
                kind = '独有'
            else:
                kind = '部分共有'
            per_str = ' / '.join(
                f"{names[aid]}:{val:.2f}" for aid, val in per.items())
            rows.append([name, str(n), kind, per_str])
        rows.sort(key=lambda x: (-int(x[1]), x[0]))

        self.diff_table.setRowCount(len(rows))
        self.diff_table.setColumnCount(4)
        for r, row in enumerate(rows):
            for col, val in enumerate(row):
                self.diff_table.setItem(r, col, QTableWidgetItem(val))

        self._draw_chart(analyses)

    def _draw_chart(self, analyses):
        """绘制首个样品的特征化合物柱状图（相对含量 Top10）"""
        # 清空图表容器
        while self.chart_box_layout.count():
            item = self.chart_box_layout.takeAt(0)
            w = item.widget()
            if w is not None:
                w.deleteLater()

        try:
            a0 = analyses[0]
            with Session(self.parent().engine) as session:
                comps = session.query(GCMSCompound).filter(
                    GCMSCompound.analysis_id == a0.id
                ).order_by(GCMSCompound.relative_content.desc()).limit(10).all()
            if not comps:
                self.chart_box_layout.addWidget(
                    QLabel("该样品暂无化合物数据，无法绘图。"))
                return
            names = [_cname(c) for c in comps][::-1]
            vals = [(c.relative_content or 0) for c in comps][::-1]
            fig = Figure(figsize=(8, 5))
            ax = fig.add_subplot(111)
            ax.barh(names, vals, color='#2E86AB')
            ax.set_title(f"特征化合物（{a0.name or a0.number}）Top10 相对含量")
            ax.set_xlabel('相对含量')
            fig.tight_layout()
            canvas = FigureCanvas(fig)
            self.chart_box_layout.addWidget(canvas)
        except Exception as e:
            self.chart_box_layout.addWidget(
                QLabel(f"图表生成失败：{e}"))
