#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
智能分析模块
提供配方相似度分析、成本预测、库存优化建议与数据可视化
（依据字节码接口重建，保持与 main.py 调用约定一致）
"""
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from io import BytesIO

from PyQt6.QtWidgets import QMessageBox
from PyQt6.QtGui import QPixmap
from PyQt6.QtCore import Qt, QThread, pyqtSignal

import pandas as pd
PANDAS_AVAILABLE = True

try:
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei']
    plt.rcParams['axes.unicode_minus'] = False
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    plt = None
    MATPLOTLIB_AVAILABLE = False

try:
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.metrics.pairwise import cosine_similarity
    from sklearn.linear_model import LinearRegression
    import numpy as np
    SKLEARN_AVAILABLE = True
except ImportError:
    SKLEARN_AVAILABLE = False
    np = None

from models import (Ingredient, Formula, StockRecord,
                    ingredient_formula, FormulaUsage)


@dataclass
class AnalysisResult:
    """分析结果数据类"""
    title: str = ""
    summary: str = ""
    data: pd.DataFrame = None
    suggestions: list = field(default_factory=list)
    created_at: datetime = field(default_factory=datetime.now)


class FormulaAnalyzer:
    """配方分析器：相似度与关联分析"""

    def __init__(self, session):
        self.session = session

    def _formula_texts(self):
        """收集每个配方的文本特征（名称/描述/内容/成分名）"""
        formulas = self.session.query(Formula).all()
        texts, names = [], []
        for f in formulas:
            ing_names = [i.name for i in f.ingredients]
            text = " ".join(filter(None, [f.name, f.description,
                                          f.content, " ".join(ing_names)]))
            texts.append(text)
            names.append(f.name)
        return names, texts

    def analyze_similarity(self, target_formula):
        """分析目标配方与其他配方的相似度"""
        names, texts = self._formula_texts()
        if target_formula.name not in names or len(names) < 2:
            return None

        results = []
        if SKLEARN_AVAILABLE:
            vectorizer = TfidfVectorizer(analyzer='char_wb', ngram_range=(1, 2))
            try:
                matrix = vectorizer.fit_transform(texts)
            except ValueError:
                return None
            target_idx = names.index(target_formula.name)
            sims = cosine_similarity(matrix[target_idx:target_idx + 1], matrix).flatten()
            # 成分重合度加权
            for idx, sim in enumerate(sims):
                if idx == target_idx:
                    continue
                other = self.session.query(Formula).filter(
                    Formula.name == names[idx]).first()
                overlap = self._ingredient_overlap(target_formula, other)
                results.append((names[idx], float(sim), overlap))
        else:
            # 无 sklearn 时退化为成分重合度
            for idx, name in enumerate(names):
                if name == target_formula.name:
                    continue
                other = self.session.query(Formula).filter(
                    Formula.name == name).first()
                overlap = self._ingredient_overlap(target_formula, other)
                results.append((name, overlap, overlap))

        results.sort(key=lambda x: (x[1] + x[2]) / 2, reverse=True)
        return results[:10]

    def _ingredient_overlap(self, fa, fb):
        """两个配方的成分重合度（Jaccard）"""
        if not fa or not fb:
            return 0.0
        a = set(i.id for i in fa.ingredients)
        b = set(i.id for i in fb.ingredients)
        if not a or not b:
            return 0.0
        return len(a & b) / len(a | b)


class CostPredictor:
    """成本预测器：基于消耗速率预测采购成本"""

    def __init__(self, session):
        self.session = session

    def predict(self, days: int):
        """预测未来 days 天的原料消耗成本"""
        now = datetime.now()
        start = now - timedelta(days=90)
        records = self.session.query(StockRecord).filter(
            StockRecord.is_deleted == False,
            StockRecord.operation_time >= start
        ).all()

        consumption = {}  # number -> 90天累计出库量
        for r in records:
            if r.operation_type == 'out' and r.quantity:
                consumption[r.ingredient_number] = (
                    consumption.get(r.ingredient_number, 0) - r.quantity)

        rows = []
        for number, used in consumption.items():
            if used <= 0:
                continue
            ing = self.session.query(Ingredient).filter(
                Ingredient.number == number).first()
            if not ing:
                continue
            daily_rate = used / 90.0
            need = daily_rate * days
            price = ing.price or 0.0
            rows.append({
                '原料编号': number,
                '原料名称': ing.name,
                '日均消耗': round(daily_rate, 3),
                f'{days}天需求量': round(need, 2),
                '当前单价': price,
                '预计成本': round(need * price, 2)
            })

        df = pd.DataFrame(rows)
        if not df.empty:
            df = df.sort_values('预计成本', ascending=False)
        total = float(df['预计成本'].sum()) if not df.empty else 0.0
        return df, total


class InventoryOptimizer:
    """库存优化器：安全库存与采购建议"""

    def __init__(self, session):
        self.session = session

    def get_current_stock(self):
        """按原料编号汇总当前库存（未删除记录求和）"""
        stock_sums = dict(
            self.session.query(
                StockRecord.ingredient_number,
                func_sum(StockRecord.quantity)
            ).filter(StockRecord.is_deleted == False)
             .group_by(StockRecord.ingredient_number)
             .all()
        )
        return stock_sums

    def optimize(self):
        """生成库存预警与采购建议"""
        stock = self.get_current_stock()
        now = datetime.now()
        start = now - timedelta(days=90)
        records = self.session.query(StockRecord).filter(
            StockRecord.is_deleted == False,
            StockRecord.operation_time >= start
        ).all()

        consumption = {}
        for r in records:
            if r.operation_type == 'out' and r.quantity:
                consumption[r.ingredient_number] = (
                    consumption.get(r.ingredient_number, 0) - r.quantity)

        warnings, purchases = [], []
        ingredients = self.session.query(Ingredient).all()
        for ing in ingredients:
            current = stock.get(ing.number, 0) or 0
            threshold = ing.min_stock_threshold or 0
            daily_rate = consumption.get(ing.number, 0) / 90.0
            days_left = (current / daily_rate) if daily_rate > 0 else None

            if current <= 0 and daily_rate > 0:
                warnings.append((ing.number, ing.name, current, threshold,
                                 '缺货', '立即采购'))
            elif threshold > 0 and current <= threshold:
                warnings.append((ing.number, ing.name, round(current, 2),
                                 threshold, '低于安全库存',
                                 f'建议采购 {max(threshold * 2 - current, threshold):.1f}'))
            elif days_left is not None and days_left < 14:
                warnings.append((ing.number, ing.name, round(current, 2),
                                 threshold, f'约{days_left:.0f}天后耗尽',
                                 f'建议采购 {daily_rate * 30:.1f}（30天用量）'))

            if daily_rate > 0:
                suggested = max(daily_rate * 30 - current, 0)
                if suggested > 0:
                    purchases.append({
                        '原料编号': ing.number,
                        '原料名称': ing.name,
                        '当前库存': round(current, 2),
                        '日均消耗': round(daily_rate, 3),
                        '建议采购量': round(suggested, 2),
                        '预计费用': round(suggested * (ing.price or 0), 2)
                    })
        return warnings, purchases


def func_sum(col):
    """延迟导入聚合函数，避免顶层依赖"""
    from sqlalchemy import func
    return func.sum(col)


class ChartGenerator:
    """图表生成器"""

    def __init__(self, session):
        self.session = session

    def _fig_to_pixmap(self, fig):
        buf = BytesIO()
        fig.savefig(buf, format='png', dpi=110, bbox_inches='tight')
        plt.close(fig)
        buf.seek(0)
        pix = QPixmap()
        pix.loadFromData(buf.read())
        return pix

    def generate(self, chart_type: str):
        """按类型生成图表，返回 QPixmap"""
        if not MATPLOTLIB_AVAILABLE:
            return None
        if chart_type == "原料使用频率图":
            return self._ingredient_frequency()
        if chart_type == "成本趋势图":
            return self._cost_trend()
        if chart_type == "库存分布图":
            return self._stock_distribution()
        if chart_type == "配方复杂度分析":
            return self._formula_complexity()
        return None

    def _ingredient_frequency(self):
        rows = self.session.query(
            ingredient_formula.c.ingredient_id,
            func_count(ingredient_formula.c.formula_id).label('cnt')
        ).group_by(ingredient_formula.c.ingredient_id).all()
        freq = sorted(rows, key=lambda x: x[1], reverse=True)[:15]
        if not freq:
            return None
        names, counts = [], []
        for ing_id, cnt in freq:
            ing = self.session.query(Ingredient).get(ing_id)
            names.append(ing.name if ing else f'#{ing_id}')
            counts.append(cnt)
        fig, ax = plt.subplots(figsize=(9, 5))
        ax.barh(names[::-1], counts[::-1], color='#6F4E37')
        ax.set_title('原料使用频率 Top15（被配方引用次数）')
        ax.set_xlabel('引用次数')
        return self._fig_to_pixmap(fig)

    def _cost_trend(self):
        formulas = self.session.query(Formula).filter(
            Formula.total_cost.isnot(None)).order_by(Formula.created_at).all()
        if len(formulas) < 2:
            return None
        dates = [f.created_at.strftime('%m-%d') for f in formulas]
        costs = [f.total_cost for f in formulas]
        fig, ax = plt.subplots(figsize=(9, 5))
        ax.plot(dates, costs, marker='o', color='#D4A574', linewidth=2)
        ax.set_title('配方成本趋势')
        ax.set_ylabel('成本')
        ax.tick_params(axis='x', rotation=45)
        return self._fig_to_pixmap(fig)

    def _stock_distribution(self):
        opt = InventoryOptimizer(self.session)
        stock = {k: v for k, v in opt.get_current_stock().items() if v and v > 0}
        if not stock:
            return None
        items = sorted(stock.items(), key=lambda x: x[1], reverse=True)[:10]
        names, values = [], []
        for number, qty in items:
            ing = self.session.query(Ingredient).filter(
                Ingredient.number == number).first()
            names.append(ing.name if ing else number)
            values.append(qty)
        other = sum(stock.values()) - sum(values)
        if other > 0:
            names.append('其他')
            values.append(other)
        fig, ax = plt.subplots(figsize=(8, 6))
        ax.pie(values, labels=names, autopct='%1.1f%%', startangle=90)
        ax.set_title('库存分布（按数量 Top10）')
        return self._fig_to_pixmap(fig)

    def _formula_complexity(self):
        formulas = self.session.query(Formula).all()
        if not formulas:
            return None
        counts = [len(f.ingredients) for f in formulas]
        names = [f.name for f in formulas]
        fig, ax = plt.subplots(figsize=(9, 5))
        ax.bar(names, counts, color='#2E86AB')
        ax.set_title('配方复杂度分析（成分数量）')
        ax.set_ylabel('成分数量')
        ax.tick_params(axis='x', rotation=60)
        return self._fig_to_pixmap(fig)


def func_count(col):
    from sqlalchemy import func
    return func.count(col)


class AIAnalysisModule:
    """智能分析模块主类（与 main.py 调用约定保持一致）"""

    def __init__(self, session, parent=None):
        self.session = session
        self.parent = parent
        self.analyzer = FormulaAnalyzer(session)
        self.cost_predictor = CostPredictor(session)
        self.inventory_optimizer = InventoryOptimizer(session)
        self.chart_generator = ChartGenerator(session)

    # ---------- 配方相似度 ----------
    def load_formulas_to_combo(self, combo):
        """刷新配方下拉框"""
        current = combo.currentText()
        combo.clear()
        formulas = self.session.query(Formula).order_by(
            Formula.updated_at.desc()).all()
        for f in formulas:
            combo.addItem(f.name)
        idx = combo.findText(current)
        if idx >= 0:
            combo.setCurrentIndex(idx)

    def run_similarity_analysis(self, combo, result_widget):
        """执行相似度分析，结果写入 result_widget (QTextEdit)"""
        try:
            name = combo.currentText()
            if not name:
                QMessageBox.warning(self.parent, "提示", "请先选择目标配方")
                return
            target = self.session.query(Formula).filter(
                Formula.name == name).first()
            if not target:
                QMessageBox.warning(self.parent, "提示", "未找到目标配方")
                return
            results = self.analyzer.analyze_similarity(target)
            if results is None:
                result_widget.setPlainText(
                    "数据不足：至少需要两个配方且配方需包含成分信息才能进行相似度分析。")
                return
            lines = [f"目标配方：{target.name}", "-" * 50]
            lines.append(f"{'配方名称':<20}{'文本相似度':<12}{'成分重合度':<12}")
            for fname, sim, overlap in results:
                lines.append(f"{fname:<20}{sim:>10.2%}{overlap:>11.2%}")
            lines.append("")
            lines.append("💡 优化建议：")
            if results:
                best = results[0]
                lines.append(
                    f"· 与「{best[0]}」关联度最高，可参考其配方思路进行优化；")
            lines.append("· 成分重合度高（>50%）的配方可考虑合并或复用；")
            lines.append("· 相似度仅基于文本与成分信息，实际调香请结合感官评价。")
            result_widget.setPlainText("\n".join(lines))
        except Exception as e:
            QMessageBox.critical(self.parent, "错误", f"相似度分析失败：{e}")

    # ---------- 成本预测 ----------
    def run_cost_prediction(self, days_spin, result_widget):
        """执行成本预测（days_spin 为 QSpinBox）"""
        try:
            days = days_spin.value()
            df, total = self.cost_predictor.predict(days)
            lines = [f"📈 未来 {days} 天采购成本预测（基于近90天消耗）", "=" * 50]
            if df is None or df.empty:
                lines.append("暂无足够的出库记录，无法预测。")
                lines.append("提示：录入配方使用（出库）记录后可获得预测结果。")
            else:
                lines.append(df.to_string(index=False))
                lines.append("")
                lines.append(f"预计总采购成本：¥{total:,.2f}")
                lines.append("")
                top = df.iloc[0]
                lines.append(f"💡 成本大头：「{top['原料名称']}」"
                             f"预计 ¥{top['预计成本']:,.2f}，"
                             f"可优先洽谈批量采购。")
            result_widget.setPlainText("\n".join(lines))
        except Exception as e:
            QMessageBox.critical(self.parent, "错误", f"成本预测失败：{e}")

    # ---------- 库存优化 ----------
    def run_inventory_optimization(self, result_widget):
        """生成库存优化与采购建议"""
        try:
            warnings, purchases = self.inventory_optimizer.optimize()
            lines = ["📦 库存智能优化建议", "=" * 50]
            if warnings:
                lines.append(f"⚠️ 库存预警（{len(warnings)} 项）：")
                lines.append(f"{'编号':<12}{'名称':<16}{'库存':<10}"
                             f"{'安全线':<10}{'状态':<14}{'建议'}")
                for w in warnings:
                    lines.append(f"{w[0]:<12}{w[1]:<16}{w[2]!s:<10}"
                                 f"{w[3]!s:<10}{w[4]:<14}{w[5]}")
                lines.append("")
            else:
                lines.append("✅ 当前库存状态良好，无预警项。")
                lines.append("")
            if purchases:
                df = pd.DataFrame(purchases).sort_values(
                    '预计费用', ascending=False)
                lines.append(f"🎯 采购建议（按费用排序）：")
                lines.append(df.to_string(index=False))
                total = df['预计费用'].sum()
                lines.append("")
                lines.append(f"30天用量补货总预算：约 ¥{total:,.2f}")
            result_widget.setPlainText("\n".join(lines))
        except Exception as e:
            QMessageBox.critical(self.parent, "错误", f"库存优化失败：{e}")

    # ---------- 数据可视化 ----------
    def generate_chart(self, chart_type_combo, chart_label):
        """生成图表并显示在 chart_label (QLabel) 上"""
        try:
            chart_type = chart_type_combo.currentText()
            pixmap = self.chart_generator.generate(chart_type)
            if pixmap is None:
                QMessageBox.information(
                    self.parent, "提示",
                    f"无法生成「{chart_type}」：数据不足（需要更多配方/库存记录）。")
                return
            chart_label.setText("")
            chart_label.setPixmap(pixmap)
            chart_label.setScaledContents(False)
            chart_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        except Exception as e:
            QMessageBox.critical(self.parent, "错误", f"图表生成失败：{e}")
