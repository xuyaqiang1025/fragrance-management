#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# 优先使用本地Python环境，避免虚拟环境冲突
import sys
import os
import math

# 控制台输出统一为 UTF-8（打包为无窗口程序时避免GBK编码报错）
for _stream in (sys.stdout, sys.stderr):
    if _stream is not None and hasattr(_stream, 'reconfigure'):
        try:
            _stream.reconfigure(encoding='utf-8', errors='replace')
        except Exception:
            pass

# 强制使用本地Python环境的包路径，优先级高于虚拟环境
local_python_path = r"C:\Python313\Lib\site-packages"
if os.path.exists(local_python_path) and local_python_path not in sys.path:
    # 将本地Python包路径插入到最前面，确保优先级
    sys.path.insert(0, local_python_path)
    print(f"✅ 优先使用本地Python环境: {local_python_path}")

# 添加当前目录和上级目录到Python路径
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
sys.path.insert(0, current_dir)
sys.path.insert(0, parent_dir)

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                           QHBoxLayout, QPushButton, QLabel, QStackedWidget,
                           QTableWidget, QTableWidgetItem, QLineEdit, QFormLayout,
                           QMessageBox, QHeaderView, QFileDialog, QDialog, QToolTip,
                           QComboBox, QSpinBox, QDoubleSpinBox, QListWidget, QListWidgetItem,
                           QDialogButtonBox, QTabWidget, QTextEdit, QGroupBox, QAbstractItemView,
                           QDateEdit)
from PyQt6.QtGui import QPixmap, QColor
from PyQt6.QtCore import Qt, QTimer, QThread, pyqtSignal, QSettings, QDateTime, QDate
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, joinedload

# 导入模型和分析模块
try:
    from models import Base, Ingredient, Formula, StockRecord, Supplier, GCMSAnalysis, GCMSCompound, ingredient_formula, FormulaUsage
    print("✅ 数据模型导入成功")
except ImportError as e:
    print(f"❌ 数据模型导入失败: {e}")
    sys.exit(1)

try:
    from ai_analysis_module import AIAnalysisModule
    print("✅ AI分析模块导入成功")
except ImportError as e:
    print(f"⚠️  AI分析模块导入失败: {e}")
    AIAnalysisModule = None

try:
    from flavor_expert_module import FlavorExpertModule
    print("✅ AI调香专家模块导入成功")
except ImportError as e:
    print(f"⚠️  AI调香专家模块导入失败: {e}")
    FlavorExpertModule = None

# 打包(windowed)模式下重定向输出到日志文件，避免句柄缺失
if getattr(sys, 'frozen', False) and (sys.stdout is None or sys.stderr is None):
    try:
        _log_path = os.path.join(os.path.dirname(sys.executable),
                                 'fragrance_app.log')
        _log_file = open(_log_path, 'a', encoding='utf-8')
        sys.stdout = _log_file
        sys.stderr = _log_file
    except Exception:
        pass

# 条件导入pandas和相关依赖包
try:
    import pandas as pd  # 重新启用pandas
    PANDAS_AVAILABLE = True
    print("✅ pandas导入成功")
except ImportError as e:
    print(f"⚠️  pandas导入失败: {e}")
    print("GC-MS分析功能中的数据处理功能将受限")
    PANDAS_AVAILABLE = False
    # 创建一个简单的替代类
    class MockDataFrame:
        def __init__(self, data=None, columns=None):
            self.data = data or []
            self.columns = columns or []
        def empty(self):
            return len(self.data) == 0
        def head(self, n=5):
            return self
        def groupby(self, *args):
            return self
        def sum(self):
            return self
        def to_csv(self, *args, **kwargs):
            raise NotImplementedError("pandas不可用，无法导出CSV文件。建议安装pandas或手动复制数据。")
        def to_excel(self, *args, **kwargs):
            raise NotImplementedError("pandas不可用，无法导出Excel文件。建议安装pandas或手动复制数据。")
        def reset_index(self):
            return self
        def pivot_table(self, *args, **kwargs):
            return self
        def iterrows(self):
            # 返回空迭代器
            return iter([])
        def get(self, key, default=None):
            return default
        def __getitem__(self, key):
            return []
    pd = type('pd', (), {'DataFrame': MockDataFrame})()

import requests
from io import BytesIO

# 条件导入matplotlib
try:
    import matplotlib  # 重新启用matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei']
    plt.rcParams['axes.unicode_minus'] = False
    try:
        # matplotlib>=3.5 的 Qt6 通用后端
        from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
    except ImportError:
        from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
    MATPLOTLIB_AVAILABLE = True
    print("✅ matplotlib导入成功")
except ImportError as e:
    print(f"⚠️  matplotlib导入失败: {e}")
    print("GC-MS分析功能中的图表生成功能将受限")
    MATPLOTLIB_AVAILABLE = False
    plt = None

from datetime import datetime, timedelta

# 条件导入seaborn
try:
    import seaborn as sns  # 重新启用seaborn
    SEABORN_AVAILABLE = True
    print("✅ seaborn导入成功")
except ImportError as e:
    print(f"⚠️  seaborn导入失败: {e}")
    SEABORN_AVAILABLE = False
    sns = None

import hashlib
import os

# 条件导入openpyxl
try:
    from openpyxl import Workbook  # 重新启用openpyxl
    OPENPYXL_AVAILABLE = True
    print("✅ openpyxl导入成功")
except ImportError as e:
    print(f"⚠️  openpyxl导入失败: {e}")
    OPENPYXL_AVAILABLE = False

# 确保 PyQt6-Charts 已安装: pip install PyQt6-Charts
from PyQt6.QtCharts import QChart, QChartView, QBarSeries, QBarSet, QValueAxis, QDateTimeAxis, QPieSeries, QBarCategoryAxis, QLineSeries
from sqlalchemy import func, text # 导入 func 和 text
import sqlite3

# Custom QTableWidgetItem for numeric sorting
class NumericTableWidgetItem(QTableWidgetItem):
    def __lt__(self, other):
        # Compare items based on their UserRole data (numeric value)
        try:
            data1 = self.data(Qt.ItemDataRole.UserRole)
            data2 = other.data(Qt.ItemDataRole.UserRole)
            #print(f"Comparing data: {data1} < {data2}") # Added print statement
            return self.data(Qt.ItemDataRole.UserRole) < other.data(Qt.ItemDataRole.UserRole)
        except TypeError:
            # Fallback to text comparison if user data is not comparable (e.g., None)
            return super().__lt__(other)

class PagedTableLoader(QThread):
    data_loaded = pyqtSignal(list)
    def __init__(self, engine, model_class, filter_text="", offset=0, limit=50, sort_column=None, sort_order='asc'):
        super().__init__()
        self.engine = engine
        self.model_class = model_class
        self.filter_text = filter_text
        self.offset = offset
        self.limit = limit
        self.sort_column = sort_column  # 排序列名
        self.sort_order = sort_order    # 排序方向: 'asc' 或 'desc'
        self._is_running = True

    def run(self):
        try:
            if self._is_running:
                # 在新线程中创建新的session
                from sqlalchemy.orm import sessionmaker
                Session = sessionmaker(bind=self.engine)
                session = Session()
                
                try:
                    # 构建查询
                    query = session.query(self.model_class)
                    
                    # 添加过滤条件（根据模型类型）
                    if self.filter_text:
                        from models import Ingredient, Formula
                        if self.model_class == Ingredient:
                            like = f"%{self.filter_text}%"
                            query = query.filter(
                                Ingredient.name.like(like) |
                                Ingredient.cas_number.like(like) |
                                Ingredient.english_name.like(like) |
                                Ingredient.molecular_formula.like(like)
                            )
                        elif self.model_class == Formula:
                            like = f"%{self.filter_text}%"
                            query = query.filter(
                                Formula.name.like(like) |
                                Formula.description.like(like) |
                                Formula.creator.like(like)
                            )
                        elif hasattr(self.model_class, 'name'):
                            query = query.filter(self.model_class.name.contains(self.filter_text))
                        elif hasattr(self.model_class, 'number'):
                            query = query.filter(self.model_class.number.contains(self.filter_text))
                    
                    # 添加排序
                    if self.sort_column and hasattr(self.model_class, self.sort_column):
                        sort_attr = getattr(self.model_class, self.sort_column)
                        if self.sort_order == 'desc':
                            query = query.order_by(sort_attr.desc())
                        else:
                            query = query.order_by(sort_attr.asc())
                    else:
                        # 默认按id排序，确保结果一致性
                        query = query.order_by(self.model_class.id.asc())
                    
                    # 执行分页查询
                    data = query.offset(self.offset).limit(self.limit).all()
                    if self._is_running:  # 再次检查，以防在查询过程中被终止
                        self.data_loaded.emit(data)
                finally:
                    session.close()
        except Exception as e:
            print(f"Error in PagedTableLoader: {e}")

    def stop(self):
        self._is_running = False
        self.wait()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        # 先初始化分页参数
        self.stock_record_page = 1
        self.stock_record_page_size = 50
        self.stock_summary_page = 1
        self.stock_summary_page_size = 50
        self.stock_summary_filter = ''  # ← 这里必须加上
        self.stock_summary_sort_col = 0
        self.stock_summary_sort_order = Qt.SortOrder.AscendingOrder
        self.stock_record_page_label = QLabel()  # 初始化 stock_record_page_label
        self.stock_record_prev_btn = QPushButton("上一页")  # 初始化 stock_record_prev_btn
        self.stock_record_next_btn = QPushButton("下一页")  # 初始化 stock_record_next_btn
        self.setWindowTitle("香精原料与配方管理系统")
        self.setMinimumSize(1200, 800)
        self.ingredient_page = 1
        self.ingredient_page_size = 50
        self.formula_page = 1
        self.formula_page_size = 50
        self.gcms_page = 1 # Added for GCMS pagination
        self.gcms_page_size = 50 # Added for GCMS pagination
        
        # 排序状态管理
        self.ingredient_sort_column = None
        self.ingredient_sort_order = 'asc'
        self.formula_sort_column = None
        self.formula_sort_order = 'asc'
        
        # 初始化加载器为 None
        self.ingredient_loader = None
        self.formula_loader = None
        
        # 初始化数据库
        self.init_database()
        
        # 创建主窗口部件
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        # 创建主布局
        layout = QHBoxLayout()
        main_widget.setLayout(layout)
        
        # 创建侧边栏
        sidebar = QWidget()
        sidebar.setMaximumWidth(180)  # 减少侧边栏宽度
        sidebar.setStyleSheet("""
            QWidget {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #f8f9fa, stop:1 #e9ecef);
                border-right: 2px solid #dee2e6;
                font-family: 'Microsoft YaHei UI', sans-serif;
            }
        """)
        sidebar_layout = QVBoxLayout()
        sidebar_layout.setSpacing(8)  # 减少按钮间距
        sidebar_layout.setContentsMargins(12, 20, 12, 20)
        sidebar.setLayout(sidebar_layout)
        
        # 添加标题
        title_label = QLabel("香料管理")
        title_label.setStyleSheet("""
            QLabel {
                font-size: 18px;
                font-weight: bold;
                color: #495057;
                background: none;
                border: none;
                padding: 10px 0px;
            }
        """)
        sidebar_layout.addWidget(title_label)
        
        # 添加侧边栏按钮
        buttons = [
            ("原料管理", self.show_ingredients),
            ("配方管理", self.show_formulas),
            ("库存管理", self.show_stock),
            ("供应商管理", self.show_suppliers),
            ("GC-MS分析", self.show_gcms),
            ("AI调香专家", self.show_flavor_expert),
            ("智能分析", self.show_ai_analysis),
            ("数据统计", self.show_statistics)
        ]
        
        for text, slot in buttons:
            btn = QPushButton(text)
            btn.setMinimumHeight(40)
            btn.clicked.connect(slot)
            sidebar_layout.addWidget(btn)
        
        sidebar_layout.addStretch()
        
        # 创建内容区域
        self.content = QStackedWidget()
        
        # 添加内容页面
        self.content.addWidget(self.create_ingredients_page())
        self.content.addWidget(self.create_formulas_page())
        self.content.addWidget(self.create_stock_page())
        self.content.addWidget(self.create_suppliers_page())
        self.content.addWidget(self.create_gcms_page())
        self.content.addWidget(self.create_ai_analysis_page())
        self.content.addWidget(self.create_statistics_page())
        if FlavorExpertModule is not None:
            self.flavor_expert = FlavorExpertModule(self.session, self)
            self.content.addWidget(self.flavor_expert)
        
        # 将侧边栏和内容区域添加到主布局
        layout.addWidget(sidebar)
        layout.addWidget(self.content)
        
        # 启动时刷新一次原料表
        self.refresh_ingredient_table()

        # 状态栏常驻显示当前数据库路径，避免误判数据没保存
        if hasattr(self, 'db_path'):
            self.statusBar().showMessage(f"数据库：{self.db_path}")
            self.statusBar().setToolTip(
                "所有增删改都写入这个文件；源码运行与打包exe会各自使用独立的数据库")

    def init_database(self):
        """初始化数据库"""
        try:
            # 数据库固定放在程序目录（打包后为exe所在目录），避免随工作目录漂移
            if getattr(sys, 'frozen', False):
                base_dir = os.path.dirname(os.path.abspath(sys.executable))
            else:
                base_dir = current_dir
            db_path = os.path.join(base_dir, 'fragrance_management.db')
            db_existed = os.path.exists(db_path)
            # 记录实际使用的数据库路径：
            # 从源码运行(src/)与运行打包exe(release/)会指向不同的库，
            # 打印出来便于排查「改了数据但看起来没变」的问题
            self.db_path = db_path
            print(f"数据库路径: {db_path}（{'已存在' if db_existed else '新建'}）")
            # 添加线程安全配置
            engine = create_engine(f'sqlite:///{db_path}',
                                 pool_pre_ping=True,
                                 connect_args={'check_same_thread': False})
            Base.metadata.create_all(engine)
            Session = sessionmaker(bind=engine)
            self.session = Session()
            self.engine = engine  # 保存引擎引用

            # 初始化智能分析模块
            if AIAnalysisModule is not None:
                self.ai_analysis = AIAnalysisModule(self.session, self)
            else:
                self.ai_analysis = None
            
            # 检查并添加示例配方数据（仅在数据库首次创建时，一次性播种）
            self._init_sample_data(db_existed)
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"数据库初始化失败：{str(e)}")
    
    def _init_sample_data(self, db_existed=False):
        """初始化示例数据（一次性播种：通过app_settings标记，删除后不再恢复）

        健壮性增强（修复「删除配方后重载又复活」的根因之一）：
        - 无论新库还是旧库，只要标记缺失就补写标记，避免任何历史版本在后续
          启动时被反复重新播种示例配方；
        - 播种前按编号去重，防止极端情况下重复插入；
        - 仅当是全新数据库且无任何配方时才播种，已存在数据的库绝不改动。
        """
        try:
            from models import AppSetting
            seeded = self.session.query(AppSetting).filter_by(
                key='sample_formulas_seeded').first()
            if seeded:
                return
            # 仅当是全新数据库且无任何配方时才播种示例
            if not db_existed:
                formula_count = self.session.query(Formula).count()
                if formula_count == 0:
                    # 添加示例配方
                    sample_formulas = [
                    {
                        'number': 'F001',
                        'name': '清新薄荷',
                        'creator': '调香师A',
                        'description': '清新薄荷香型配方',
                        'content': '薄荷醇:30%:3.0g,香兰素:20%:2.0g,乙基麦芽酚:10%:1.0g',
                        'evaluation': '清香怡人，薄荷味突出',
                        'total_cost': 15.5
                    },
                    {
                        'number': 'F002',
                        'name': '果香混合',
                        'creator': '调香师B',
                        'description': '混合果香配方',
                        'content': '覆盆子酮:25%:2.5g,香叶醇:20%:2.0g,柠檬醛:15%:1.5g',
                        'evaluation': '果香浓郁，层次丰富',
                        'total_cost': 22.8
                    },
                    {
                        'number': 'F003',
                        'name': '经典烟草',
                        'creator': '调香师C',
                        'description': '经典烟草香型',
                        'content': '香兰素:30%:3.0g,乙基香兰素:25%:2.5g,麦芽酚:15%:1.5g',
                        'evaluation': '醇厚烟草香，回味悠长',
                        'total_cost': 18.9
                    }
                    ]
                    existing_numbers = {
                        f.number for f in self.session.query(Formula.number).all()}
                    added = 0
                    for formula_data in sample_formulas:
                        if formula_data['number'] in existing_numbers:
                            continue
                        self.session.add(Formula(**formula_data))
                        added += 1
                    if added:
                        self.session.commit()
                        print(f"已添加 {added} 个示例配方")
            # 无论如何都补写标记，确保只播种一次（根治删除后复活）
            self.session.add(AppSetting(key='sample_formulas_seeded', value='1'))
            self.session.commit()
        except Exception as e:
            print(f"初始化示例数据时出错: {e}")
            try:
                self.session.rollback()
            except Exception:
                pass
    
    def create_ingredients_page(self):
        """创建原料管理页面"""
        page = QWidget()
        layout = QVBoxLayout()
        page.setLayout(layout)
        
        # 创建搜索和按钮区域
        top_layout = QHBoxLayout()

        # 添加搜索框
        search_label = QLabel("搜索：")
        self.ingredient_search = QLineEdit()
        self.ingredient_search.setPlaceholderText("输入原料名称、CAS号或供应商")
        self.ingredient_search.textChanged.connect(self.refresh_ingredient_table)
        top_layout.addWidget(search_label)
        top_layout.addWidget(self.ingredient_search)

        # 添加按钮
        add_btn = QPushButton("添加原料")
        add_btn.clicked.connect(self.add_ingredient_dialog)
        top_layout.addWidget(add_btn)

        import_btn = QPushButton("批量导入")
        import_btn.clicked.connect(self.import_ingredients)
        top_layout.addWidget(import_btn)

        export_btn = QPushButton("导出")
        export_btn.clicked.connect(self.export_ingredients)
        top_layout.addWidget(export_btn)

        # 添加分页控件 (放在表格下方)
        page_layout = QHBoxLayout()
        self.ingredient_page_label = QLabel()
        self.ingredient_prev_btn = QPushButton("上一页")
        self.ingredient_next_btn = QPushButton("下一页")
        self.ingredient_prev_btn.clicked.connect(self.ingredient_prev_page)
        self.ingredient_next_btn.clicked.connect(self.ingredient_next_page)
        page_layout.addWidget(self.ingredient_prev_btn)
        page_layout.addWidget(self.ingredient_page_label)
        page_layout.addWidget(self.ingredient_next_btn)
        page_layout.addStretch() # Add stretch to push pagination to the right

        # 创建表格
        self.ingredient_table = QTableWidget()
        self.ingredient_table.setColumnCount(19) # 增加成本列
        self.ingredient_table.setHorizontalHeaderLabels([
            "编号", "CAS", "原料名称", "英文名", "分子式", "化学结构式", "分子量", "沸点", "溶解性", 
            "电子烟国标最大用量", "天然存在", "调香用途", "香气香韵", "香气变调", 
            "嗅香香气（1% in PG）", "抽吸感官评价", "香韵构成", "成本(元/g)", "操作"
        ])
        self.ingredient_table.setSortingEnabled(False)  # 禁用内置排序，使用数据库级别排序
        
        # 设置表格列宽
        self.ingredient_table.setColumnWidth(0, 100)  # 编号
        self.ingredient_table.setColumnWidth(1, 100)  # CAS
        self.ingredient_table.setColumnWidth(2, 150)  # 原料名称
        self.ingredient_table.setColumnWidth(3, 150)  # 英文名
        self.ingredient_table.setColumnWidth(4, 100)  # 分子式
        self.ingredient_table.setColumnWidth(5, 150)  # 化学结构式
        self.ingredient_table.setColumnWidth(6, 80)   # 分子量
        self.ingredient_table.setColumnWidth(7, 80)   # 沸点
        self.ingredient_table.setColumnWidth(8, 100)   # 溶解性
        self.ingredient_table.setColumnWidth(9, 150)   # 国标用量
        self.ingredient_table.setColumnWidth(10, 100)  # 天然存在
        self.ingredient_table.setColumnWidth(11, 150)  # 调香用途
        self.ingredient_table.setColumnWidth(12, 150)  # 香气香韵
        self.ingredient_table.setColumnWidth(13, 150)  # 香气变化
        self.ingredient_table.setColumnWidth(14, 150)  # 吸香香气
        self.ingredient_table.setColumnWidth(15, 150)  # 感官评价
        self.ingredient_table.setColumnWidth(16, 150)  # 香的构成
        self.ingredient_table.setColumnWidth(17, 100)  # 成本(元/g)
        self.ingredient_table.setColumnWidth(18, 150)  # 操作
        
        # 连接表格头部点击事件，实现数据库级别排序
        self.ingredient_table.horizontalHeader().sectionClicked.connect(self.handle_ingredient_sort)
        
        # 添加所有控件到主布局
        layout.addLayout(top_layout)
        layout.addWidget(self.ingredient_table)
        layout.addLayout(page_layout)
        
        return page
    
    def create_formulas_page(self):
        page = QWidget()
        layout = QVBoxLayout()
        page.setLayout(layout)
        
        # 创建搜索和按钮区域
        top_layout = QHBoxLayout()

        # 添加搜索框
        search_label = QLabel("搜索：")
        self.formula_search = QLineEdit()
        self.formula_search.setPlaceholderText("输入配方名称或描述")
        self.formula_search.textChanged.connect(self.refresh_formula_table)
        top_layout.addWidget(search_label)
        top_layout.addWidget(self.formula_search)

        # 添加按钮
        add_btn = QPushButton("添加配方")
        add_btn.clicked.connect(self.add_formula_dialog)
        top_layout.addWidget(add_btn)

        # 添加导出按钮
        export_btn = QPushButton("导出")
        export_btn.clicked.connect(self.export_formulas)
        top_layout.addWidget(export_btn)

        # 添加分页控件 (放在表格下方)
        page_layout = QHBoxLayout()
        self.formula_page_label = QLabel()
        self.formula_prev_btn = QPushButton("上一页")
        self.formula_next_btn = QPushButton("下一页")
        self.formula_prev_btn.clicked.connect(self.formula_prev_page)
        self.formula_next_btn.clicked.connect(self.formula_next_page)
        page_layout.addWidget(self.formula_prev_btn)
        page_layout.addWidget(self.formula_page_label)
        page_layout.addWidget(self.formula_next_btn)
        page_layout.addStretch() # Add stretch to push pagination to the right

        # 创建表格
        self.formula_table = QTableWidget()
        self.formula_table.setColumnCount(8)
        self.formula_table.setHorizontalHeaderLabels([
            "编号", "名称", "创建人", "创建时间", "修改时间", "感官评价", "总成本(元)", "操作"
        ])
        self.formula_table.setSortingEnabled(False)  # 禁用内置排序，使用数据库级别排序
        
        # 连接表格头部点击事件，实现数据库级别排序
        self.formula_table.horizontalHeader().sectionClicked.connect(self.handle_formula_sort)
        
        # 设置表格列宽
        self.formula_table.setColumnWidth(0, 150)  # 名称
        self.formula_table.setColumnWidth(1, 200)  # 描述
        self.formula_table.setColumnWidth(2, 150)  # 创建时间
        self.formula_table.setColumnWidth(3, 150)  # 修改时间
        self.formula_table.setColumnWidth(4, 150)  # 感官评价
        self.formula_table.setColumnWidth(5, 100) # 总成本
        self.formula_table.setColumnWidth(6, 200)  # 操作
        self.formula_table.setColumnWidth(7, 100)  # 详情 (removed as a separate column)
        
        # Add double click to show formula detail
        #self.formula_table.cellDoubleClicked.connect(self.show_formula_detail)

        # 添加所有控件到主布局
        layout.addLayout(top_layout)
        layout.addWidget(self.formula_table)
        layout.addLayout(page_layout)
        
        return page
    
    def create_stock_page(self):
        page = QWidget()
        layout = QVBoxLayout()
        page.setLayout(layout)
        # 预警区域
        warning_group = QGroupBox("库存预警")
        warning_layout = QVBoxLayout()
        # 刷新、导出按钮
        warning_btn_layout = QHBoxLayout()
        refresh_warning_btn = QPushButton("刷新")
        export_warning_btn = QPushButton("导出")
        warning_btn_layout.addWidget(refresh_warning_btn)
        warning_btn_layout.addWidget(export_warning_btn)
        warning_btn_layout.addStretch()
        warning_layout.addLayout(warning_btn_layout)
        self.warning_table = QTableWidget()
        self.warning_table.setColumnCount(5)
        self.warning_table.setHorizontalHeaderLabels(["原料编号", "原料名称", "当前库存", "预警类型", "详情"])
        self.warning_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        warning_layout.addWidget(self.warning_table)
        warning_group.setLayout(warning_layout)
        layout.addWidget(warning_group)
        # 总览区域
        summary_group = QGroupBox("当前库存总览")
        summary_layout = QVBoxLayout()
        # 筛选、导出、分页控件在一行
        summary_top_layout = QHBoxLayout()
        self.stock_summary_filter_edit = QLineEdit()
        self.stock_summary_filter_edit.setPlaceholderText("筛选原料编号/名称/供应商...")
        summary_top_layout.addWidget(self.stock_summary_filter_edit)
        export_summary_btn = QPushButton("导出")
        summary_top_layout.addWidget(export_summary_btn)
        # 新增刷新按钮
        refresh_summary_btn = QPushButton("刷新")
        summary_top_layout.addWidget(refresh_summary_btn)
        # 分页控件
        self.stock_summary_prev_btn = QPushButton("上一页")
        self.stock_summary_next_btn = QPushButton("下一页")
        self.stock_summary_page_label = QLabel()
        summary_top_layout.addWidget(self.stock_summary_prev_btn)
        summary_top_layout.addWidget(self.stock_summary_page_label)
        summary_top_layout.addWidget(self.stock_summary_next_btn)
        summary_top_layout.addStretch()
        summary_layout.addLayout(summary_top_layout)
        self.stock_summary_table = QTableWidget()
        self.stock_summary_table.setColumnCount(5)
        self.stock_summary_table.setHorizontalHeaderLabels(["原料编号", "原料名称", "当前库存（g）", "预警阈值（g）", "供应商"])
        self.stock_summary_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.stock_summary_table.setSortingEnabled(True)
        self.stock_summary_table.horizontalHeader().sectionClicked.connect(self.handle_stock_summary_sort)
        summary_layout.addWidget(self.stock_summary_table)
        summary_group.setLayout(summary_layout)
        layout.addWidget(summary_group)
        # 库存记录区域
        records_group = QGroupBox("库存记录")
        records_layout = QVBoxLayout()
        # 搜索、添加、分页在同一行
        top_layout = QHBoxLayout()
        self.stock_search = QLineEdit()
        self.stock_search.setPlaceholderText("搜索原料编号/名称/供应商/批次...")
        top_layout.addWidget(self.stock_search)
        add_btn = QPushButton("添加库存记录")
        top_layout.addWidget(add_btn)
        # 分页控件
        top_layout.addWidget(self.stock_record_prev_btn)
        top_layout.addWidget(self.stock_record_page_label)
        top_layout.addWidget(self.stock_record_next_btn)
        top_layout.addStretch()
        records_layout.addLayout(top_layout)
        self.stock_table = QTableWidget()
        self.stock_table.setColumnCount(10)
        self.stock_table.setHorizontalHeaderLabels([
            "原料编号", "原料名称", "数量（g）", "供应商", "批次号", "操作类型", "操作时间", "有效期", "操作人", "删除"
        ])
        self.stock_table.setSortingEnabled(True)
        self.stock_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        records_layout.addWidget(self.stock_table)
        records_group.setLayout(records_layout)
        layout.addWidget(records_group)
        # 信号连接
        refresh_warning_btn.clicked.connect(self.refresh_stock_warning)
        export_warning_btn.clicked.connect(lambda: self.export_table_data(self.warning_table, "库存预警"))
        export_summary_btn.clicked.connect(lambda: self.export_table_data(self.stock_summary_table, "库存总览"))
        refresh_summary_btn.clicked.connect(self.refresh_stock_summary)  # 新增刷新按钮信号
        add_btn.clicked.connect(self.add_stock_dialog)
        self.stock_search.textChanged.connect(self.refresh_stock_table)
        self.stock_record_prev_btn.clicked.connect(self.stock_record_prev_page)
        self.stock_record_next_btn.clicked.connect(self.stock_record_next_page)
        # 初始化
        self.refresh_stock_warning()
        self.refresh_stock_summary()
        self.refresh_stock_table()
        # 信号连接
        self.stock_summary_filter_edit.textChanged.connect(self.handle_stock_summary_filter)
        self.stock_summary_prev_btn.clicked.connect(self.stock_summary_prev_page)
        self.stock_summary_next_btn.clicked.connect(self.stock_summary_next_page)
        return page

    def delete_stock_record(self, record):
        """逻辑删除库存记录"""
        record.is_deleted = True
        self.session.commit()
        self.refresh_stock_table()
        self.refresh_stock_summary()

    def refresh_stock_warning(self):
        self.warning_table.setRowCount(0)
        warnings = []
        # 只遍历有库存记录的原料
        ingredient_ids = set(r.ingredient_id for r in self.session.query(StockRecord).filter(StockRecord.is_deleted == False).all())
        for ingredient in self.session.query(Ingredient).filter(Ingredient.id.in_(ingredient_ids)).all():
            current_stock = sum(r.quantity for r in ingredient.stock_records if not r.is_deleted)
            threshold = ingredient.min_stock_threshold or 0
            if current_stock < threshold:
                warnings.append([
                    ingredient.number,
                    ingredient.name,
                    f"{current_stock:.2f}",
                    "库存不足",
                    f"最小库存: {threshold}g"
                ])
            elif current_stock < threshold * 1.2 and threshold > 0:
                warnings.append([
                    ingredient.number,
                    ingredient.name,
                    f"{current_stock:.2f}",
                    "接近阈值",
                    f"最小库存: {threshold}g"
                ])
        self.warning_table.setRowCount(len(warnings))
        for row, data in enumerate(warnings):
            for col, value in enumerate(data):
                self.warning_table.setItem(row, col, QTableWidgetItem(str(value)))

    def refresh_stock_summary(self):
        self.session.expire_all()
        self.stock_summary_table.setRowCount(0)
        # 用 ingredient_number 分组统计所有未删除的库存
        stock_sums = dict(
            self.session.query(
                StockRecord.ingredient_number,
                func.sum(StockRecord.quantity)
            ).filter(StockRecord.is_deleted == False)
             .group_by(StockRecord.ingredient_number)
             .all()
        )
        # 只显示库存大于0的原料编号
        numbers = [num for num, qty in stock_sums.items() if qty > 0]
        print("stock_sums:", stock_sums)
        print("numbers:", numbers)
        print("Ingredient表：", [(i.id, i.number, i.name) for i in self.session.query(Ingredient).all()])
        if not numbers:
            self.stock_summary_table.setRowCount(0)
            self.stock_summary_page_label.setText("第1/1页 共0条")
            self.stock_summary_prev_btn.setEnabled(False)
            self.stock_summary_next_btn.setEnabled(False)
            return
        query = self.session.query(Ingredient).filter(Ingredient.number.in_(numbers))
        # 筛选
        if self.stock_summary_filter:
            like = f"%{self.stock_summary_filter}%"
            query = query.filter(
                Ingredient.number.like(like) |
                Ingredient.name.like(like)
            )
        all_ings = list(query)
        print("all_ings:", [(ing.id, ing.number, ing.name) for ing in all_ings])  # 调试输出
        # 排序
        def get_sort_key(ing):
            if self.stock_summary_sort_col == 0:
                return ing.number
            elif self.stock_summary_sort_col == 1:
                return ing.name
            elif self.stock_summary_sort_col == 2:
                return stock_sums.get(ing.number, 0)
            elif self.stock_summary_sort_col == 3:
                return ing.min_stock_threshold or 0
            elif self.stock_summary_sort_col == 4:
                last_in = self.session.query(StockRecord).filter(
                    StockRecord.ingredient_number == ing.number,
                    StockRecord.operation_type == 'in',
                    StockRecord.is_deleted == False
                ).order_by(StockRecord.operation_time.desc()).first()
                return last_in.supplier if last_in and last_in.supplier else ""
            return ing.number
        all_ings.sort(key=get_sort_key, reverse=self.stock_summary_sort_order == Qt.SortOrder.DescendingOrder)
        # 分页
        total = len(all_ings)
        max_page = max(1, (total + self.stock_summary_page_size - 1) // self.stock_summary_page_size)
        self.stock_summary_page = min(self.stock_summary_page, max_page)
        offset = (self.stock_summary_page - 1) * self.stock_summary_page_size
        page_ings = all_ings[offset:offset + self.stock_summary_page_size]
        self.stock_summary_page_label.setText(f"第{self.stock_summary_page}/{max_page}页 共{total}条")
        self.stock_summary_prev_btn.setEnabled(self.stock_summary_page > 1)
        self.stock_summary_next_btn.setEnabled(self.stock_summary_page < max_page)
        # 填充表格
        self.stock_summary_table.setRowCount(len(page_ings))  # <--- 关键修复
        for row, ingredient in enumerate(page_ings):
            current_stock = stock_sums.get(ingredient.number, 0)
            last_in = self.session.query(StockRecord).filter(
                StockRecord.ingredient_number == ingredient.number,
                StockRecord.operation_type == 'in',
                StockRecord.is_deleted == False
            ).order_by(StockRecord.operation_time.desc()).first()
            supplier = last_in.supplier if last_in and last_in.supplier else ""
            for col, value in enumerate([
                ingredient.number,
                ingredient.name,
                f"{current_stock:.2f}",
                ingredient.min_stock_threshold or 0,
                supplier
            ]):
                if col == 3:
                    from PyQt6.QtWidgets import QDoubleSpinBox
                    spin = QDoubleSpinBox()
                    spin.setRange(0, 999999)
                    spin.setDecimals(2)
                    spin.setSingleStep(1)
                    spin.setValue(float(value))
                    spin.ingredient_id = ingredient.id
                    spin.valueChanged.connect(self.update_min_stock_threshold)
                    self.stock_summary_table.setCellWidget(row, col, spin)
                else:
                    self.stock_summary_table.setItem(row, col, QTableWidgetItem(str(value)))

    def update_min_stock_threshold(self, value):
        spin = self.sender()
        if not hasattr(spin, 'ingredient_id'):
            return
        ing = self.session.query(Ingredient).get(spin.ingredient_id)
        if ing:
            ing.min_stock_threshold = value
            self.session.commit()
            self.refresh_stock_warning()

    def refresh_stock_table(self):
        self.stock_table.setRowCount(0)
        keyword = self.stock_search.text().strip()
        query = self.session.query(StockRecord).filter(StockRecord.is_deleted == False)
        if keyword:
            query = query.filter(
                StockRecord.ingredient_number.like(f"%{keyword}%") |
                StockRecord.ingredient_name.like(f"%{keyword}%") |
                StockRecord.supplier.like(f"%{keyword}%") |
                StockRecord.batch_number.like(f"%{keyword}%")
            )
        total = query.count()
        max_page = max(1, (total + self.stock_record_page_size - 1) // self.stock_record_page_size)
        self.stock_record_page = min(self.stock_record_page, max_page)
        offset = (self.stock_record_page - 1) * self.stock_record_page_size
        records = query.order_by(StockRecord.operation_time.desc()).offset(offset).limit(self.stock_record_page_size).all()
        self.stock_record_page_label.setText(f"第{self.stock_record_page}/{max_page}页 共{total}条")
        self.stock_record_prev_btn.setEnabled(self.stock_record_page > 1)
        self.stock_record_next_btn.setEnabled(self.stock_record_page < max_page)
        self.stock_table.setRowCount(len(records))
        for row, r in enumerate(records):
            values = [
                r.ingredient_number,
                r.ingredient_name,
                f"{r.quantity:.2f}",
                r.supplier,
                r.batch_number,
                r.operation_type,
                r.operation_time.strftime("%Y-%m-%d %H:%M:%S") if r.operation_time else "",
                r.expiration_date.strftime("%Y-%m-%d") if r.expiration_date else "",
                r.operator,
                "删除"
            ]
            for col, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                if col == 9:
                    btn = QPushButton("删除")
                    btn.clicked.connect(lambda _, rec=r: self.delete_stock_record(rec))
                    self.stock_table.setCellWidget(row, col, btn)
                else:
                    self.stock_table.setItem(row, col, item)
        self.refresh_stock_summary()  # 自动联动刷新库存总览

    def export_table_data(self, table, title):
        """导出表格数据到Excel"""
        if table.rowCount() == 0:
            QMessageBox.warning(self, "警告", "没有数据可导出！")
            return
            
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            f"导出{title}",
            "",
            "Excel Files (*.xlsx);;CSV Files (*.csv)"
        )
        
        if not file_path:
            return
            
        try:
            data = []
            headers = []
            
            # 获取表头
            for col in range(table.columnCount()):
                headers.append(table.horizontalHeaderItem(col).text())
            
            # 获取数据
            for row in range(table.rowCount()):
                row_data = []
                for col in range(table.columnCount()):
                    item = table.item(row, col)
                    row_data.append(item.text() if item else "")
                data.append(row_data)
            
            # 创建DataFrame并导出
            df = pd.DataFrame(data, columns=headers)
            
            if file_path.endswith('.xlsx'):
                df.to_excel(file_path, index=False)
            else:
                df.to_csv(file_path, index=False, encoding='utf-8-sig')
                
            QMessageBox.information(self, "成功", f"{title}导出成功！")
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"导出失败：{str(e)}")

    def create_suppliers_page(self):
        """创建供应商管理页面"""
        page = QWidget()
        layout = QVBoxLayout()
        page.setLayout(layout)
        
        # 创建搜索和添加区域
        top_layout = QHBoxLayout()
        self.supplier_search = QLineEdit()
        self.supplier_search.setPlaceholderText("搜索供应商...")
        self.supplier_search.textChanged.connect(self.refresh_supplier_table)
        add_button = QPushButton("添加供应商")
        add_button.clicked.connect(self.add_supplier_dialog)
        export_button = QPushButton("导出")
        export_button.clicked.connect(self.export_suppliers)
        
        top_layout.addWidget(self.supplier_search)
        top_layout.addWidget(add_button)
        top_layout.addWidget(export_button)
        
        # 创建表格
        self.supplier_table = QTableWidget()
        self.supplier_table.setColumnCount(7)
        self.supplier_table.setHorizontalHeaderLabels([
            "ID", "名称", "联系人", "电话", "邮箱", "地址", "操作"
        ])
        header = self.supplier_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(6, QHeaderView.ResizeMode.ResizeToContents)
        
        layout.addLayout(top_layout)
        layout.addWidget(self.supplier_table)
        
        # 初始化数据
        self.refresh_supplier_table()
        
        return page
    
    def refresh_supplier_table(self):
        """刷新供应商表格"""
        try:
            # 获取搜索关键词
            keyword = self.supplier_search.text().strip() if hasattr(self, 'supplier_search') else ""
            
            # 查询供应商
            query = self.session.query(Supplier)
            if keyword:
                query = query.filter(
                    (Supplier.name.like(f"%{keyword}%")) |
                    (Supplier.contact_person.like(f"%{keyword}%")) |
                    (Supplier.phone.like(f"%{keyword}%")) |
                    (Supplier.email.like(f"%{keyword}%"))
                )
            
            suppliers = query.all()
            
            # 更新表格
            self.supplier_table.setRowCount(len(suppliers))
            
            for row, supplier in enumerate(suppliers):
                self.supplier_table.setItem(row, 0, QTableWidgetItem(str(supplier.id)))
                self.supplier_table.setItem(row, 1, QTableWidgetItem(supplier.name or ""))
                self.supplier_table.setItem(row, 2, QTableWidgetItem(supplier.contact_person or ""))
                self.supplier_table.setItem(row, 3, QTableWidgetItem(supplier.phone or ""))
                self.supplier_table.setItem(row, 4, QTableWidgetItem(supplier.email or ""))
                self.supplier_table.setItem(row, 5, QTableWidgetItem(supplier.address or ""))
                
                # 操作按钮
                btn_layout = QHBoxLayout()
                edit_btn = QPushButton("编辑")
                edit_btn.clicked.connect(lambda checked, s=supplier: self.edit_supplier_dialog(s))
                delete_btn = QPushButton("删除")
                delete_btn.clicked.connect(lambda checked, s=supplier: self.delete_supplier(s))
                delete_btn.setStyleSheet("color: red;")
                
                btn_layout.addWidget(edit_btn)
                btn_layout.addWidget(delete_btn)
                
                btn_widget = QWidget()
                btn_widget.setLayout(btn_layout)
                self.supplier_table.setCellWidget(row, 6, btn_widget)
                
        except Exception as e:
            QMessageBox.warning(self, "错误", f"刷新供应商表格失败：{str(e)}")

    def add_supplier_dialog(self):
        """添加供应商对话框"""
        dialog = SupplierDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            data = dialog.get_data()
            try:
                supplier = Supplier(**data)
                self.session.add(supplier)
                self.session.commit()
                self.refresh_supplier_table()
                QMessageBox.information(self, "成功", "供应商添加成功！")
            except Exception as e:
                self.session.rollback()
                QMessageBox.critical(self, "错误", f"添加供应商失败：{str(e)}")

    def edit_supplier_dialog(self, supplier):
        """编辑供应商对话框"""
        dialog = SupplierDialog(self, supplier)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            data = dialog.get_data()
            try:
                for key, value in data.items():
                    setattr(supplier, key, value)
                supplier.updated_at = datetime.now()
                self.session.commit()
                self.refresh_supplier_table()
                QMessageBox.information(self, "成功", "供应商信息更新成功！")
            except Exception as e:
                self.session.rollback()
                QMessageBox.critical(self, "错误", f"更新供应商失败：{str(e)}")

    def delete_supplier(self, supplier):
        """删除供应商"""
        reply = QMessageBox.question(
            self, "确认删除", 
            f"确定要删除供应商 '{supplier.name}' 吗？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            try:
                self.session.delete(supplier)
                self.session.commit()
                self.refresh_supplier_table()
                QMessageBox.information(self, "成功", "供应商删除成功！")
            except Exception as e:
                self.session.rollback()
                QMessageBox.critical(self, "错误", f"删除供应商失败：{str(e)}")

    def export_suppliers(self):
        """导出供应商数据"""
        try:
            suppliers = self.session.query(Supplier).all()
            
            if not suppliers:
                QMessageBox.information(self, "提示", "没有供应商数据可导出")
                return
            
            filename, _ = QFileDialog.getSaveFileName(
                self, "导出供应商数据", "suppliers_export.csv", "CSV Files (*.csv)"
            )
            
            if filename:
                # 使用内置csv模块代替pandas
                import csv
                with open(filename, 'w', newline='', encoding='utf-8-sig') as csvfile:
                    fieldnames = ['ID', '名称', '联系人', '电话', '邮箱', '地址', '创建时间', '更新时间']
                    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                    
                    writer.writeheader()
                    for supplier in suppliers:
                        writer.writerow({
                            'ID': supplier.id,
                            '名称': supplier.name,
                            '联系人': supplier.contact_person,
                            '电话': supplier.phone,
                            '邮箱': supplier.email,
                            '地址': supplier.address,
                            '创建时间': supplier.created_at.strftime('%Y-%m-%d %H:%M:%S') if supplier.created_at else '',
                            '更新时间': supplier.updated_at.strftime('%Y-%m-%d %H:%M:%S') if supplier.updated_at else ''
                        })
                
                QMessageBox.information(self, "成功", f"供应商数据已导出到：{filename}")
                
        except Exception as e:
            QMessageBox.critical(self, "错误", f"导出失败：{str(e)}")
    
    def create_gcms_page(self):
        page = QWidget()
        layout = QVBoxLayout()
        page.setLayout(layout)

        # 创建搜索和按钮区域
        top_layout = QHBoxLayout()

        # 添加搜索框
        search_label = QLabel("搜索：")
        self.gcms_search = QLineEdit()
        self.gcms_search.setPlaceholderText("输入样品名称或描述")
        self.gcms_search.textChanged.connect(self.refresh_gcms_table)
        top_layout.addWidget(search_label)
        top_layout.addWidget(self.gcms_search)

        # 添加按钮
        add_btn = QPushButton("添加分析")
        add_btn.clicked.connect(self.add_gcms_dialog)
        top_layout.addWidget(add_btn)

        analysis_btn = QPushButton("分析")
        analysis_btn.clicked.connect(self.show_gcms_analysis_func)
        top_layout.addWidget(analysis_btn)

        export_btn = QPushButton("导出")
        export_btn.clicked.connect(self.export_gcms_analyses) # Connect export button
        top_layout.addWidget(export_btn)

        # 添加分页控件 (放在表格下方)
        page_layout = QHBoxLayout()
        self.gcms_page_label = QLabel()
        self.gcms_prev_btn = QPushButton("上一页")
        self.gcms_next_btn = QPushButton("下一页")
        self.gcms_prev_btn.clicked.connect(self.gcms_prev_page)
        self.gcms_next_btn.clicked.connect(self.gcms_next_page)
        page_layout.addWidget(self.gcms_prev_btn)
        page_layout.addWidget(self.gcms_page_label)
        page_layout.addWidget(self.gcms_next_btn)
        page_layout.addStretch() # Add stretch to push pagination to the right

        # 创建表格
        self.gcms_table = QTableWidget()
        self.gcms_table.setColumnCount(7) # 增加谱图列
        self.gcms_table.setHorizontalHeaderLabels([
            "编号", "样品名称", "供应商", "仪器参数", "整体调香思路", "谱图", "操作"
        ])

        # Set column resize mode
        header = self.gcms_table.horizontalHeader()
        for i in range(header.count()):
             header.setSectionResizeMode(i, QHeaderView.ResizeMode.Interactive)
        header.setSectionsMovable(True)

        # 设置表格列宽 (Adjust widths based on new columns)
        self.gcms_table.setColumnWidth(0, 100)  # 编号
        self.gcms_table.setColumnWidth(1, 150)  # 样品名称
        self.gcms_table.setColumnWidth(2, 150)  # 供应商
        self.gcms_table.setColumnWidth(3, 150)  # 仪器参数
        self.gcms_table.setColumnWidth(4, 200)  # 整体调香思路
        self.gcms_table.setColumnWidth(5, 100)  # 谱图
        self.gcms_table.setColumnWidth(6, 150)  # 操作 (Edit, Delete)

        # 添加所有控件到主布局
        layout.addLayout(top_layout)
        layout.addWidget(self.gcms_table)
        layout.addLayout(page_layout)

        # Load column widths from settings
        self.load_gcms_column_widths()

        return page
    
    def create_ai_analysis_page(self):
        """创建智能分析页面"""
        page = QWidget()
        layout = QVBoxLayout()
        page.setLayout(layout)
        
        # 页面标题和状态指示器
        header_layout = QHBoxLayout()
        title_label = QLabel("🤖 智能分析模块")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; color: #2E86AB; margin: 10px;")
        header_layout.addWidget(title_label)
        
        # 依赖库状态指示器
        status_label = QLabel()
        status_text = "🟢 核心功能正常"
        try:
            # 检查关键依赖
            import pandas as pd
            import matplotlib.pyplot as plt
            try:
                from sklearn.feature_extraction.text import TfidfVectorizer
                status_text = "🟢 全功能可用 (包含机器学习)"
            except ImportError:
                status_text = "🟡 基础功能可用 (缺少scikit-learn)"
        except ImportError:
            status_text = "🔴 功能受限 (缺少依赖库)"
            
        status_label.setText(status_text)
        status_label.setStyleSheet("font-size: 12px; color: #666; margin: 5px;")
        header_layout.addWidget(status_label)
        header_layout.addStretch()
        
        layout.addLayout(header_layout)
        
        # 创建选项卡
        tabs = QTabWidget()
        
        # 配方相似度分析选项卡
        similarity_tab = QWidget()
        similarity_layout = QVBoxLayout()
        
        # 功能描述
        desc_label = QLabel("📊 配方相似度分析")
        desc_label.setStyleSheet("font-weight: bold; font-size: 14px; color: #2E86AB;")
        similarity_layout.addWidget(desc_label)
        
        info_label = QLabel("基于TF-IDF算法和成分相似性分析配方关联度，提供智能配方推荐和优化建议")
        info_label.setWordWrap(True)
        info_label.setStyleSheet("color: #666; margin-bottom: 10px;")
        similarity_layout.addWidget(info_label)
        
        # 配方选择区域
        formula_group = QGroupBox("配方选择")
        formula_group_layout = QVBoxLayout()
        
        formula_select_layout = QHBoxLayout()
        formula_select_layout.addWidget(QLabel("目标配方:"))
        self.target_formula_combo = QComboBox()
        self.target_formula_combo.setMinimumWidth(300)
        # 配方下拉框将在show_ai_analysis方法中初始化
        formula_select_layout.addWidget(self.target_formula_combo)
        
        # 刷新配方列表按钮
        refresh_btn = QPushButton("🔄 刷新")
        refresh_btn.clicked.connect(lambda: self.ai_analysis.load_formulas_to_combo(self.target_formula_combo))
        formula_select_layout.addWidget(refresh_btn)
        formula_select_layout.addStretch()
        
        formula_group_layout.addLayout(formula_select_layout)
        formula_group.setLayout(formula_group_layout)
        similarity_layout.addWidget(formula_group)
        
        # 分析控制区域
        analysis_layout = QHBoxLayout()
        similarity_btn = QPushButton("🔍 执行相似度分析")
        similarity_btn.setStyleSheet("QPushButton { background-color: #4CAF50; color: white; font-weight: bold; padding: 8px; }")
        similarity_btn.clicked.connect(lambda: self.ai_analysis.run_similarity_analysis(self.target_formula_combo, self.similarity_result))
        analysis_layout.addWidget(similarity_btn)
        analysis_layout.addStretch()
        
        similarity_layout.addLayout(analysis_layout)
        
        # 结果显示区域
        result_group = QGroupBox("分析结果")
        result_layout = QVBoxLayout()
        self.similarity_result = QTextEdit()
        self.similarity_result.setMinimumHeight(250)
        self.similarity_result.setStyleSheet("background-color: #f8f9fa; border: 1px solid #dee2e6;")
        result_layout.addWidget(self.similarity_result)
        result_group.setLayout(result_layout)
        similarity_layout.addWidget(result_group)
        
        similarity_tab.setLayout(similarity_layout)
        tabs.addTab(similarity_tab, "配方相似度")
        
        # 成本预测选项卡
        cost_tab = QWidget()
        cost_layout = QVBoxLayout()
        
        # 功能描述
        cost_desc_label = QLabel("💰 成本预测与优化")
        cost_desc_label.setStyleSheet("font-weight: bold; font-size: 14px; color: #2E86AB;")
        cost_layout.addWidget(cost_desc_label)
        
        cost_info_label = QLabel("基于历史数据和市场趋势预测原料价格变化，提供成本优化和采购时机建议")
        cost_info_label.setWordWrap(True)
        cost_info_label.setStyleSheet("color: #666; margin-bottom: 10px;")
        cost_layout.addWidget(cost_info_label)
        
        # 预测参数设置
        param_group = QGroupBox("预测参数")
        param_layout = QHBoxLayout()
        param_layout.addWidget(QLabel("预测时间范围:"))
        self.predict_days = QSpinBox()
        self.predict_days.setRange(1, 365)
        self.predict_days.setValue(30)
        self.predict_days.setSuffix(" 天")
        param_layout.addWidget(self.predict_days)
        
        # 预测按钮
        predict_btn = QPushButton("📈 执行价格预测")
        predict_btn.setStyleSheet("QPushButton { background-color: #FF9800; color: white; font-weight: bold; padding: 8px; }")
        predict_btn.clicked.connect(lambda: self.ai_analysis.run_cost_prediction(self.predict_days, self.cost_result))
        param_layout.addWidget(predict_btn)
        param_layout.addStretch()
        
        param_group.setLayout(param_layout)
        cost_layout.addWidget(param_group)
        
        # 结果显示
        cost_result_group = QGroupBox("预测结果")
        cost_result_layout = QVBoxLayout()
        self.cost_result = QTextEdit()
        self.cost_result.setMinimumHeight(250)
        self.cost_result.setStyleSheet("background-color: #f8f9fa; border: 1px solid #dee2e6;")
        cost_result_layout.addWidget(self.cost_result)
        cost_result_group.setLayout(cost_result_layout)
        cost_layout.addWidget(cost_result_group)
        
        cost_tab.setLayout(cost_layout)
        tabs.addTab(cost_tab, "成本预测")
        
        # 库存优化选项卡
        inventory_tab = QWidget()
        inventory_layout = QVBoxLayout()
        
        # 功能描述
        inv_desc_label = QLabel("📦 库存智能优化")
        inv_desc_label.setStyleSheet("font-weight: bold; font-size: 14px; color: #2E86AB;")
        inventory_layout.addWidget(inv_desc_label)
        
        inv_info_label = QLabel("分析原料使用模式和库存状态，提供智能采购建议、安全库存设置和库存预警")
        inv_info_label.setWordWrap(True)
        inv_info_label.setStyleSheet("color: #666; margin-bottom: 10px;")
        inventory_layout.addWidget(inv_info_label)
        
        # 分析控制
        inv_control_layout = QHBoxLayout()
        inventory_btn = QPushButton("🎯 生成采购建议")
        inventory_btn.setStyleSheet("QPushButton { background-color: #2196F3; color: white; font-weight: bold; padding: 8px; }")
        inventory_btn.clicked.connect(lambda: self.ai_analysis.run_inventory_optimization(self.inventory_result))
        inv_control_layout.addWidget(inventory_btn)
        inv_control_layout.addStretch()
        
        inventory_layout.addLayout(inv_control_layout)
        
        # 结果显示
        inv_result_group = QGroupBox("优化建议")
        inv_result_layout = QVBoxLayout()
        self.inventory_result = QTextEdit()
        self.inventory_result.setMinimumHeight(250)
        self.inventory_result.setStyleSheet("background-color: #f8f9fa; border: 1px solid #dee2e6;")
        inv_result_layout.addWidget(self.inventory_result)
        inv_result_group.setLayout(inv_result_layout)
        inventory_layout.addWidget(inv_result_group)
        
        inventory_tab.setLayout(inventory_layout)
        tabs.addTab(inventory_tab, "库存优化")
        
        # 数据可视化选项卡
        viz_tab = QWidget()
        viz_layout = QVBoxLayout()
        
        # 功能描述
        viz_desc_label = QLabel("📊 数据可视化")
        viz_desc_label.setStyleSheet("font-weight: bold; font-size: 14px; color: #2E86AB;")
        viz_layout.addWidget(viz_desc_label)
        
        viz_info_label = QLabel("生成各种分析图表和可视化报告，包括原料使用统计、成本趋势、库存分布等")
        viz_info_label.setWordWrap(True)
        viz_info_label.setStyleSheet("color: #666; margin-bottom: 10px;")
        viz_layout.addWidget(viz_info_label)
        
        # 图表控制区域
        chart_control_group = QGroupBox("图表设置")
        chart_control_layout = QVBoxLayout()
        
        # 图表类型选择
        chart_type_layout = QHBoxLayout()
        chart_type_layout.addWidget(QLabel("图表类型:"))
        self.chart_type = QComboBox()
        self.chart_type.addItems([
            "原料使用频率图",
            "成本趋势图", 
            "库存分布图",
            "配方复杂度分析"
        ])
        self.chart_type.setMinimumWidth(200)
        chart_type_layout.addWidget(self.chart_type)
        
        # 生成图表按钮
        chart_btn = QPushButton("📈 生成图表")
        chart_btn.setStyleSheet("QPushButton { background-color: #9C27B0; color: white; font-weight: bold; padding: 8px; }")
        chart_btn.clicked.connect(lambda: self.ai_analysis.generate_chart(self.chart_type, self.chart_label))
        chart_type_layout.addWidget(chart_btn)
        chart_type_layout.addStretch()
        
        chart_control_layout.addLayout(chart_type_layout)
        chart_control_group.setLayout(chart_control_layout)
        viz_layout.addWidget(chart_control_group)
        
        # 图表显示区域
        chart_display_group = QGroupBox("图表显示")
        chart_display_layout = QVBoxLayout()
        self.chart_label = QLabel("📊 点击'生成图表'按钮开始可视化分析")
        self.chart_label.setMinimumHeight(350)
        self.chart_label.setStyleSheet("""
            QLabel {
                border: 2px dashed #ccc; 
                background-color: #f9f9f9;
                color: #666;
                font-size: 14px;
            }
        """)
        self.chart_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        chart_display_layout.addWidget(self.chart_label)
        chart_display_group.setLayout(chart_display_layout)
        viz_layout.addWidget(chart_display_group)
        
        viz_tab.setLayout(viz_layout)
        tabs.addTab(viz_tab, "数据可视化")
        
        layout.addWidget(tabs)
        
        return page
        
    def create_statistics_page(self):
        """创建数据统计页面"""
        page = QWidget()
        layout = QVBoxLayout()
        page.setLayout(layout)
        
        # 添加统计类型选择
        type_layout = QHBoxLayout()
        type_label = QLabel("统计类型：")
        self.stat_type = QComboBox()
        self.stat_type.addItems([
            "原料使用频率统计",
            "配方成本统计",
            "GC-MS分析统计",
            "GC-MS化合物频次统计"
        ])
        self.stat_type.currentTextChanged.connect(self.update_statistics)
        type_layout.addWidget(type_label)
        type_layout.addWidget(self.stat_type)
        
        # 添加时间范围选择
        time_layout = QHBoxLayout()
        time_label = QLabel("时间范围：")
        self.time_range = QComboBox()
        self.time_range.addItems([
            "全部",
            "最近一周",
            "最近一月",
            "最近三月",
            "最近一年"
        ])
        self.time_range.currentTextChanged.connect(self.update_statistics)
        time_layout.addWidget(time_label)
        time_layout.addWidget(self.time_range)
        
        # 添加导出按钮
        export_btn = QPushButton("导出统计结果")
        export_btn.clicked.connect(self.export_statistics)
        time_layout.addWidget(export_btn)
        
        # 创建图表显示区域
        from PyQt6.QtCharts import QChartView
        self.chart_view = QChartView()
        self.chart_view.setMinimumHeight(400)
        
        # 创建数据表格
        self.stat_table = QTableWidget()
        self.stat_table.setMinimumHeight(300)
        
        # 添加所有控件到主布局
        layout.addLayout(type_layout)
        layout.addLayout(time_layout)
        layout.addWidget(self.chart_view)
        layout.addWidget(self.stat_table)
        
        return page
        
    def update_statistics(self):
        """更新统计数据"""
        # Removed permission check
        
        stat_type = self.stat_type.currentText()
        time_range = self.time_range.currentText()
        
        # 根据统计类型和时间范围获取数据
        if stat_type == "原料使用频率统计":
            self.show_ingredient_usage_stats(time_range)
        elif stat_type == "配方成本统计":
            self.show_formula_cost_stats(time_range)
        elif stat_type == "GC-MS分析统计":
            self.show_gcms_analysis_stats(time_range)
        elif stat_type == "GC-MS化合物频次统计":
            self.show_gcms_compound_stats(time_range)
        # Removed Operation Log Stats as UserOperation table is removed
        # elif stat_type == "操作日志统计":
        #     self.show_operation_log_stats(time_range)
            
    def show_ingredient_usage_stats(self, time_range):
        """显示原料使用频率统计（按配方引用次数 Top20）"""
        # 获取时间范围（按配方创建时间过滤）
        start_date = self.get_start_date(time_range)

        # 查询数据
        query = self.session.query(
            ingredient_formula.c.ingredient_id,
            Ingredient.name,
            func.count(ingredient_formula.c.formula_id).label('usage_count')
        ).join(
            Ingredient,
            ingredient_formula.c.ingredient_id == Ingredient.id
        ).join(
            Formula,
            ingredient_formula.c.formula_id == Formula.id
        )

        if start_date:
            query = query.filter(Formula.created_at >= start_date)

        results = query.group_by(
            ingredient_formula.c.ingredient_id,
            Ingredient.name
        ).order_by(
            text('usage_count DESC')
        ).limit(20).all()

        if not results:
            self._show_empty_chart(
                "原料使用频率统计",
                f"所选时间范围（{time_range}）内暂无配方成分数据。")
            return

        # 全局引用总次数：用于计算真实占比，
        # 不能用 Top20 之和做分母（那样占比会被放大且合计不为 100%）
        total_query = self.session.query(
            func.count(ingredient_formula.c.formula_id)
        ).select_from(ingredient_formula).join(
            Formula, ingredient_formula.c.formula_id == Formula.id
        )
        if start_date:
            total_query = total_query.filter(Formula.created_at >= start_date)
        total_usage = total_query.scalar() or 0

        # 创建图表
        chart = QChart()
        chart.setTitle(f"原料使用频率统计 Top20（{time_range}）")

        # 创建柱状图系列
        series = QBarSeries()
        bar_set = QBarSet("使用次数")

        # 添加数据
        categories = []
        for result in results:
            bar_set.append(result.usage_count)
            categories.append(result.name or '')

        series.append(bar_set)
        chart.addSeries(series)

        # 设置坐标轴
        axis_x = QBarCategoryAxis()
        axis_x.append(categories)
        chart.addAxis(axis_x, Qt.AlignmentFlag.AlignBottom)
        series.attachAxis(axis_x)

        axis_y = QValueAxis()
        max_usage = max(r.usage_count for r in results)
        axis_y.setRange(0, max_usage * 1.1 if max_usage > 0 else 10)
        axis_y.setTitleText("被配方引用次数")
        chart.addAxis(axis_y, Qt.AlignmentFlag.AlignLeft)
        series.attachAxis(axis_y)

        # 更新图表视图
        self.chart_view.setChart(chart)

        # 更新数据表格
        self.stat_table.setColumnCount(3)
        self.stat_table.setHorizontalHeaderLabels(["原料名称", "使用次数", "占比"])
        self.stat_table.setRowCount(len(results))

        for row, result in enumerate(results):
            self.stat_table.setItem(row, 0, QTableWidgetItem(result.name or ''))
            self.stat_table.setItem(row, 1, QTableWidgetItem(str(result.usage_count)))
            usage_rate = (result.usage_count / total_usage * 100) if total_usage > 0 else 0
            self.stat_table.setItem(row, 2, QTableWidgetItem(f"{usage_rate:.2f}%"))
            
    def show_formula_cost_stats(self, time_range):
        """显示配方成本统计"""
        # 获取时间范围
        start_date = self.get_start_date(time_range)
        
        # 查询数据
        query = self.session.query(
            Formula.name,
            Formula.total_cost,
            Formula.created_at
        )
        
        if start_date:
            query = query.filter(Formula.created_at >= start_date)
            
        results = query.order_by(Formula.created_at).all()

        # 仅保留可绘制的数据点（成本与时间均有效）
        points = [r for r in results
                  if r.created_at is not None and r.total_cost is not None]

        if not points:
            self._show_empty_chart(
                "配方成本统计",
                f"所选时间范围（{time_range}）内暂无带成本数据的配方。\n"
                f"（共 {len(results)} 个配方，但均未记录成本）")
            return

        # 创建图表
        chart = QChart()
        chart.setTitle(f"配方成本统计（{time_range}）")

        # 创建折线图系列
        series = QLineSeries()
        series.setName("配方成本")

        # 添加数据
        for result in points:
            # 将Python datetime转换为QDateTime
            qdt = QDateTime.fromString(
                result.created_at.strftime("%Y-%m-%d %H:%M:%S"),
                "yyyy-MM-dd hh:mm:ss")
            series.append(qdt.toMSecsSinceEpoch(), result.total_cost)

        chart.addSeries(series)

        # 设置坐标轴
        axis_x = QDateTimeAxis()
        axis_x.setFormat("yyyy-MM-dd")
        axis_x.setTitleText("日期")
        chart.addAxis(axis_x, Qt.AlignmentFlag.AlignBottom)
        series.attachAxis(axis_x)

        axis_y = QValueAxis()
        # 只对非空成本求最大值，避免 None 参与比较导致 TypeError
        costs = [r.total_cost for r in points]
        max_cost = max(costs)
        axis_y.setRange(0, max_cost * 1.1 if max_cost > 0 else 10)
        axis_y.setTitleText("成本")
        chart.addAxis(axis_y, Qt.AlignmentFlag.AlignLeft)
        series.attachAxis(axis_y)

        # 更新图表视图
        self.chart_view.setChart(chart)

        # 更新数据表格
        self.stat_table.setColumnCount(3)
        self.stat_table.setHorizontalHeaderLabels(["配方名称", "成本", "创建时间"])
        self.stat_table.setRowCount(len(results))

        for row, result in enumerate(results):
            self.stat_table.setItem(row, 0, QTableWidgetItem(result.name or ""))
            self.stat_table.setItem(row, 1, QTableWidgetItem(
                f"{result.total_cost:,.2f}" if result.total_cost is not None else "未记录"))
            self.stat_table.setItem(row, 2, QTableWidgetItem(
                result.created_at.strftime("%Y-%m-%d %H:%M")
                if result.created_at else ""
            ))
            
    def _show_empty_chart(self, title, message):
        """在图表区显示空数据提示，并清空统计表格"""
        chart = QChart()
        chart.setTitle(title)
        self.chart_view.setChart(chart)
        self.stat_table.setColumnCount(1)
        self.stat_table.setHorizontalHeaderLabels(["提示"])
        self.stat_table.setRowCount(1)
        self.stat_table.setItem(0, 0, QTableWidgetItem(message))

    def show_gcms_analysis_stats(self, time_range):
        """显示GC-MS分析统计（按供应商分布 + 样品明细）

        注意：GCMSAnalysis 实际字段为 number/name/supplier/analysis_time，
        原先误用了不存在的 status/sample_name/created_at，导致统计恒为「未知」。
        """
        try:
            start_date = self.get_start_date(time_range)

            query = self.session.query(GCMSAnalysis)
            if start_date:
                # 无分析时间的记录视为不限时间范围，避免被误过滤掉
                query = query.filter(
                    (GCMSAnalysis.analysis_time >= start_date) |
                    (GCMSAnalysis.analysis_time.is_(None))
                )
            results = query.order_by(GCMSAnalysis.analysis_time.desc()).all()

            if not results:
                self._show_empty_chart(
                    "GC-MS分析统计",
                    f"所选时间范围（{time_range}）内暂无GC-MS分析记录。")
                return

            # 一次性统计各分析的化合物数量，避免逐条查询
            analysis_ids = [a.id for a in results]
            counts = dict(
                self.session.query(
                    GCMSCompound.analysis_id,
                    func.count(GCMSCompound.id)
                ).filter(GCMSCompound.analysis_id.in_(analysis_ids))
                 .group_by(GCMSCompound.analysis_id).all()
            )

            # 按供应商统计样品分布
            supplier_counts = {}
            for analysis in results:
                supplier = (analysis.supplier or '').strip() or '未填写供应商'
                supplier_counts[supplier] = supplier_counts.get(supplier, 0) + 1

            chart = QChart()
            chart.setTitle(f"GC-MS分析统计（{time_range}，共 {len(results)} 条）")
            series = QPieSeries()
            for supplier, count in sorted(
                    supplier_counts.items(), key=lambda kv: kv[1], reverse=True):
                slice_ = series.append(f"{supplier} ({count})", count)
                slice_.setLabelVisible(True)
            chart.addSeries(series)
            self.chart_view.setChart(chart)

            # 明细表格
            headers = ["编号", "样品名称", "供应商", "化合物数", "分析时间"]
            self.stat_table.setColumnCount(len(headers))
            self.stat_table.setHorizontalHeaderLabels(headers)
            self.stat_table.setRowCount(len(results))

            for row, analysis in enumerate(results):
                analysis_time = analysis.analysis_time
                values = [
                    analysis.number or '',
                    analysis.name or '',
                    analysis.supplier or '',
                    str(counts.get(analysis.id, 0)),
                    analysis_time.strftime("%Y-%m-%d %H:%M")
                    if analysis_time else "未记录",
                ]
                for col, value in enumerate(values):
                    self.stat_table.setItem(row, col, QTableWidgetItem(str(value)))

        except Exception as e:
            QMessageBox.warning(self, "统计错误", f"GC-MS分析统计时发生错误：{str(e)}")
            self._show_empty_chart("GC-MS分析统计 - 错误", str(e))

    # Removed show_operation_log_stats method

    def show_gcms_compound_stats(self, time_range):
        """显示GC-MS化合物频次统计（跨样品按化合物汇总出现次数 Top15）

        与「GC-MS分析统计」互补：后者看样品/供应商分布，本统计看
        哪些化合物在多个样品中高频出现，辅助判断特征/标志性组分。
        """
        try:
            start_date = self.get_start_date(time_range)

            # 仅统计时间范围内的分析所包含的化合物
            analysis_q = self.session.query(GCMSAnalysis.id)
            if start_date:
                analysis_q = analysis_q.filter(
                    (GCMSAnalysis.analysis_time >= start_date) |
                    (GCMSAnalysis.analysis_time.is_(None))
                )
            analysis_ids = [a[0] for a in analysis_q.all()]
            if not analysis_ids:
                self._show_empty_chart(
                    "GC-MS化合物频次统计",
                    f"所选时间范围（{time_range}）内暂无GC-MS分析记录。")
                return

            rows = self.session.query(
                GCMSCompound.name_cn,
                GCMSCompound.name_en,
                func.count(GCMSCompound.id).label('cnt')
            ).filter(GCMSCompound.analysis_id.in_(analysis_ids)).group_by(
                GCMSCompound.name_cn, GCMSCompound.name_en
            ).order_by(text('cnt DESC')).limit(15).all()

            if not rows:
                self._show_empty_chart(
                    "GC-MS化合物频次统计",
                    f"所选时间范围（{time_range}）内暂无化合物数据。")
                return

            def _cname(cn, en):
                return (cn or en or '').strip() or '未知化合物'

            names, counts = [], []
            for cn, en, cnt in rows:
                names.append(_cname(cn, en))
                counts.append(cnt)

            chart = QChart()
            chart.setTitle(f"GC-MS化合物频次 Top15（{time_range}）")
            series = QBarSeries()
            bar_set = QBarSet("出现样品数")
            for c in counts:
                bar_set.append(c)
            series.append(bar_set)
            chart.addSeries(series)
            axis_x = QBarCategoryAxis()
            axis_x.append(names)
            chart.addAxis(axis_x, Qt.AlignmentFlag.AlignBottom)
            series.attachAxis(axis_x)
            axis_y = QValueAxis()
            max_c = max(counts)
            axis_y.setRange(0, max_c * 1.1 if max_c > 0 else 10)
            axis_y.setTitleText("出现样品数")
            chart.addAxis(axis_y, Qt.AlignmentFlag.AlignLeft)
            series.attachAxis(axis_y)
            self.chart_view.setChart(chart)

            self.stat_table.setColumnCount(3)
            self.stat_table.setHorizontalHeaderLabels(["化合物", "出现样品数", "占比"])
            self.stat_table.setRowCount(len(rows))
            total = sum(counts)
            for row, (cn, en, cnt) in enumerate(rows):
                self.stat_table.setItem(row, 0, QTableWidgetItem(_cname(cn, en)))
                self.stat_table.setItem(row, 1, QTableWidgetItem(str(cnt)))
                rate = (cnt / total * 100) if total else 0
                self.stat_table.setItem(row, 2, QTableWidgetItem(f"{rate:.2f}%"))
        except Exception as e:
            QMessageBox.warning(self, "统计错误", f"GC-MS化合物频次统计时发生错误：{str(e)}")
            self._show_empty_chart("GC-MS化合物频次统计 - 错误", str(e))

    def get_start_date(self, time_range):
        """获取时间范围的开始日期"""
        now = datetime.now()
        if time_range == "最近一周":
            return now - timedelta(days=7)
        elif time_range == "最近一月":
            return now - timedelta(days=30)
        elif time_range == "最近三月":
            return now - timedelta(days=90)
        elif time_range == "最近一年":
            return now - timedelta(days=365)
        return None
        
    def export_statistics(self):
        """导出统计数据"""
        # Removed permission check
            
        file_path, _ = QFileDialog.getSaveFileName(
            self, "导出统计数据", "统计数据.xlsx",
            "Excel 文件 (*.xlsx)"
        )
        
        if not file_path:
            return
            
        try:
            # 创建Excel工作簿
            wb = Workbook()
            
            # 添加数据表格
            ws = wb.active
            ws.title = "统计数据"
            
            # 写入表头
            headers = []
            for col in range(self.stat_table.columnCount()):
                header_item = self.stat_table.horizontalHeaderItem(col)
                if header_item:
                    headers.append(header_item.text())
                else:
                    headers.append("") # Add empty header if item is None
            ws.append(headers)
            
            # 写入数据
            for row in range(self.stat_table.rowCount()):
                row_data = []
                for col in range(self.stat_table.columnCount()):
                    item = self.stat_table.item(row, col)
                    row_data.append(item.text() if item else "")
                ws.append(row_data)
                
            # 保存文件
            wb.save(file_path)
            QMessageBox.information(self, "导出成功", "统计数据已成功导出！")
            
        except Exception as e:
            QMessageBox.critical(self, "导出失败", f"导出统计数据时发生错误：{str(e)}")
    
    def show_ingredients(self):
        self.content.setCurrentIndex(0)
        self.refresh_ingredient_table()
    
    def show_formulas(self):
        self.content.setCurrentIndex(1)
        self.refresh_formula_table()
    
    def show_stock(self):
        self.content.setCurrentIndex(2)
    
    def show_suppliers(self):
        self.content.setCurrentIndex(3)
    
    def show_gcms(self):
        self.content.setCurrentIndex(4)
        self.refresh_gcms_table()
    
    def show_ai_analysis(self):
        """显示智能分析页面"""
        self.content.setCurrentIndex(5)
        # 刷新配方下拉框
        if hasattr(self, 'target_formula_combo') and hasattr(self, 'ai_analysis'):
            self.ai_analysis.load_formulas_to_combo(self.target_formula_combo)

    def show_flavor_expert(self):
        """显示AI调香专家页面"""
        if not hasattr(self, 'flavor_expert'):
            QMessageBox.information(self, "提示", "AI调香专家模块未安装")
            return
        self.content.setCurrentIndex(self.content.indexOf(self.flavor_expert))
        self.flavor_expert.load_formula_combo()
        self.flavor_expert.load_gcms_combo()

    def show_statistics(self):
        """显示数据统计页面"""
        self.content.setCurrentIndex(6)
        # 更新统计数据
        if hasattr(self, 'stat_type'):
            self.update_statistics()

    def import_ingredients(self):
        """批量导入原料数据，支持去重并覆盖（xlsx/xlsm/xls/csv/tsv/txt及伪Excel）"""
        from table_io import read_table_any, SUPPORTED_FILE_FILTER
        file_path, _ = QFileDialog.getOpenFileName(self, "选择原料数据文件", "", SUPPORTED_FILE_FILTER)
        if not file_path:
            return
        try:
            # 先回滚任何未完成的事务
            self.session.rollback()

            df = read_table_any(file_path)
            df.columns = [str(c).strip() for c in df.columns]
            field_map = {
                '编号': 'number',
                'CAS': 'cas_number',
                '原料名称': 'name',
                '英文名': 'english_name',
                '分子式': 'molecular_formula',
                '化学结构式': 'chemical_structure',
                '分子量': 'molecular_weight',
                '沸点': 'boiling_point',
                '溶解性': 'solubility',
                '电子烟国标最大用量': 'max_limit_gb',
                '天然存在': 'natural_occurrence',
                '调香用途': 'perfume_usage',
                '香气香韵': 'aroma_character',
                '香气变调': 'aroma_change',
                '嗅香香气（1% in PG）': 'sniff_aroma',
                '抽吸感官评价': 'sensory_evaluation',
                '香韵构成': 'aroma_composition',
                '成本(元/g)': 'price',
                '成本': 'price',
                '价格': 'price',
            }
            # 检查必填列
            required_columns = ['原料名称']  # 只要求原料名称为必填
            for col in required_columns:
                if col not in df.columns:
                    QMessageBox.warning(self, "表头错误", f"缺少必填列：{col}")
                    return
            
            add_count, update_count, skip_count = 0, 0, 0
            error_messages = []
            
            # 预先收集所有现有的编号，避免重复
            existing_numbers = set()
            all_ingredients = self.session.query(Ingredient).all()
            for ing in all_ingredients:
                if ing.number:
                    existing_numbers.add(ing.number)
            
            for row_idx, row in df.iterrows():
                try:
                    # 处理数据，确保必填字段不为空
                    data = {}
                    for k in field_map:
                        if k in row and pd.notna(row[k]) and str(row[k]).strip():
                            data[field_map[k]] = str(row[k]).strip()

                    # 成本列为数值型
                    if 'price' in data:
                        try:
                            data['price'] = float(
                                str(data['price'])
                                .replace('¥', '').replace('￥', '').strip())
                        except ValueError:
                            data['price'] = 0.0
                    
                    # 确保必填字段存在
                    if 'name' not in data or not data['name']:
                        print(f"跳过第{row_idx+1}行：原料名称为空")
                        skip_count += 1
                        continue
                    
                    # 检查是否已存在（根据CAS或名称）
                    exists = None
                    if data.get('cas_number'):
                        exists = self.session.query(Ingredient).filter_by(cas_number=data['cas_number']).first()
                    if not exists and data.get('name'):
                        exists = self.session.query(Ingredient).filter_by(name=data['name']).first()
                    
                    if exists:
                        # 更新现有原料
                        # 处理编号冲突问题
                        if 'number' in data:
                            new_number = data['number']
                            # 如果新编号与现有编号冲突（且不是当前原料的编号）
                            if new_number in existing_numbers and exists.number != new_number:
                                # 检查是否有其他原料使用这个编号
                                conflict_ingredient = self.session.query(Ingredient).filter(
                                    Ingredient.number == new_number,
                                    Ingredient.id != exists.id
                                ).first()
                                
                                if conflict_ingredient:
                                    # 生成新的唯一编号
                                    import time
                                    unique_number = f"{new_number}_IMPORT_{int(time.time())}_{row_idx}"
                                    data['number'] = unique_number
                                    existing_numbers.add(unique_number)
                                    error_messages.append(f"第{row_idx+1}行：编号'{new_number}'已存在，自动修改为'{unique_number}'")
                                else:
                                    existing_numbers.add(new_number)
                            elif new_number not in existing_numbers:
                                existing_numbers.add(new_number)
                        
                        # 更新所有字段
                        for key, value in data.items():
                            setattr(exists, key, value)
                        update_count += 1
                    else:
                        # 新增原料
                        # 如果没有编号，自动生成一个
                        if 'number' not in data or not data['number']:
                            import time
                            unique_number = f"ING_{int(time.time())}_{row_idx}"
                            data['number'] = unique_number
                            existing_numbers.add(unique_number)
                        else:
                            # 检查编号是否冲突
                            new_number = data['number']
                            if new_number in existing_numbers:
                                # 生成新的唯一编号
                                import time
                                unique_number = f"{new_number}_IMPORT_{int(time.time())}_{row_idx}"
                                data['number'] = unique_number
                                existing_numbers.add(unique_number)
                                error_messages.append(f"第{row_idx+1}行：编号'{new_number}'已存在，自动修改为'{unique_number}'")
                            else:
                                existing_numbers.add(new_number)
                        
                        ing = Ingredient(**data)
                        self.session.add(ing)
                        add_count += 1
                        
                except Exception as e:
                    error_messages.append(f"第{row_idx+1}行处理失败: {str(e)}")
                    skip_count += 1
                    continue
            
            # 提交事务
            self.session.commit()
            self.refresh_ingredient_table()
            
            # 显示导入结果
            result_message = f"导入完成！\n成功导入 {add_count} 条新原料\n更新 {update_count} 条原料\n跳过 {skip_count} 条记录"
            
            if error_messages:
                result_message += f"\n\n注意事项：\n" + "\n".join(error_messages[:10])  # 只显示前10个错误
                if len(error_messages) > 10:
                    result_message += f"\n... 还有 {len(error_messages) - 10} 个问题"
            
            QMessageBox.information(self, "导入完成", result_message)
            
        except Exception as e:
            self.session.rollback()
            QMessageBox.critical(self, "导入失败", f"导入过程中发生错误：{str(e)}")

    def export_ingredients(self):
        """导出当前筛选后的原料信息为Excel或CSV文件"""
        from PyQt6.QtWidgets import QFileDialog
        
        # 检查pandas可用性
        if not PANDAS_AVAILABLE:
            QMessageBox.warning(self, "功能受限", 
                              "pandas不可用，无法导出Excel/CSV文件。\n\n"
                              "建议：\n"
                              "1. 安装pandas: pip install pandas\n"
                              "2. 或手动复制表格数据\n"
                              "3. 或查看'依赖包问题解决方案.md'")
            return
            
        # Get the current filtered data
        keyword = self.ingredient_search.text().strip() if hasattr(self, 'ingredient_search') else ''
        query = self.session.query(Ingredient)
        if keyword:
            like = f"%{keyword}%"
            query = query.filter(
                Ingredient.name.like(like) |
                Ingredient.cas_number.like(like) |
                Ingredient.english_name.like(like) |
                Ingredient.molecular_formula.like(like)
            )
        ingredients = query.all()
        if not ingredients:
            QMessageBox.information(self, "导出", "没有可导出的原料数据！")
            return
        # Assemble export data
        data = []
        columns = [
            '编号', 'CAS', '原料名称', '英文名', '分子式', '化学结构式', '分子量', '沸点', '溶解性', '电子烟国标最大用量',
            '天然存在', '调香用途', '香气香韵', '香气变调', '嗅香香气（1% in PG）', '抽吸感官评价', '香韵构成'
        ]
        for ing in ingredients:
            data.append([
                ing.number, ing.cas_number, ing.name, ing.english_name, ing.molecular_formula,
                ing.chemical_structure, ing.molecular_weight, ing.boiling_point, ing.solubility,
                ing.max_limit_gb, ing.natural_occurrence, ing.perfume_usage, ing.aroma_character,
                ing.aroma_change, ing.sniff_aroma, ing.sensory_evaluation, ing.aroma_composition
            ])
        df = pd.DataFrame(data, columns=columns)
        # Select export filename
        file_path, file_type = QFileDialog.getSaveFileName(self, "导出原料信息", "原料信息.xlsx", "Excel 文件 (*.xlsx);;CSV 文件 (*.csv)")
        if not file_path:
            return
        try:
            if file_path.endswith('.csv'):
                df.to_csv(file_path, index=False, encoding='utf-8-sig')
            else:
                df.to_excel(file_path, index=False)
            QMessageBox.information(self, "导出成功", f"原料信息已导出到：\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", f"导出过程中发生错误：{str(e)}")


    def refresh_ingredient_table(self):
        # 如果存在旧的加载器，先停止它
        if self.ingredient_loader and self.ingredient_loader.isRunning():
            self.ingredient_loader.stop()
            
        keyword = self.ingredient_search.text().strip() if hasattr(self, 'ingredient_search') else ''
        
        # 计算总数和分页信息
        from models import Ingredient
        session = self.session
        query = session.query(Ingredient)
        if keyword:
            like = f"%{keyword}%"
            query = query.filter(
                Ingredient.name.like(like) |
                Ingredient.cas_number.like(like) |
                Ingredient.english_name.like(like) |
                Ingredient.molecular_formula.like(like)
            )
        total = query.count()
        max_page = max(1, (total + self.ingredient_page_size - 1) // self.ingredient_page_size)
        self.ingredient_page = min(self.ingredient_page, max_page)
        offset = (self.ingredient_page - 1) * self.ingredient_page_size
        self.ingredient_page_label.setText(f"第{self.ingredient_page}/{max_page}页 共{total}条")
        # Ensure buttons exist before enabling/disabling
        if hasattr(self, 'ingredient_prev_btn') and hasattr(self, 'ingredient_next_btn'):
             self.ingredient_prev_btn.setEnabled(self.ingredient_page > 1)
             self.ingredient_next_btn.setEnabled(self.ingredient_page < max_page)
        
        self.ingredient_table.setRowCount(0)
        
        # Create new loader with enhanced filter logic and sorting
        self.ingredient_loader = PagedTableLoader(
            self.engine, 
            Ingredient, 
            keyword, 
            offset, 
            self.ingredient_page_size,
            self.ingredient_sort_column,
            self.ingredient_sort_order
        )
        self.ingredient_loader.data_loaded.connect(self._fill_ingredient_table)
        self.ingredient_loader.start()

    def handle_ingredient_sort(self, logical_index):
        """处理原料表格头部点击排序事件"""
        # 定义列索引到数据库字段的映射
        column_mapping = {
            0: 'number',           # 编号
            1: 'cas_number',       # CAS
            2: 'name',            # 原料名称
            3: 'english_name',    # 英文名
            4: 'molecular_formula', # 分子式
            5: 'chemical_structure', # 化学结构式
            6: 'molecular_weight', # 分子量
            7: 'boiling_point',   # 沸点
            8: 'solubility',      # 溶解性
            9: 'max_limit_gb',    # 电子烟国标最大用量
            10: 'natural_occurrence', # 天然存在
            11: 'perfume_usage',  # 调香用途
            12: 'aroma_character', # 香气香韵
            13: 'aroma_change',   # 香气变调
            14: 'sniff_aroma',    # 嗅香香气
            15: 'sensory_evaluation', # 抽吸感官评价
            16: 'aroma_composition', # 香韵构成
            17: 'price',          # 成本(元/g)
            18: None              # 操作列不支持排序
        }
        
        # 操作列不支持排序
        if logical_index == 18 or logical_index not in column_mapping:
            return
            
        column_field = column_mapping[logical_index]
        if not column_field:
            return
            
        # 更新排序状态
        if self.ingredient_sort_column == column_field:
            # 同一列，切换排序方向
            self.ingredient_sort_order = 'desc' if self.ingredient_sort_order == 'asc' else 'asc'
        else:
            # 不同列，重置为升序
            self.ingredient_sort_column = column_field
            self.ingredient_sort_order = 'asc'
            
        # 重置到第一页
        self.ingredient_page = 1
        
        # 更新表格头部排序指示器
        self.update_ingredient_sort_indicator()
        
        # 刷新表格数据
        self.refresh_ingredient_table()
    
    def update_ingredient_sort_indicator(self):
        """更新原料表格头部的排序指示器"""
        # 定义字段到列索引的反向映射
        field_to_column = {
            'number': 0, 'cas_number': 1, 'name': 2, 'english_name': 3,
            'molecular_formula': 4, 'chemical_structure': 5, 'molecular_weight': 6,
            'boiling_point': 7, 'solubility': 8, 'max_limit_gb': 9,
            'natural_occurrence': 10, 'perfume_usage': 11, 'aroma_character': 12,
            'aroma_change': 13, 'sniff_aroma': 14, 'sensory_evaluation': 15,
            'aroma_composition': 16, 'price': 17
        }
        
        header = self.ingredient_table.horizontalHeader()
        
        # 清除所有列的排序指示器
        for i in range(header.count() - 1):  # 排除操作列
            header.setSortIndicatorShown(False)
            
        # 设置当前排序列的指示器
        if self.ingredient_sort_column and self.ingredient_sort_column in field_to_column:
            column_index = field_to_column[self.ingredient_sort_column]
            sort_order = Qt.SortOrder.AscendingOrder if self.ingredient_sort_order == 'asc' else Qt.SortOrder.DescendingOrder
            header.setSortIndicator(column_index, sort_order)
            header.setSortIndicatorShown(True)

    def _fill_ingredient_table(self, ingredients):
        self.ingredient_table.setRowCount(len(ingredients))
        # self.ingredient_table.setColumnCount(18) # Already set in create_ingredients_page
        self.ingredient_table.setWordWrap(True)
        
        # 临时禁用排序，提高性能
        self.ingredient_table.setSortingEnabled(False)
        
        for row_idx, ing in enumerate(ingredients):
            values = [
                ing.number, ing.cas_number, ing.name, ing.english_name, ing.molecular_formula,
                ing.chemical_structure, ing.molecular_weight, ing.boiling_point, ing.solubility,
                ing.max_limit_gb, ing.natural_occurrence, ing.perfume_usage, ing.aroma_character,
                ing.aroma_change, ing.sniff_aroma, ing.sensory_evaluation, ing.aroma_composition,
                getattr(ing, 'price', 0), # 新增成本列
            ]
            for col_idx, value in enumerate(values):
                if col_idx == 5 and value and isinstance(value, str) and value.startswith("https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/"):
                    label = QLabel()
                    label.setAlignment(Qt.AlignmentFlag.AlignCenter)
                    label.setText("点击加载图片")
                    def load_img(event, url=value, label=label):
                        try:
                            resp = requests.get(url, timeout=5)
                            if resp.status_code == 200:
                                pix = QPixmap()
                                pix.loadFromData(resp.content)
                                label.setPixmap(pix.scaledToHeight(60))
                                label.setToolTip("点击放大")
                                label.mousePressEvent = lambda e, img=pix: self.show_image_dialog(img, "化学结构式")
                            else:
                                label.setText("图片加载失败")
                        except Exception:
                            label.setText("图片加载失败")
                    label.mousePressEvent = load_img
                    self.ingredient_table.setCellWidget(row_idx, col_idx, label)
                elif col_idx == 16 and value and isinstance(value, str) and ':' in value:
                    label = QLabel()
                    label.setAlignment(Qt.AlignmentFlag.AlignCenter)
                    label.setText("点击加载饼图")
                    def load_pie(event, value=value, label=label):
                        try:
                            parts = [p.strip() for p in value.split(',') if ':' in p]
                            labels = []
                            sizes = []
                            for p in parts:
                                n, v = p.split(':', 1)
                                labels.append(n.strip())
                                v = v.replace('%','').strip()
                                try:
                                    sizes.append(float(v))
                                except:
                                    sizes.append(0)
                            if sum(sizes) == 0:
                                raise ValueError
                            fig, ax = plt.subplots(figsize=(3,3), dpi=150)
                            ax.pie(sizes, labels=labels, autopct='%1.0f%%', textprops={'fontsize': 10})
                            ax.axis('equal')
                            plt.tight_layout(pad=0.1)
                            buf = BytesIO()
                            plt.savefig(buf, format='png', bbox_inches='tight', pad_inches=0.05)
                            plt.close(fig)
                            buf.seek(0)
                            pix = QPixmap()
                            pix.loadFromData(buf.read())
                            label.setPixmap(pix.scaled(60,60))
                            label.setToolTip("点击放大")
                            label.mousePressEvent = lambda e, img=pix: self.show_image_dialog(img, "香韵构成饼图")
                        except Exception:
                            label.setText("饼图错误")
                    label.mousePressEvent = load_pie
                    self.ingredient_table.setCellWidget(row_idx, col_idx, label)
                elif col_idx in [0, 1, 6, 17]:  # 编号、CAS、分子量、成本列使用数字排序
                    item = NumericTableWidgetItem(str(value) if value is not None else "")
                    if col_idx == 0:  # 编号列
                        # 提取编号中的数字部分用于排序
                        try:
                            # 从类似 "0002-ESY-RH002-V1" 中提取数字
                            number_part = str(value).split('-')[0] if value else "0"
                            item.setData(Qt.ItemDataRole.UserRole, int(number_part))
                        except:
                            item.setData(Qt.ItemDataRole.UserRole, 0)
                    elif col_idx == 1:  # CAS列
                        # 提取CAS号中的第一个数字用于排序
                        try:
                            cas_parts = str(value).split('-') if value else ["0"]
                            item.setData(Qt.ItemDataRole.UserRole, int(cas_parts[0]))
                        except:
                            item.setData(Qt.ItemDataRole.UserRole, 0)
                    elif col_idx == 6:  # 分子量列
                        try:
                            # 提取数字部分
                            import re
                            numbers = re.findall(r'\d+\.?\d*', str(value)) if value else ["0"]
                            item.setData(Qt.ItemDataRole.UserRole, float(numbers[0]) if numbers else 0.0)
                        except:
                            item.setData(Qt.ItemDataRole.UserRole, 0.0)
                    elif col_idx == 17:  # 成本列
                        try:
                            item.setData(Qt.ItemDataRole.UserRole, float(value) if value else 0.0)
                        except:
                            item.setData(Qt.ItemDataRole.UserRole, 0.0)
                    
                    item.setToolTip(str(value) if value is not None else "")
                    item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
                    item.setFlags(item.flags() ^ Qt.ItemFlag.ItemIsEditable)
                    self.ingredient_table.setItem(row_idx, col_idx, item)
                else:
                    item = QTableWidgetItem(str(value) if value is not None else "")
                    item.setToolTip(str(value) if value is not None else "")
                    item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
                    item.setFlags(item.flags() ^ Qt.ItemFlag.ItemIsEditable)
                    self.ingredient_table.setItem(row_idx, col_idx, item)
            
            # Add Edit and Delete buttons for each row
            edit_btn = QPushButton("编辑")
            del_btn = QPushButton("删除")
            
            # Connect buttons with row data
            edit_btn.clicked.connect(lambda _, i=ingredients[row_idx]: self.edit_ingredient_dialog(i))
            del_btn.clicked.connect(lambda _, i=ingredients[row_idx]: self.delete_ingredient(i))
            
            btn_widget = QWidget()
            btn_layout = QHBoxLayout(btn_widget)
            btn_layout.setContentsMargins(0,0,0,0)
            btn_layout.addWidget(edit_btn)
            btn_layout.addWidget(del_btn)
            btn_widget.setLayout(btn_layout)
            
            # Set the button widget in the last column
            self.ingredient_table.setCellWidget(row_idx, 18, btn_widget)

        # 注意：我们使用数据库级别排序，不需要重新启用QTableWidget的内置排序

        if not hasattr(self, '_ingredient_table_resized'):
            self.ingredient_table.resizeColumnsToContents()
            self.ingredient_table.resizeRowsToContents()
            self._ingredient_table_resized = True

    def ingredient_prev_page(self):
        if self.ingredient_page > 1:
            self.ingredient_page -= 1
            self.refresh_ingredient_table()

    def ingredient_next_page(self):
        self.ingredient_page += 1
        self.refresh_ingredient_table()

    def add_ingredient_dialog(self):
        dlg = IngredientDialog(self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            data = dlg.get_data()
            ing = Ingredient(**data)
            self.session.add(ing)
            self.session.commit()
            self.refresh_ingredient_table()
            QMessageBox.information(self, "添加成功", "原料已添加！")

    def edit_ingredient_dialog(self, ingredient):
        print(f"编辑前的原料信息: ID={ingredient.id}, 名称={ingredient.name}, CAS={ingredient.cas_number}")
        dlg = IngredientDialog(self, ingredient)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            data = dlg.get_data()
            print(f"从对话框获取的数据: {data}")
            
            try:
                # 重新从数据库查询这个对象，确保获取最新状态
                fresh_ingredient = self.session.query(Ingredient).filter_by(id=ingredient.id).first()
                if not fresh_ingredient:
                    QMessageBox.critical(self, "错误", "找不到要编辑的原料！")
                    return
                
                # 更新数据，处理数据类型转换
                for key, value in data.items():
                    if key == 'price':  # 价格字段需要转换为浮点数
                        try:
                            setattr(fresh_ingredient, key, float(value) if value else 0.0)
                        except ValueError:
                            setattr(fresh_ingredient, key, 0.0)
                    else:
                        setattr(fresh_ingredient, key, value)
                
                # 手动更新 updated_at 字段
                from datetime import datetime
                fresh_ingredient.updated_at = datetime.now()
                
                print(f"更新后的原料信息: {fresh_ingredient.name}, CAS={fresh_ingredient.cas_number}")
                
                # 提交事务
                self.session.commit()
                print("数据库提交成功")
                
                # 刷新表格
                self.refresh_ingredient_table()
                print("表格刷新完成")
                
                QMessageBox.information(self, "修改成功", "原料信息已更新！")
                
            except Exception as e:
                print(f"数据库操作失败: {e}")
                self.session.rollback()
                QMessageBox.critical(self, "修改失败", f"修改原料信息时发生错误：{str(e)}")

    def delete_ingredient(self, ingredient):
        """删除原料（先提示关联影响，异常时回滚）"""
        try:
            # 统计关联影响：被多少配方引用、有多少库存记录
            formula_count = self.session.query(
                func.count(ingredient_formula.c.formula_id)
            ).filter(
                ingredient_formula.c.ingredient_id == ingredient.id
            ).scalar() or 0
            stock_count = self.session.query(StockRecord).filter_by(
                ingredient_id=ingredient.id, is_deleted=False).count()

            msg = f"确定要删除原料：{ingredient.name}（CAS: {ingredient.cas_number}）吗？"
            tips = []
            if formula_count:
                tips.append(f"该原料被 {formula_count} 个配方引用，删除后会从这些配方中移除")
            if stock_count:
                tips.append(f"存在 {stock_count} 条关联库存记录，删除后这些记录将失去原料关联")
            if tips:
                msg += "\n\n注意：\n· " + "\n· ".join(tips)

            reply = QMessageBox.question(
                self, "确认删除", msg,
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply != QMessageBox.StandardButton.Yes:
                return

            self.session.delete(ingredient)
            self.session.commit()
            self.refresh_ingredient_table()
            QMessageBox.information(self, "删除成功", "原料已删除！")
        except Exception as e:
            self.session.rollback()
            QMessageBox.critical(self, "删除失败", f"删除原料时发生错误：{str(e)}")

    def show_image_dialog(self, pixmap, title="图片预览"):
        dlg = QDialog(self)
        dlg.setWindowTitle(title)
        dlg.setWindowFlags(dlg.windowFlags() | Qt.WindowType.WindowMinimizeButtonHint | Qt.WindowType.WindowMaximizeButtonHint)
        vbox = QVBoxLayout()
        label = QLabel()
        label.setPixmap(pixmap.scaled(600, 600, Qt.AspectRatioMode.KeepAspectRatio))
        vbox.addWidget(label)
        dlg.setLayout(vbox)
        dlg.resize(620, 620)
        dlg.exec()

    def refresh_formula_table(self):
        # 如果存在旧的加载器，先停止它
        if self.formula_loader and self.formula_loader.isRunning():
            self.formula_loader.stop()
            
        keyword = self.formula_search.text().strip() if hasattr(self, 'formula_search') else ''
        query = self.session.query(Formula)
        if keyword:
            like = f"%{keyword}%"
            query = query.filter(
                Formula.name.like(like) |
                Formula.description.like(like) |
                Formula.creator.like(like)
            )
        total = query.count()
        max_page = max(1, (total + self.formula_page_size - 1) // self.formula_page_size)
        self.formula_page = min(self.formula_page, max_page)
        offset = (self.formula_page - 1) * self.formula_page_size
        self.formula_page_label.setText(f"第{self.formula_page}/{max_page}页 共{total}条")

        # Ensure buttons exist before enabling/disabling
        if hasattr(self, 'formula_prev_btn') and hasattr(self, 'formula_next_btn'):
            self.formula_prev_btn.setEnabled(self.formula_page > 1)
            self.formula_next_btn.setEnabled(self.formula_page < max_page)
        
        self.formula_table.setRowCount(0)
        
        # Create new loader with sorting support
        self.formula_loader = PagedTableLoader(
            self.engine, 
            Formula, 
            keyword, 
            offset, 
            self.formula_page_size,
            self.formula_sort_column,
            self.formula_sort_order
        )
        self.formula_loader.data_loaded.connect(self._fill_formula_table)
        self.formula_loader.start()

    def handle_formula_sort(self, logical_index):
        """处理配方表格头部点击排序事件"""
        # 定义列索引到数据库字段的映射
        column_mapping = {
            0: 'number',          # 编号
            1: 'name',           # 名称
            2: 'creator',        # 创建人
            3: 'created_at',     # 创建时间
            4: 'updated_at',     # 修改时间
            5: 'evaluation',     # 感官评价
            6: 'total_cost',     # 总成本(元)
            7: None              # 操作列不支持排序
        }
        
        # 操作列不支持排序
        if logical_index == 7 or logical_index not in column_mapping:
            return
            
        column_field = column_mapping[logical_index]
        if not column_field:
            return
            
        # 更新排序状态
        if self.formula_sort_column == column_field:
            # 同一列，切换排序方向
            self.formula_sort_order = 'desc' if self.formula_sort_order == 'asc' else 'asc'
        else:
            # 不同列，重置为升序
            self.formula_sort_column = column_field
            self.formula_sort_order = 'asc'
            
        # 重置到第一页
        self.formula_page = 1
        
        # 更新表格头部排序指示器
        self.update_formula_sort_indicator()
        
        # 刷新表格数据
        self.refresh_formula_table()
    
    def update_formula_sort_indicator(self):
        """更新配方表格头部的排序指示器"""
        # 定义字段到列索引的反向映射
        field_to_column = {
            'number': 0, 'name': 1, 'creator': 2, 'created_at': 3,
            'updated_at': 4, 'evaluation': 5, 'total_cost': 6
        }
        
        header = self.formula_table.horizontalHeader()
        
        # 清除所有列的排序指示器
        for i in range(header.count() - 1):  # 排除操作列
            header.setSortIndicatorShown(False)
            
        # 设置当前排序列的指示器
        if self.formula_sort_column and self.formula_sort_column in field_to_column:
            column_index = field_to_column[self.formula_sort_column]
            sort_order = Qt.SortOrder.AscendingOrder if self.formula_sort_order == 'asc' else Qt.SortOrder.DescendingOrder
            header.setSortIndicator(column_index, sort_order)
            header.setSortIndicatorShown(True)

    def _fill_formula_table(self, formulas):
        self.formula_table.setRowCount(len(formulas))
        # self.formula_table.setColumnCount(7) # Already set in create_formulas_page
        for row, formula in enumerate(formulas):
            # Updated columns based on create_formulas_page
            for col, value in enumerate([
                formula.number, formula.name, formula.creator, formula.created_at,
                formula.updated_at, formula.evaluation, formula.total_cost
            ]):
                if col == 6:  # 总成本(元)列，数字排序
                    item = NumericTableWidgetItem(str(value) if value is not None else "")
                    try:
                        item.setData(Qt.ItemDataRole.UserRole, float(value))
                    except Exception:
                        item.setData(Qt.ItemDataRole.UserRole, -1.0)
                    item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
                    item.setFlags(item.flags() ^ Qt.ItemFlag.ItemIsEditable)
                    self.formula_table.setItem(row, col, item)
                else:
                    item = QTableWidgetItem(str(value) if value is not None else "")
                    item.setToolTip(str(value) if value is not None else "")
                    item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
                    item.setFlags(item.flags() ^ Qt.ItemFlag.ItemIsEditable)
                    self.formula_table.setItem(row, col, item)

            # Add Edit and Delete buttons
            btn_widget = QWidget()
            btn_layout = QHBoxLayout()
            btn_layout.setContentsMargins(0, 0, 0, 0)
            edit_btn = QPushButton("编辑")
            del_btn = QPushButton("删除")
            detail_btn = QPushButton("详情") # Add Detail button here

            edit_btn.clicked.connect(lambda checked, f=formulas[row]: self.edit_formula_dialog(f))
            del_btn.clicked.connect(lambda checked, f=formulas[row]: self.delete_formula(f))
            detail_btn.clicked.connect(lambda checked, f=formulas[row]: self.show_formula_detail(f)) # Connect Detail button

            btn_layout.addWidget(edit_btn)
            btn_layout.addWidget(del_btn)
            btn_layout.addWidget(detail_btn) # Add Detail button to the layout
            btn_widget.setLayout(btn_layout)

            # Set buttons in the '操作' column (index 7 based on new headers)
            self.formula_table.setCellWidget(row, 7, btn_widget)

        if not hasattr(self, '_formula_table_resized'):
            self.formula_table.resizeColumnsToContents()
            self.formula_table.resizeRowsToContents()
            self._formula_table_resized = True


    def formula_prev_page(self):
        if self.formula_page > 1:
            self.formula_page -= 1
            self.refresh_formula_table()

    def formula_next_page(self):
        self.formula_page += 1
        self.refresh_formula_table()

    def add_formula_dialog(self):
        dlg = FormulaDialog(self, self.session)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            data = dlg.get_data()
            formula = Formula(**data)
            self.session.add(formula)
            self.session.commit()
            self.refresh_formula_table()
            QMessageBox.information(self, "添加成功", "配方已添加！")

    def edit_formula_dialog(self, formula):
        formula_data = {
            'number': formula.number,
            'name': formula.name,
            'creator': formula.creator,
            'evaluation': formula.evaluation,
            'description': formula.description,  # 添加描述字段
            'content': formula.content,
            'created_at': formula.created_at,
            'updated_at': formula.updated_at
        }
        # Pass the formula object itself to the dialog for potential updates
        dialog = FormulaDialog(self, self.session, formula_data, existing_formula=formula)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            new_data = dialog.get_data()
            # Update the existing formula object
            formula.number = new_data['number']
            formula.name = new_data['name']
            formula.creator = new_data['creator']
            formula.evaluation = new_data['evaluation']
            formula.content = new_data['content']
            formula.total_cost = new_data['total_cost'] # Update total cost
            formula.updated_at = datetime.now() # Update modified time

            # Update ingredient associations and percentages
            # Clear existing associations
            self.session.query(ingredient_formula).filter_by(formula_id=formula.id).delete()

            # Add new associations from the dialog data
            for ingredient_name_percentage_amount in new_data['content'].split(','):
                if ':' in ingredient_name_percentage_amount:
                    parts = ingredient_name_percentage_amount.split(':')
                    ingredient_name = parts[0].strip()
                    percentage_str = parts[1].strip().replace('%', '')
                    amount_str = parts[2].strip().replace('g', '') if len(parts) > 2 else "0"

                    ingredient = self.session.query(Ingredient).filter_by(name=ingredient_name).first()
                    if ingredient:
                        try:
                            percentage = float(percentage_str)
                            # Create a new entry in the association table
                            insert_stmt = ingredient_formula.insert().values(
                                formula_id=formula.id,
                                ingredient_id=ingredient.id,
                                percentage=percentage
                            )
                            self.session.execute(insert_stmt)
                        except ValueError:
                            print(f"Could not convert percentage to float: {percentage_str}")


            self.session.commit()
            self.refresh_formula_table()
            QMessageBox.information(self, "修改成功", "配方信息已更新！")


    def delete_formula(self, formula):
        """删除配方"""
        # 表格行里的 formula 来自分页线程中已关闭的会话（游离实例），
        # 在主会话中按 id 重新取回，保证删除操作落在正确的会话上并真正落库。
        formula = self.session.query(Formula).filter_by(id=formula.id).first()
        if formula is None:
            QMessageBox.information(self, "提示", "该配方已被删除或不存在。")
            return
        # 检查是否有相关的配方使用记录
        usage_count = self.session.query(FormulaUsage).filter_by(formula_id=formula.id).count()
        
        warning_message = f"确定要删除配方：{formula.name} 吗？"
        if usage_count > 0:
            warning_message += f"\n\n注意：该配方有 {usage_count} 条使用记录，删除配方将同时删除这些记录及相关的库存记录。"
        
        reply = QMessageBox.question(
            self, "确认删除",
            warning_message,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            try:
                # 1. 首先删除相关的库存记录（通过FormulaUsage关联）
                usage_records = self.session.query(FormulaUsage).filter_by(formula_id=formula.id).all()
                for usage in usage_records:
                    # 删除与该使用记录关联的库存记录
                    stock_records = self.session.query(StockRecord).filter_by(formula_usage_id=usage.id).all()
                    for stock_record in stock_records:
                        self.session.delete(stock_record)
                    
                    # 删除配方使用记录
                    self.session.delete(usage)
                
                # 2. 删除配方本身
                self.session.delete(formula)
                
                # 3. 提交所有更改
                self.session.commit()
                
                self.refresh_formula_table()
                
                if usage_count > 0:
                    QMessageBox.information(
                        self, "删除成功", 
                        f"配方已删除！\n同时删除了 {usage_count} 条使用记录及相关库存记录。"
                    )
                else:
                    QMessageBox.information(self, "删除成功", "配方已删除！")
                    
            except Exception as e:
                self.session.rollback()
                QMessageBox.critical(
                    self, "删除失败", 
                    f"删除配方时发生错误：{str(e)}"
                )

    def export_formulas(self):
        """导出当前筛选后的配方信息为CSV文件"""
        from PyQt6.QtWidgets import QFileDialog
        import csv
        from datetime import datetime
        
        # Get the current filtered data
        keyword = self.formula_search.text().strip() if hasattr(self, 'formula_search') else ''
        query = self.session.query(Formula)
        if keyword:
            like = f"%{keyword}%"
            query = query.filter(
                Formula.name.like(like) |
                Formula.description.like(like) |
                Formula.creator.like(like)
            )
        formulas = query.all()
        if not formulas:
            QMessageBox.information(self, "导出", "没有可导出的配方数据！")
            return
            
        # Select export filename (只支持CSV，因为没有pandas)
        file_path, file_type = QFileDialog.getSaveFileName(
            self, 
            "导出配方信息", 
            "配方信息.csv", 
            "CSV 文件 (*.csv)"
        )
        if not file_path:
            return
            
        try:
            with open(file_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
                fieldnames = [
                    '编号', '配方名称', '创建时间', '修改时间', '创建人', 
                    '描述', '配方内容', '感官评价', '总成本'
                ]
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                
                for formula in formulas:
                    # 格式化日期时间
                    created_at = ''
                    if formula.created_at:
                        created_at = formula.created_at.strftime('%Y-%m-%d %H:%M:%S')
                    
                    updated_at = ''
                    if formula.updated_at:
                        updated_at = formula.updated_at.strftime('%Y-%m-%d %H:%M:%S')
                    
                    writer.writerow({
                        '编号': formula.number or '',
                        '配方名称': formula.name or '',
                        '创建时间': created_at,
                        '修改时间': updated_at,
                        '创建人': formula.creator or '',
                        '描述': formula.description or '',
                        '配方内容': formula.content or '',
                        '感官评价': formula.evaluation or '',
                        '总成本': formula.total_cost or 0
                    })
            
            QMessageBox.information(self, "导出成功", f"配方信息已导出到：\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", f"导出过程中发生错误：{str(e)}")


    def show_formula_detail(self, formula):
        """显示配方详情"""
        if formula:
            detail_dialog = FormulaDetailDialog(formula, self.session, self)
            detail_dialog.exec()
        else:
            QMessageBox.warning(self, "警告", "无法找到配方详情")


    def refresh_gcms_table(self):
        """刷新GCMS分析表格"""
        self.gcms_table.setRowCount(0)
        offset = (self.gcms_page - 1) * self.gcms_page_size
        query = self.session.query(GCMSAnalysis)
        total = query.count()
        analyses = query.order_by(GCMSAnalysis.number).offset(offset).limit(self.gcms_page_size).all()
        for row, analysis in enumerate(analyses):
            self.gcms_table.insertRow(row)
            self.gcms_table.setItem(row, 0, QTableWidgetItem(analysis.number or ""))
            self.gcms_table.setItem(row, 1, QTableWidgetItem(analysis.name or ""))
            self.gcms_table.setItem(row, 2, QTableWidgetItem(analysis.supplier or ""))
            self.gcms_table.setItem(row, 3, QTableWidgetItem(analysis.instrument_params or ""))
            # 整体调香思路
            perfume_idea_textedit = QTextEdit()
            perfume_idea_textedit.setPlainText(analysis.perfume_idea or "")
            perfume_idea_textedit.setReadOnly(True)
            perfume_idea_textedit.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
            self.gcms_table.setCellWidget(row, 4, perfume_idea_textedit)
            self.adjust_gcms_row_height(row, perfume_idea_textedit)
            # 谱图列
            if analysis.spectrum_image:
                label = QLabel()
                label.setAlignment(Qt.AlignmentFlag.AlignCenter)
                label.setText("加载图片")
                def load_img(event, path=analysis.spectrum_image, label=label):
                    from PyQt6.QtGui import QPixmap
                    import os
                    if os.path.exists(path):
                        pix = QPixmap(path)
                        label.setPixmap(pix.scaledToHeight(60))
                        label.setToolTip("点击放大")
                        label.mousePressEvent = lambda e, img=pix: self.show_image_dialog(img, "谱图")
                    else:
                        label.setText("图片不存在")
                label.mousePressEvent = load_img
                self.gcms_table.setCellWidget(row, 5, label)
            else:
                self.gcms_table.setCellWidget(row, 5, QLabel("无谱图"))
            # 操作按钮
            btn_widget = QWidget()
            btn_layout = QHBoxLayout(btn_widget)
            btn_layout.setContentsMargins(0, 0, 0, 0)
            btn_layout.setSpacing(2)
            edit_btn = QPushButton("编辑")
            edit_btn.setFixedSize(60, 25)
            edit_btn.clicked.connect(lambda checked, a=analysis: self.edit_gcms_dialog(a))
            delete_btn = QPushButton("删除")
            delete_btn.setFixedSize(60, 25)
            delete_btn.clicked.connect(lambda checked, a=analysis: self.delete_gcms(a))
            btn_layout.addWidget(edit_btn)
            btn_layout.addWidget(delete_btn)
            self.gcms_table.setCellWidget(row, 6, btn_widget)
            self.adjust_gcms_row_height(row, self.gcms_table.item(row, 4))
        self.gcms_page_label.setText(f"第 {self.gcms_page} 页，共 {max(1, (total + self.gcms_page_size - 1) // self.gcms_page_size)} 页")

    def adjust_gcms_row_height(self, row, text_edit):
        """Adjust row height based on QTextEdit content, with a maximum height"""
        # Check if the item is a QTextEdit widget
        if not isinstance(text_edit, QTextEdit):
            # If not a QTextEdit, cannot get document size, skip adjustment or set a default height
            # print(f"Warning: Expected QTextEdit for row height adjustment, but got {type(text_edit)}")
            # Optionally set a default height if needed for consistency
            # default_height = self.gcms_table.verticalHeader().defaultSectionSize()
            # self.gcms_table.setRowHeight(row, default_height)
            return # Skip adjustment if not a QTextEdit

        # Calculate the ideal height of the document
        doc_height = text_edit.document().size().height()
        # Define a maximum allowed row height (adjust this value as needed)
        max_height = 150 # Example maximum height
        # Calculate the new height, ensuring it doesn't exceed the maximum
        new_height = min(int(doc_height + 2), max_height) # Add a small margin and cap at max_height
        # Ensure the row height is at least the default height
        default_height = self.gcms_table.verticalHeader().defaultSectionSize()
        final_height = max(new_height, default_height)
        self.gcms_table.setRowHeight(row, final_height) # Directly set the calculated height


    def gcms_prev_page(self):
        """GCMS分析表格上一页"""
        if self.gcms_page > 1:
            self.gcms_page -= 1
            self.refresh_gcms_table()


    def gcms_next_page(self):
        """GCMS分析表格下一页"""
        max_page = max(1, (self.session.query(GCMSAnalysis).count() + self.gcms_page_size - 1) // self.gcms_page_size)
        if self.gcms_page < max_page:
             self.gcms_page += 1
             self.refresh_gcms_table()


    def add_gcms_dialog(self):
        dlg = GCMSAnalysisDialog(self, self.session)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            data = dlg.get_data()
            print(f"从对话框获取的数据: {data}") # 添加打印语句
            
            # 从 data 中分离化合物数据
            compounds_data = data.pop('compounds', []) # 移除 compounds 键并获取其值，默认为空列表

            try:
                # 使用剩余的数据创建 GCMSAnalysis 实例
                anal = GCMSAnalysis(**data)
                
                # 为每个化合物数据创建 GCMSCompound 实例并添加到分析记录中
                for compound_dict in compounds_data:
                    # 移除可能的 id 键，因为这是新创建的化合物
                    compound_dict.pop('id', None)
                    # 创建 GCMSCompound 实例
                    compound = GCMSCompound(**compound_dict)
                    # 将化合物添加到分析记录的 compounds 列表中
                    anal.compounds.append(compound)

                self.session.add(anal)
                print("GCMSAnalysis 对象及关联的 GCMSCompound 对象已添加到 session。") # 添加打印语句
                self.session.commit()
                print("数据库 commit 成功。") # 添加打印语句
                self.refresh_gcms_table()
                QMessageBox.information(self, "添加成功", "GC-MS分析已添加！")

            except Exception as e:
                print(f"添加 GCMS 分析时发生错误: {e}")
                self.session.rollback()
                QMessageBox.critical(self, "添加失败", f"添加 GC-MS 分析时发生错误：{str(e)}")


    def edit_gcms_dialog(self, analysis):
        from sqlalchemy.orm import joinedload
        # 强制立即加载compounds
        analysis = self.session.query(GCMSAnalysis).options(joinedload(GCMSAnalysis.compounds)).get(analysis.id)
        dlg = GCMSAnalysisDialog(self, self.session, analysis)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            # 获取对话框中的新数据
            data = dlg.get_data()
            # 更新主分析信息
            analysis.number = data.get('number', '')
            analysis.name = data.get('name', '')
            analysis.instrument_params = data.get('instrument_params', '')
            analysis.supplier = data.get('supplier', '')
            analysis.perfume_idea = data.get('perfume_idea', '')
            analysis.analysis_time = data.get('analysis_time', None)
            analysis.spectrum_image = data.get('spectrum_image', '')  # 新增，保证谱图字段同步
            # 先删除原有的化合物信息
            if hasattr(analysis, 'compounds'):
                for c in list(analysis.compounds):
                    self.session.delete(c)
                analysis.compounds.clear()
            # 添加新化合物信息
            for compound_dict in data.get('compounds', []):
                compound = GCMSCompound(**compound_dict)
                analysis.compounds.append(compound)
            self.session.commit()
            self.refresh_gcms_table()
            QMessageBox.information(self, "修改成功", "GC-MS分析信息已更新！")


    def delete_gcms(self, analysis):
        """删除GC-MS分析（连带删除其化合物，失败时回滚）"""
        try:
            compound_count = self.session.query(GCMSCompound).filter_by(
                analysis_id=analysis.id).count()
            msg = f"确定要删除GC-MS分析：{analysis.name} 吗？"
            if compound_count:
                msg += f"\n\n注意：该分析包含 {compound_count} 条化合物记录，将一并删除。"

            reply = QMessageBox.question(
                self, "确认删除", msg,
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply != QMessageBox.StandardButton.Yes:
                return

            # 显式删除化合物，避免依赖数据库级联配置
            self.session.query(GCMSCompound).filter_by(
                analysis_id=analysis.id).delete(synchronize_session=False)
            self.session.delete(analysis)
            self.session.commit()
            self.refresh_gcms_table()
            QMessageBox.information(self, "删除成功", "GC-MS分析已删除！")
        except Exception as e:
            self.session.rollback()
            QMessageBox.critical(self, "删除失败", f"删除GC-MS分析时发生错误：{str(e)}")

    def show_gcms_result(self, analysis):
        dlg = GCMSResultDialog(self, analysis, self.session)
        dlg.exec()

    def show_gcms_analysis_func(self):
        dlg = GCMSAnalysisFuncDialog(self, self.session)
        dlg.exec()

    def load_gcms_column_widths(self):
        """Load GCMS table column widths from settings"""
        settings = QSettings("YourCompany", "FragranceManagement") # Replace YourCompany with your company name
        header = self.gcms_table.horizontalHeader()
        widths = settings.value("gcms_table_column_widths")
        if widths is not None:
            # QSettings stores list of ints as string list like '@VariantList(\0\0\0\0\x1\0\0\0\x2\0\0\0\x3\0\0\0\x4)'
            # We need to parse this back to a list of integers.
            # A simpler way might be to store as a string: "150,100,..."
            try:
                if isinstance(widths, str):
                     # If it's a comma-separated string, parse it
                     widths = [int(w) for w in widths.split(',') if w]
                elif isinstance(widths, list):
                     # If it's already a list, ensure elements are integers
                     widths = [int(w) for w in widths]
                else:
                     # If neither, something unexpected happened, treat as empty
                     widths = []

                if len(widths) == header.count():
                    for i in range(header.count()):
                        header.resizeSection(i, widths[i])
            except Exception as e:
                 print(f"Error loading column widths: {e}")
                 # If parsing fails, fallback to default or interactive resize
                 for i in range(header.count()):
                    header.setSectionResizeMode(i, QHeaderView.ResizeMode.Interactive)


    def save_gcms_column_widths(self):
        """Save GCMS table column widths to settings"""
        settings = QSettings("YourCompany", "FragranceManagement") # Replace YourCompany with your company name
        header = self.gcms_table.horizontalHeader()
        widths = [header.sectionSize(i) for i in range(header.count())]
        settings.setValue("gcms_table_column_widths", ",".join(map(str, widths))) # Store as comma-separated string

    def closeEvent(self, event):
        """Handle close event to save settings and cleanup threads"""
        # Save column widths setting
        self.save_gcms_column_widths()
        
        # Clean up any running loaders
        if hasattr(self, 'ingredient_loader') and self.ingredient_loader:
            self.ingredient_loader.stop()
            self.ingredient_loader.wait()  # 等待线程完全结束
            
        if hasattr(self, 'formula_loader') and self.formula_loader:
            self.formula_loader.stop()
            self.formula_loader.wait()  # 等待线程完全结束
            
        # Clean up AI analysis worker threads
        if hasattr(self, 'ai_module') and self.ai_module:
            try:
                # 停止所有可能的工作线程
                if hasattr(self.ai_module, '_current_worker') and self.ai_module._current_worker:
                    if self.ai_module._current_worker.isRunning():
                        self.ai_module._current_worker.quit()
                        self.ai_module._current_worker.wait(3000)  # 等待最多3秒
            except Exception as e:
                print(f"清理AI模块线程时出错: {e}")

        # Close database session
        if hasattr(self, 'session'):
            self.session.close()
            
        super().closeEvent(event)

    def export_gcms_analyses(self):
        """Export the current filtered GCMS analysis data to Excel or CSV file"""
        from PyQt6.QtWidgets import QFileDialog
        
        # 检查pandas可用性
        if not PANDAS_AVAILABLE:
            QMessageBox.warning(self, "功能受限", 
                              "pandas不可用，无法导出Excel/CSV文件。\n\n"
                              "建议：\n"
                              "1. 安装pandas: pip install pandas\n"
                              "2. 或手动复制表格数据\n"
                              "3. 或查看'依赖包问题解决方案.md'")
            return
            
        # Get the current filtered data
        keyword = self.gcms_search.text().strip() if hasattr(self, 'gcms_search') else ''
        query = self.session.query(GCMSAnalysis)
        if keyword:
            like = f"%{keyword}%"
            query = query.filter(
                GCMSAnalysis.name.like(like) |
                GCMSAnalysis.number.like(like) |
                GCMSAnalysis.supplier.like(like)
            )
        analyses = query.all()
        if not analyses:
            QMessageBox.information(self, "导出", "没有可导出的GCMS分析数据！")
            return
        # Assemble export data. Include main analysis fields and a summary of compounds.
        data = []
        columns = [
            '编号', '名称', '供应商', '仪器参数', '整体调香思路', '化合物数量' # Simplified columns
        ]
        for ana in analyses:
            compound_count = len(ana.compounds) if ana.compounds else 0
            data.append([
                ana.number,
                ana.name,
                ana.supplier,
                ana.instrument_params,
                ana.perfume_idea,
                compound_count
            ])
        df = pd.DataFrame(data, columns=columns)
        # Select export filename
        file_path, file_type = QFileDialog.getSaveFileName(self, "导出GCMS分析信息", "GCMS分析信息.xlsx", "Excel 文件 (*.xlsx);;CSV 文件 (*.csv)")
        if not file_path:
            return
        try:
            if file_path.endswith('.csv'):
                df.to_csv(file_path, index=False, encoding='utf-8-sig')
            else:
                df.to_excel(file_path, index=False)
            QMessageBox.information(self, "导出成功", f"GCMS分析信息已导出到：\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", f"导出过程中发生错误：{str(e)}")


    def import_gcms_compounds(self):
        """批量导入GC-MS分析的化合物信息（支持 xlsx/xlsm/xls/csv/tsv/txt/HTML/XML 多格式）"""
        from table_io import read_table_any, SUPPORTED_FILE_FILTER
        file_path, _ = QFileDialog.getOpenFileName(self, "选择化合物数据文件", "", SUPPORTED_FILE_FILTER)
        if not file_path:
            return

        try:
            df = read_table_any(file_path)
            if df.empty:
                QMessageBox.warning(self, "格式错误", "文件为空或无法识别为表格！")
                return

            # Assume required columns exist (adjust column names as per your file)
            required_columns = ['样品编号或名称', '化合物名称', '相对含量'] # Example column names
            if not all(col in df.columns for col in required_columns):
                QMessageBox.warning(self, "表头错误", f"导入文件缺少以下列：{', '.join(required_columns)}")
                return

            imported_count = 0
            for index, row in df.iterrows():
                sample_identifier = row['样品编号或名称']
                compound_name = row['化合物名称']
                relative_content = row['相对含量']
                cas_number = row.get('CAS号', None) # Optional column

                # Find the corresponding GCMSAnalysis record
                # Assuming '样品编号或名称' column in import file matches 'number' or 'name' in GCMSAnalysis
                analysis = self.session.query(GCMSAnalysis).filter(
                    (GCMSAnalysis.number == str(sample_identifier)) | (GCMSAnalysis.name == str(sample_identifier))
                ).first()

                if analysis:
                    # Create GCMSCompound record
                    compound = GCMSCompound(
                        analysis_id=analysis.id,
                        name_cn=compound_name if isinstance(compound_name, str) and any(c >= '\u4e00' and c <= '\u9fa5' for c in compound_name) else None, # Simple check for Chinese
                        name_en=compound_name if isinstance(compound_name, str) and not any(c >= '\u4e00' and c <= '\u9fa5' for c in compound_name) else None,
                        cas_number=str(cas_number) if cas_number is not None else None,
                        relative_content=float(relative_content) if relative_content is not None else 0
                        # Add other fields as needed from your import file
                    )
                    self.session.add(compound)
                    imported_count += 1
                else:
                    print(f"Warning: GCMSAnalysis with identifier '{sample_identifier}' not found.")

            self.session.commit()
            QMessageBox.information(self, "导入完成", f"成功导入 {imported_count} 条化合物记录！")

        except Exception as e:
            self.session.rollback() # Rollback changes in case of error
            QMessageBox.critical(self, "导入失败", f"导入过程中发生错误：{str(e)}")
            print(f"GCMS Compounds import failed: {e}")

    def export_analysis_image(self, pixmap, base_filename):
        """Export the analysis result image."""
        if pixmap is None or pixmap.isNull():
            QMessageBox.warning(self, "导出失败", "没有可导出的图片。")
            return

        from PyQt6.QtWidgets import QFileDialog
        file_path, _ = QFileDialog.getSaveFileName(self, "导出图片", f"{base_filename}.png", "PNG 图片 (*.png);;JPEG 图片 (*.jpg)")
        if file_path:
            if pixmap.save(file_path):
                QMessageBox.information(self, "导出成功", f"图片已导出到：\n{file_path}")
            else:
                QMessageBox.critical(self, "导出失败", "保存图片失败。")

    def export_analysis_data(self, analysis_type):
        """Export the analysis result data to Excel or CSV file."""
        data_to_export = None
        filename = ""
        if analysis_type == "feature":
            data_to_export = self._feature_result_data
            filename = "特征筛查结果数据"
        elif analysis_type == "differential":
            data_to_export = self._differential_result_data
            filename = "差异分析结果数据"
        elif analysis_type == "visualization":
            data_to_export = self._visualization_result_data
            filename = "可视化结果数据"

        if data_to_export is None or data_to_export.empty:
            QMessageBox.warning(self, "导出失败", "没有可导出的数据。")
            return

        from PyQt6.QtWidgets import QFileDialog

        file_path, file_type = QFileDialog.getSaveFileName(self, "导出数据", f"{filename}.xlsx", "Excel 文件 (*.xlsx);;CSV 文件 (*.csv)") # Fixed file type string
        if not file_path:
            return
        try:
            if file_path.endswith('.csv'):
                # Adjust index based on how the data is stored (True if index is meaningful, False otherwise)
                data_to_export.to_csv(file_path, index=True if analysis_type != "feature" else False, encoding='utf-8-sig')
            else:
                # Adjust index based on how the data is stored
                data_to_export.to_excel(file_path, index=True if analysis_type != "feature" else False)
            QMessageBox.information(self, "导出成功", f"数据已导出到：\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", f"导出过程中发生错误：{str(e)}")

    def show_ingredient_detail_by_id(self, ingredient_identifier):
        """切换到原料管理页面并根据标识查找显示原料 (现在支持名称或ID搜索) """
        # 切换到原料管理页面
        self.content.setCurrentIndex(0)

        # 在搜索框中输入原料标识（名称或ID）
        if hasattr(self, 'ingredient_search'):
            # Directly set the text with the provided identifier (name or ID string)
            self.ingredient_search.setText(str(ingredient_identifier))
            # 触发搜索
            self.refresh_ingredient_table()

        # 可选：在表格中选中找到的原料行（如果只有一行）
        # if self.ingredient_table.rowCount() == 1:
        #     self.ingredient_table.selectRow(0)

    def add_stock_dialog(self):
        dialog = StockRecordDialog(self, self.session)
        if dialog.exec():
            data = dialog.get_data()
            if data:
                record = StockRecord(**data)
                self.session.add(record)
                self.session.commit()
                self.refresh_stock_table()
                self.refresh_stock_summary()
                self.refresh_stock_warning()

    def handle_stock_summary_filter(self, text):
        self.stock_summary_filter = text.strip()
        self.stock_summary_page = 1
        self.refresh_stock_summary()

    def handle_stock_summary_sort(self, col):
        if self.stock_summary_sort_col == col:
            self.stock_summary_sort_order = Qt.SortOrder.DescendingOrder if self.stock_summary_sort_order == Qt.SortOrder.AscendingOrder else Qt.SortOrder.AscendingOrder
        else:
            self.stock_summary_sort_col = col
            self.stock_summary_sort_order = Qt.SortOrder.AscendingOrder
        self.refresh_stock_summary()

    def stock_summary_prev_page(self):
        if self.stock_summary_page > 1:
            self.stock_summary_page -= 1
            self.refresh_stock_summary()

    def stock_summary_next_page(self):
        self.stock_summary_page += 1
        self.refresh_stock_summary()

    def stock_record_prev_page(self):
        if self.stock_record_page > 1:
            self.stock_record_page -= 1
            self.refresh_stock_table()
    def stock_record_next_page(self):
        self.stock_record_page += 1
        self.refresh_stock_table()
    
    def use_formula_dialog(self, formula):
        """使用配方对话框 - 改进版本"""
        dialog = FormulaUsageDialog(self, formula, self.session)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            # 刷新相关界面
            self.refresh_stock_table()
            self.refresh_stock_summary()
            self.refresh_stock_warning()


class IngredientDialog(QDialog):
    def __init__(self, parent=None, ingredient=None):
        super().__init__(parent)
        self.setWindowTitle("添加原料" if ingredient is None else "编辑原料")
        self.setMinimumSize(550, 400)  # 设置最小尺寸
        self.resize(650, 550)  # 调整窗口尺寸以适应屏幕
        self.ingredient = ingredient
        self.image_label = QLabel()
        self.image_label.setFixedHeight(120)
        self.image_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        form = QFormLayout()
        self.fields = {}
        labels = [
            ("编号", 'number'), ("CAS", 'cas_number'), ("原料名称", 'name'), ("英文名", 'english_name'),
            ("分子式", 'molecular_formula'), ("化学结构式", 'chemical_structure'), ("分子量", 'molecular_weight'),
            ("沸点", 'boiling_point'), ("溶解性", 'solubility'), ("电子烟国标最大用量", 'max_limit_gb'),
            ("天然存在", 'natural_occurrence'), ("调香用途", 'perfume_usage'), ("香气香韵", 'aroma_character'),
            ("香气变调", 'aroma_change'), ("嗅香香气（1% in PG）", 'sniff_aroma'), ("抽吸感官评价", 'sensory_evaluation'),
            ("香韵构成", 'aroma_composition'), ("成本(元/g)", 'price') # 新增成本字段
        ]
        for label, key in labels:
            line = QLineEdit()
            if ingredient:
                line.setText(str(getattr(ingredient, key, "") or ""))
            self.fields[key] = line
            form.addRow(label, line)
        # CAS号变化时自动获取结构式
        self.fields['cas_number'].editingFinished.connect(self.update_structure_image)
        vbox = QVBoxLayout()
        vbox.addLayout(form)
        vbox.addWidget(QLabel("结构式预览："))
        vbox.addWidget(self.image_label)
        btn_layout = QHBoxLayout()
        ok_btn = QPushButton("保存")
        cancel_btn = QPushButton("取消")
        ok_btn.clicked.connect(self.accept)
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(ok_btn)
        btn_layout.addWidget(cancel_btn)
        vbox.addLayout(btn_layout)
        self.setLayout(vbox)
        # 初始加载结构式
        self.update_structure_image()
    def update_structure_image(self):
        cas = self.fields['cas_number'].text().strip()
        if not cas:
            self.image_label.clear()
            return
        # Use PubChem API to get structure image
        try:
            url = f"https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/{cas}/PNG"
            resp = requests.get(url, timeout=5)
            if resp.status_code == 200:
                pix = QPixmap()
                pix.loadFromData(resp.content)
                self.image_label.setPixmap(pix.scaledToHeight(100))
                self.fields['chemical_structure'].setText(url)
            else:
                self.image_label.setText("未找到结构式")
        except Exception:
            self.image_label.setText("获取失败")
    def get_data(self):
        return {k: self.fields[k].text().strip() for k in self.fields}


class FormulaDialog(QDialog):
    def __init__(self, parent, session, formula_data=None, existing_formula=None):
        super().__init__(parent)
        self.setWindowTitle("编辑配方" if existing_formula else "添加配方")
        self.resize(700, 600)
        self.session = session
        self.formula_data = formula_data
        self.existing_formula = existing_formula # Store existing formula object if editing
        self.selected_ingredients = []
        if self.existing_formula:
            # 编辑时根据content内容自动加载所有原料
            self.selected_ingredients = self._get_ingredients_from_content()
            self.init_input_amounts()
        else:
            # 新建时先选择原料
            self.init_select_ingredients()

    def _get_ingredients_from_content(self):
        # 根据content内容获取所有涉及的原料对象
        ingredients = []
        name_set = set()
        content = getattr(self.existing_formula, 'content', '')
        if content:
            for item in content.split(','):
                if ':' in item:
                    parts = item.split(':')
                    name = parts[0].strip()
                    if name and name not in name_set:
                        ing = self.session.query(Ingredient).filter_by(name=name).first()
                        if ing:
                            ingredients.append(ing)
                            name_set.add(name)
        return ingredients

    def init_select_ingredients(self):
        layout = QVBoxLayout()
        layout.addWidget(QLabel("请选择需要参与本次配方的原料："))
        
        # 添加搜索框
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("🔍 搜索原料:"))
        self.ingredient_search = QLineEdit()
        self.ingredient_search.setPlaceholderText("输入原料名称、编号、CAS号或英文名进行筛选...")
        self.ingredient_search.textChanged.connect(self.filter_ingredients)
        search_layout.addWidget(self.ingredient_search)
        
        # 添加快速选择按钮
        quick_select_layout = QHBoxLayout()
        select_all_btn = QPushButton("✅ 全选")
        select_all_btn.clicked.connect(self.select_all_ingredients)
        select_none_btn = QPushButton("❌ 全不选")
        select_none_btn.clicked.connect(self.select_none_ingredients)
        quick_select_layout.addWidget(select_all_btn)
        quick_select_layout.addWidget(select_none_btn)
        quick_select_layout.addStretch()
        
        layout.addLayout(search_layout)
        layout.addLayout(quick_select_layout)
        
        self.list_widget = QListWidget()
        self.ingredients = self.session.query(Ingredient).all()
        self.ingredient_selection_state = {}  # 存储选择状态，使用原料ID作为键

        # Pre-select ingredients if editing an existing formula
        selected_ingredient_ids = []
        if self.existing_formula:
            # Fetch existing ingredient associations for this formula
            existing_associations = self.session.query(ingredient_formula).filter_by(formula_id=self.existing_formula.id).all()
            selected_ingredient_ids = [assoc.ingredient_id for assoc in existing_associations]

        # 初始化选择状态
        for ing in self.ingredients:
            self.ingredient_selection_state[ing.id] = ing.id in selected_ingredient_ids

        layout.addWidget(self.list_widget)
        
        # 添加统计信息
        self.ingredient_count_label = QLabel()
        layout.addWidget(self.ingredient_count_label)
        
        # 现在可以安全地填充列表了（因为label已经创建）
        self.populate_ingredient_list()
        
        # 监听选择状态变化，同步到状态字典
        self.list_widget.itemChanged.connect(self.sync_selection_state)
        
        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btns.accepted.connect(self.select_ingredients)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)
        self.setLayout(layout)
    
    def populate_ingredient_list(self, filter_text=""):
        """填充原料列表"""
        self.list_widget.clear()
        filter_text = filter_text.lower()
        
        visible_count = 0
        for ing in self.ingredients:
            # 检查是否匹配筛选条件 - 增加对原料编号的支持
            if (not filter_text or 
                filter_text in ing.name.lower() or 
                filter_text in (ing.cas_number or "").lower() or 
                filter_text in (ing.english_name or "").lower() or
                filter_text in (ing.number or "").lower()):  # 新增：支持按原料编号筛选
                
                # 显示格式：原料名称（编号：编号 CAS：CAS号）
                display_text = f"{ing.name}（编号:{ing.number} CAS:{ing.cas_number}）"
                item = QListWidgetItem(display_text)
                item.setData(Qt.ItemDataRole.UserRole, ing.id)
                # 从状态字典中获取选择状态
                is_selected = self.ingredient_selection_state.get(ing.id, False)
                item.setCheckState(Qt.CheckState.Checked if is_selected else Qt.CheckState.Unchecked)
                self.list_widget.addItem(item)
                visible_count += 1
        
        self.update_ingredient_count(visible_count)
    
    def filter_ingredients(self):
        """筛选原料列表"""
        search_text = self.ingredient_search.text().strip()
        self.populate_ingredient_list(search_text)
    
    def select_all_ingredients(self):
        """全选当前显示的原料"""
        # 临时断开信号连接，避免重复触发同步
        self.list_widget.itemChanged.disconnect()
        
        # 收集当前显示的原料ID并更新状态
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            ingredient_id = item.data(Qt.ItemDataRole.UserRole)
            item.setCheckState(Qt.CheckState.Checked)
            self.ingredient_selection_state[ingredient_id] = True
        
        # 重新连接信号
        self.list_widget.itemChanged.connect(self.sync_selection_state)
        self.update_ingredient_count()
    
    def select_none_ingredients(self):
        """取消选择所有原料"""
        # 临时断开信号连接，避免重复触发同步
        self.list_widget.itemChanged.disconnect()
        
        # 收集当前显示的原料ID并更新状态
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            ingredient_id = item.data(Qt.ItemDataRole.UserRole)
            item.setCheckState(Qt.CheckState.Unchecked)
            self.ingredient_selection_state[ingredient_id] = False
        
        # 重新连接信号
        self.list_widget.itemChanged.connect(self.sync_selection_state)
        self.update_ingredient_count()
    
    def sync_selection_state(self, changed_item):
        """同步选择状态到状态字典"""
        changed_ingredient_id = changed_item.data(Qt.ItemDataRole.UserRole)
        new_state = changed_item.checkState() == Qt.CheckState.Checked
        
        # 更新状态字典
        self.ingredient_selection_state[changed_ingredient_id] = new_state
        
        # 更新统计
        self.update_ingredient_count()
    
    def update_ingredient_count(self, visible_count=None):
        """更新原料统计信息"""
        # 安全检查：确保label已经创建
        if not hasattr(self, 'ingredient_count_label'):
            return
            
        if visible_count is None:
            visible_count = self.list_widget.count()
        
        selected_count = 0
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                selected_count += 1
        
        total_count = len(self.ingredients)
        self.ingredient_count_label.setText(
            f"📊 显示: {visible_count}/{total_count} 个原料，已选择: {selected_count} 个"
        )

    def select_ingredients(self):
        # 从状态字典中获取所有被选中的原料
        self.selected_ingredients = []
        for ingredient_id, is_selected in self.ingredient_selection_state.items():
            if is_selected:
                ingredient = self.session.query(Ingredient).get(ingredient_id)
                if ingredient:
                    self.selected_ingredients.append(ingredient)

        if not self.selected_ingredients:
            QMessageBox.warning(self, "提示", "请至少选择一个原料！")
            return
        self.init_input_amounts()

    def init_input_amounts(self):
        # 若无布局，先设置一个主布局
        if self.layout() is None:
            self.setLayout(QVBoxLayout())
        # Clear original layout
        for i in reversed(range(self.layout().count())):
            widget = self.layout().itemAt(i).widget()
            if widget:
                widget.setParent(None)

        form = QFormLayout()
        self.fields = {}
        self.fields['number'] = QLineEdit()
        self.fields['name'] = QLineEdit()
        self.fields['creator'] = QLineEdit()
        self.fields['created_at'] = QLineEdit()
        self.fields['updated_at'] = QLineEdit()
        self.fields['evaluation'] = QLineEdit()
        self.fields['description'] = QLineEdit() # Added description field

        # Set initial values if editing
        if self.formula_data:
            self.fields['number'].setText(self.formula_data.get('number', '') or "")
            self.fields['name'].setText(self.formula_data.get('name', '') or "")
            self.fields['creator'].setText(self.formula_data.get('creator', '') or "")
            self.fields['evaluation'].setText(self.formula_data.get('evaluation', '') or "")
            self.fields['description'].setText(self.formula_data.get('description', '') or "")
            created_at = self.formula_data.get('created_at')
            if isinstance(created_at, datetime):
                created_at_str = created_at.strftime("%Y-%m-%d %H:%M:%S")
            else:
                created_at_str = str(created_at) if created_at else ""
            self.fields['created_at'].setText(created_at_str)
            updated_at = self.formula_data.get('updated_at')
            if isinstance(updated_at, datetime):
                updated_at_str = updated_at.strftime("%Y-%m-%d %H:%M:%S")
            else:
                updated_at_str = str(updated_at) if updated_at else ""
            self.fields['updated_at'].setText(updated_at_str)
        else:
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.fields['created_at'].setText(now_str)
            self.fields['updated_at'].setText(now_str)
        self.fields['created_at'].setReadOnly(True)
        self.fields['updated_at'].setReadOnly(True)

        form.addRow("编号", self.fields['number'])
        form.addRow("配方名称", self.fields['name'])
        form.addRow("创建人", self.fields['creator'])
        form.addRow("创建时间", self.fields['created_at'])
        form.addRow("修改时间", self.fields['updated_at'])
        form.addRow("感官评价", self.fields['evaluation'])
        form.addRow("描述", self.fields['description'])

        # 添加导入按钮
        import_layout = QHBoxLayout()
        import_btn = QPushButton("📁 导入配方")
        import_btn.setToolTip("从CSV文件导入配方数据")
        import_btn.clicked.connect(self.import_formula_data)
        import_layout.addWidget(import_btn)
        import_layout.addStretch()  # 添加弹性空间
        self.layout().addLayout(import_layout)

        self.amount_table = QTableWidget()
        self.amount_table.setRowCount(len(self.selected_ingredients))
        self.amount_table.setColumnCount(6)  # 增加2列：稀释倍数和稀释溶剂
        self.amount_table.setHorizontalHeaderLabels(["原料名称", "CAS", "用量(g)", "百分比(%)", "稀释倍数", "稀释溶剂"])
        self.amount_spins = []
        self.dilution_spins = []  # 稀释倍数输入框列表
        self.solvent_combos = []  # 稀释溶剂下拉框列表

        # Set initial amounts if editing
        initial_amounts = {}
        if self.formula_data and 'content' in self.formula_data:
            # 解析content字符串，填充用量、百分比、稀释倍数和稀释溶剂
            for item in self.formula_data['content'].split(','):
                if ':' in item:
                    parts = item.split(':')
                    name_in_content = parts[0].strip()
                    percent = 0
                    amount = 0
                    dilution = 1.0  # 默认稀释倍数
                    solvent = "PG"  # 默认溶剂
                    
                    if len(parts) > 1:
                        percent_str = parts[1].strip().replace('%', '')
                        try:
                            percent = float(percent_str)
                        except Exception:
                            percent = 0
                    if len(parts) > 2:
                        try:
                            amount = float(parts[2].strip().replace('g', ''))
                        except Exception:
                            amount = 0
                    if len(parts) > 3:  # 稀释倍数
                        try:
                            dilution = float(parts[3].strip())
                        except Exception:
                            dilution = 1.0
                    if len(parts) > 4:  # 稀释溶剂
                        solvent = parts[4].strip()
                        
                    initial_amounts[name_in_content] = (amount, percent, dilution, solvent)

        for row, ing in enumerate(self.selected_ingredients):
            self.amount_table.setItem(row, 0, QTableWidgetItem(ing.name))
            self.amount_table.setItem(row, 1, QTableWidgetItem(ing.cas_number))
            
            # 用量输入框
            spin = QDoubleSpinBox()
            spin.setRange(0, 999999)
            spin.setDecimals(4)
            spin.setSingleStep(0.01)
            # 设置初始用量
            initial_amount = 0
            if ing.name in initial_amounts:
                initial_amount = initial_amounts[ing.name][0]
            spin.setValue(initial_amount)
            spin.valueChanged.connect(self.update_percent)
            self.amount_table.setCellWidget(row, 2, spin)
            
            # 初始化百分比
            percent_item = QTableWidgetItem("0.00")
            percent_item.setFlags(percent_item.flags() ^ Qt.ItemFlag.ItemIsEditable)
            if ing.name in initial_amounts:
                percent_item.setText(f"{initial_amounts[ing.name][1]:.2f}")
            self.amount_table.setItem(row, 3, percent_item)
            
            # 稀释倍数输入框
            dilution_spin = QDoubleSpinBox()
            dilution_spin.setRange(1, 1000)  # 稀释倍数从1倍到1000倍
            dilution_spin.setDecimals(1)
            dilution_spin.setSingleStep(0.1)
            dilution_spin.setToolTip("稀释倍数，1表示不稀释")
            # 设置初始稀释倍数
            initial_dilution = 1.0
            if ing.name in initial_amounts and len(initial_amounts[ing.name]) > 2:
                initial_dilution = initial_amounts[ing.name][2]
            dilution_spin.setValue(initial_dilution)
            self.amount_table.setCellWidget(row, 4, dilution_spin)
            
            # 稀释溶剂下拉框
            solvent_combo = QComboBox()
            solvent_options = ["PG", "乙醇", "MCT", "甘油", "其他"]
            solvent_combo.addItems(solvent_options)
            solvent_combo.setToolTip("稀释溶剂")
            # 设置初始稀释溶剂
            initial_solvent = "PG"
            if ing.name in initial_amounts and len(initial_amounts[ing.name]) > 3:
                initial_solvent = initial_amounts[ing.name][3]
            solvent_combo.setCurrentText(initial_solvent)
            self.amount_table.setCellWidget(row, 5, solvent_combo)
            
            self.amount_spins.append(spin)
            self.dilution_spins.append(dilution_spin)
            self.solvent_combos.append(solvent_combo)

        self.amount_table.resizeColumnsToContents()
        self.amount_table.resizeRowsToContents()
        self.layout().addLayout(form)
        self.layout().addWidget(QLabel("请填写每种原料的用量，系统自动计算百分比："))
        self.layout().addWidget(self.amount_table)
        self.total_label = QLabel("总用量: 0.00  总成本: 0.00")
        self.layout().addWidget(self.total_label)
        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        self.layout().addWidget(btns)
        self._current_total = 0
        self._current_cost = 0
        self._current_percent = []
        self.update_percent()

    def import_formula_data(self):
        """导入配方用量数据（支持 xlsx/xlsm/xls/csv/tsv/txt 及伪Excel文件）"""
        try:
            from table_io import read_table_any, SUPPORTED_FILE_FILTER
        except ImportError as e:
            QMessageBox.critical(self, "导入失败", f"表格读取模块加载失败：{e}")
            return

        file_path, _ = QFileDialog.getOpenFileName(
            self, "导入配方数据", "", SUPPORTED_FILE_FILTER)

        if not file_path:
            return

        try:
            df = read_table_any(file_path)
            if df.empty:
                QMessageBox.warning(self, "警告", "文件为空或格式不正确")
                return

            # 列名兼容：忽略大小写与空白，常见别名统一到标准列名
            col_alias = {
                '原料名称': '原料名称', '原料': '原料名称', '名称': '原料名称',
                '用量': '用量', '用量(g)': '用量', '用量g': '用量',
                '添加量': '用量',
                '稀释倍数': '稀释倍数', '稀释': '稀释倍数',
                '稀释溶剂': '稀释溶剂', '溶剂': '稀释溶剂',
            }
            lower_map = {str(c).strip().lower(): c for c in df.columns}
            renamed, used_std = {}, set()
            for src, std in col_alias.items():
                if std in used_std:
                    continue
                col = lower_map.get(src.lower())
                if col is not None:
                    renamed[col] = std
                    used_std.add(std)
            df = df.rename(columns=renamed)

            required_columns = ['原料名称', '用量']
            optional_columns = ['稀释倍数', '稀释溶剂']
            missing = [c for c in required_columns if c not in df.columns]
            if missing:
                QMessageBox.warning(
                    self,
                    "格式错误",
                    f"文件必须包含以下列：{'、'.join(required_columns)}\n"
                    f"可选列：{'、'.join(optional_columns)}\n\n"
                    f"当前文件列：{'、'.join(str(c) for c in df.columns)}\n"
                    f"缺少：{'、'.join(missing)}"
                )
                return

            data = df.to_dict('records')

            # 导入数据到表格
            imported_count = 0
            unmatched = []
            for csv_row in data:
                ingredient_name = csv_row.get('原料名称', '').strip()
                if not ingredient_name:
                    continue
                    
                # 查找对应的原料行
                matched_row = None
                for table_row in range(self.amount_table.rowCount()):
                    table_ingredient_name = self.amount_table.item(table_row, 0).text()
                    if table_ingredient_name == ingredient_name:
                        matched_row = table_row
                        break
                if matched_row is None:
                    unmatched.append(ingredient_name)
                    continue

                table_row = matched_row
                # 设置用量
                try:
                    amount = float(csv_row.get('用量', '0') or 0)
                    self.amount_spins[table_row].setValue(amount)
                except ValueError:
                    pass

                # 设置稀释倍数
                if '稀释倍数' in csv_row:
                    try:
                        dilution = float(csv_row.get('稀释倍数', '1.0') or 1.0)
                        self.dilution_spins[table_row].setValue(dilution)
                    except ValueError:
                        pass

                # 设置稀释溶剂
                if '稀释溶剂' in csv_row:
                    solvent = (csv_row.get('稀释溶剂') or 'PG').strip()
                    if solvent:
                        combo = self.solvent_combos[table_row]
                        index = combo.findText(solvent)
                        if index >= 0:
                            combo.setCurrentIndex(index)
                        else:
                            combo.setCurrentText(solvent)

                imported_count += 1

            self.update_percent()  # 重新计算百分比
            msg = f"成功导入 {imported_count} 个原料的数据"
            if unmatched:
                shown = "、".join(unmatched[:8])
                more = f" 等 {len(unmatched)} 项" if len(unmatched) > 8 else ""
                msg += (f"\n\n以下原料未在当前配方的原料清单中找到，已跳过：\n"
                        f"{shown}{more}")
            QMessageBox.information(self, "导入成功", msg)
            
        except Exception as e:
            QMessageBox.critical(self, "导入失败", f"导入过程中发生错误：{str(e)}")

    def update_percent(self):
        total = sum(spin.value() for spin in self.amount_spins)
        cost = 0
        self._current_percent = [] # Reset percentage list
        for row, spin in enumerate(self.amount_spins):
            val = spin.value()
            percent = (val / total * 100) if total > 0 else 0
            percent_item = self.amount_table.item(row, 3)
            percent_item.setText(f"{percent:.2f}")
            self._current_percent.append(percent) # Store percentage

            try:
                price = float(self.selected_ingredients[row].price or 0)
            except Exception:
                price = 0
            cost += price * val

        self.total_label.setText(f"总用量: {total:.2f}  总成本: {cost:.2f}")
        self._current_total = total
        self._current_cost = cost

    def accept(self):
        if self._current_total <= 0:
            QMessageBox.warning(self, "校验失败", "请填写原料用量，总用量需大于0！")
            return
        super().accept()

    def get_data(self):
        # Formula content format: IngredientName:Percentage%:Amountg:DilutionRatio:Solvent,... Only export ingredients with amount > 0
        content = []
        for row, spin in enumerate(self.amount_spins):
            val = spin.value()
            if val > 0:
                name = self.amount_table.item(row, 0).text()
                percent = self.amount_table.item(row, 3).text() # Get percentage text
                dilution = self.dilution_spins[row].value()  # 获取稀释倍数
                solvent = self.solvent_combos[row].currentText()  # 获取稀释溶剂
                # 新格式：名称:百分比%:用量g:稀释倍数:稀释溶剂
                content.append(f"{name}:{percent}%:{val}g:{dilution}:{solvent}")

        # Calculate total cost here before returning
        total_cost = 0
        for row, spin in enumerate(self.amount_spins):
             val = spin.value()
             try:
                 price = float(self.selected_ingredients[row].price or 0)
             except Exception:
                 price = 0
             total_cost += price * val

        # Convert created_at string back to datetime object
        created_at_str = self.fields['created_at'].text().strip()
        try:
            created_at_obj = datetime.strptime(created_at_str, "%Y-%m-%d %H:%M:%S")
        except ValueError:
            # Handle potential format issues, maybe return None or the original string
            # Depending on how strict you want validation, you might show an error here
            print(f"Warning: Could not parse created_at string '{created_at_str}' to datetime.")
            created_at_obj = None # Or created_at_str

        return {
            'number': self.fields['number'].text().strip(),
            'name': self.fields['name'].text().strip(),
            'creator': self.fields['creator'].text().strip(),
            'created_at': created_at_obj, # Return as datetime object
            'updated_at': datetime.now(), # Set updated_at to current time
            'content': ', '.join(content),
            'evaluation': self.fields['evaluation'].text().strip(),
            'total_cost': total_cost, # Return calculated total cost as float
            'description': self.fields['description'].text().strip(), # Return description
        }


class FormulaDetailDialog(QDialog):
    def __init__(self, formula, session, parent=None):
        super().__init__(parent)
        self.formula = formula
        self.session = session
        self.main_window = parent
        self.setWindowTitle(f"配方详情 - {formula.name}")
        self.resize(600, 400)
        
        layout = QVBoxLayout()
        
        # 基本信息
        info_group = QGroupBox("基本信息")
        info_layout = QFormLayout()
        info_layout.addRow("配方名称:", QLabel(formula.name))
        info_layout.addRow("编号:", QLabel(formula.number or ""))
        info_layout.addRow("版本:", QLabel(formula.version or ""))
        info_layout.addRow("创建人:", QLabel(formula.creator or ""))
        info_layout.addRow("创建时间:", QLabel(formula.created_at.strftime("%Y-%m-%d %H:%M:%S") if formula.created_at else ""))
        info_layout.addRow("总成本:", QLabel(f"{formula.total_cost:.2f}元/g" if formula.total_cost else ""))
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)
        
        # 原料列表
        ingredients_group = QGroupBox("原料列表")
        ingredients_layout = QVBoxLayout()
        self.ingredients_table = QTableWidget()
        self.ingredients_table.setColumnCount(4)
        self.ingredients_table.setHorizontalHeaderLabels(["原料名称", "百分比", "用量(g)", "成本"])
        self.ingredients_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        ingredients_layout.addWidget(self.ingredients_table)
        ingredients_group.setLayout(ingredients_layout)
        layout.addWidget(ingredients_group)
        
        # 感官评价
        if formula.evaluation:
            eval_group = QGroupBox("感官评价")
            eval_layout = QVBoxLayout()
            eval_text = QTextEdit()
            eval_text.setPlainText(formula.evaluation)
            eval_text.setReadOnly(True)
            eval_layout.addWidget(eval_text)
            eval_group.setLayout(eval_layout)
            layout.addWidget(eval_group)
        
        # 按钮区域
        button_layout = QHBoxLayout()
        use_btn = QPushButton("使用配方")
        use_btn.clicked.connect(self.use_formula)
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.accept)
        button_layout.addWidget(use_btn)
        button_layout.addWidget(close_btn)
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
        self._load_formula_ingredients()
    
    def use_formula(self):
        """使用配方"""
        # 直接使用传入的主窗口引用
        if self.main_window and hasattr(self.main_window, 'use_formula_dialog'):
            self.main_window.use_formula_dialog(self.formula)
            self.close()  # 关闭配方详情对话框
        else:
            QMessageBox.information(self, "提示", "找不到主窗口，无法使用配方功能")
    
    def _load_formula_ingredients(self):
        """加载配方原料列表"""
        if not self.formula.content:
            return
            
        ingredients = []
        for item in self.formula.content.split(','):
            if not item.strip():
                continue
            parts = item.strip().split(':')
            if len(parts) < 3:
                continue
                
            name = parts[0].strip()
            percentage = float(parts[1].rstrip('%'))
            amount = float(parts[2].rstrip('g'))
            
            # 新增：解析稀释倍数和稀释溶剂（向后兼容）
            dilution = 1.0
            solvent = "PG"
            if len(parts) > 3:
                try:
                    dilution = float(parts[3])
                except:
                    dilution = 1.0
            if len(parts) > 4:
                solvent = parts[4]
            
            # 查询原料成本
            ingredient = self.session.query(Ingredient).filter_by(name=name).first()
            cost = ingredient.price * amount if ingredient and ingredient.price else 0
            
            ingredients.append((name, percentage, amount, cost, dilution, solvent))
        
        # 更新表格列数以显示稀释信息
        self.ingredients_table.setColumnCount(6)
        self.ingredients_table.setHorizontalHeaderLabels(["原料名称", "百分比", "用量(g)", "成本", "稀释倍数", "稀释溶剂"])
        self.ingredients_table.setRowCount(len(ingredients))
        
        for row, (name, percentage, amount, cost, dilution, solvent) in enumerate(ingredients):
            self.ingredients_table.setItem(row, 0, QTableWidgetItem(name))
            self.ingredients_table.setItem(row, 1, QTableWidgetItem(f"{percentage:.2f}%"))
            self.ingredients_table.setItem(row, 2, QTableWidgetItem(f"{amount:.2f}"))
            self.ingredients_table.setItem(row, 3, QTableWidgetItem(f"{cost:.2f}"))
            self.ingredients_table.setItem(row, 4, QTableWidgetItem(f"{dilution:.1f}"))
            self.ingredients_table.setItem(row, 5, QTableWidgetItem(solvent))


class GCMSAnalysisDialog(QDialog):
    def __init__(self, parent=None, session=None, analysis=None):
        super().__init__(parent)
        self.setWindowTitle("添加GC-MS分析" if analysis is None else "编辑GC-MS分析")
        self.resize(700, 600)
        self.session = session
        self.analysis = analysis
        self.spectrum_image_path = None
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowMinimizeButtonHint | Qt.WindowType.WindowMaximizeButtonHint)
        main_layout = QVBoxLayout()
        basic_info_group = QWidget()
        basic_info_layout = QVBoxLayout(basic_info_group)
        basic_info_layout.addWidget(QLabel("样品基本信息"))
        form = QFormLayout()
        self.fields = {}
        fields_info = [
            ("编号", 'number', QLineEdit),
            ("样品名称", 'name', QLineEdit),
            ("分析时间", 'analysis_time', QLineEdit),
            ("仪器参数", 'instrument_params', QLineEdit),
            ("供应商", 'supplier', QLineEdit),
            ("整体调香思路", 'perfume_idea', QTextEdit),
        ]
        for label, key, widget_type in fields_info:
            if widget_type == QTextEdit:
                field_widget = widget_type()
                field_widget.setMinimumHeight(60)
            else:
                field_widget = widget_type()
            if analysis:
                value = getattr(analysis, key, "")
                if value is not None:
                    if key == 'analysis_time' and isinstance(value, datetime):
                        field_widget.setText(value.strftime("%Y-%m-%d %H:%M:%S"))
                    elif isinstance(field_widget, QTextEdit):
                         field_widget.setPlainText(str(value))
                    elif isinstance(field_widget, QLineEdit):
                        field_widget.setText(str(value))
            self.fields[key] = field_widget
            form.addRow(label, field_widget)
        # 谱图导入与预览
        spectrum_layout = QHBoxLayout()
        self.spectrum_label = QLabel()
        self.spectrum_label.setFixedHeight(60)
        self.spectrum_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        spectrum_btn = QPushButton("导入谱图")
        spectrum_btn.clicked.connect(self.import_spectrum_image)
        spectrum_layout.addWidget(QLabel("谱图："))
        spectrum_layout.addWidget(self.spectrum_label)
        spectrum_layout.addWidget(spectrum_btn)
        form.addRow(spectrum_layout)
        # 保证表单被添加到主布局
        main_layout.addLayout(form)
        
        # 化合物信息表格
        # 筛选输入框
        self.compound_filter = QLineEdit()
        self.compound_filter.setPlaceholderText("筛选化合物（支持CAS号、英文名、中文名模糊搜索）")
        self.compound_filter.textChanged.connect(self.filter_compounds_table)
        main_layout.addWidget(self.compound_filter)

        self.compound_table = QTableWidget()
        self.compound_table.setColumnCount(7)
        self.compound_table.setHorizontalHeaderLabels([
            "CAS号", "组分RT", "化合物名称", "中文名", "匹配因子", "分子式", "相对含量mg/L"
        ])
        self.compound_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.compound_table.setSortingEnabled(True)
        main_layout.addWidget(self.compound_table)
        
        # 导入化合物按钮
        import_btn = QPushButton("导入化合物")
        import_btn.clicked.connect(self.import_compounds)
        main_layout.addWidget(import_btn)
        
        # 如果是编辑，加载已有化合物
        if analysis:
            self._load_existing_compounds()
        # 预加载已有谱图
        if analysis and getattr(analysis, 'spectrum_image', None):
            self.spectrum_image_path = analysis.spectrum_image
            self.load_spectrum_preview(self.spectrum_image_path)
        # 添加确定和取消按钮
        btn_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        main_layout.addWidget(btn_box)
        self.setLayout(main_layout)

    def import_spectrum_image(self):
        from PyQt6.QtWidgets import QFileDialog
        file_path, _ = QFileDialog.getOpenFileName(self, "选择谱图图片", "", "图片文件 (*.png *.jpg *.jpeg *.bmp)")
        if file_path:
            self.spectrum_image_path = file_path
            self.load_spectrum_preview(file_path)
    def load_spectrum_preview(self, file_path):
        from PyQt6.QtGui import QPixmap
        pix = QPixmap(file_path)
        if not pix.isNull():
            self.spectrum_label.setPixmap(pix.scaledToHeight(50))
            self.spectrum_label.setToolTip("点击放大")
            self.spectrum_label.mousePressEvent = lambda e, img=pix: self.parent().show_image_dialog(img, "谱图预览")
        else:
            self.spectrum_label.setText("图片加载失败")

    def get_data(self):
        data = {}
        for key, field_widget in self.fields.items():
            if key == 'analysis_time':
                 continue
            if isinstance(field_widget, QTextEdit):
                data[key] = field_widget.toPlainText().strip()
            elif isinstance(field_widget, QLineEdit):
                data[key] = field_widget.text().strip()
        # 单独处理 analysis_time
        time_str = self.fields['analysis_time'].text().strip()
        try:
            data['analysis_time'] = datetime.strptime(time_str, "%Y-%m-%d %H:%M:%S") if time_str else None
        except ValueError:
            print(f"Warning: Could not parse analysis_time string '{time_str}'. Setting to None.")
            data['analysis_time'] = None
        data.pop('operator', None)
        data.pop('status', None)
        # 谱图图片路径
        data['spectrum_image'] = self.spectrum_image_path or ''
        # 新增：收集表格里的化合物数据
        compounds = []
        for row in range(self.compound_table.rowCount()):
            compound = {
                'cas': self.compound_table.item(row, 0).text() if self.compound_table.item(row, 0) else '',
                'rt': self.compound_table.item(row, 1).text() if self.compound_table.item(row, 1) else '',
                'name_en': self.compound_table.item(row, 2).text() if self.compound_table.item(row, 2) else '',
                'name_cn': self.compound_table.item(row, 3).text() if self.compound_table.item(row, 3) else '',
                'match_factor': self.compound_table.item(row, 4).text() if self.compound_table.item(row, 4) else '',
                'formula': self.compound_table.item(row, 5).text() if self.compound_table.item(row, 5) else '',
                'relative_content': float(self.compound_table.item(row, 6).text()) if self.compound_table.item(row, 6) and self.compound_table.item(row, 6).text() else 0.0,
            }
            compounds.append(compound)
        data['compounds'] = compounds
        return data

    def _load_existing_compounds(self):
        """加载已存在的化合物数据到表格"""
        self.compound_table.setRowCount(0)  # 清空表格
        if self.analysis and self.analysis.compounds:
            for compound in self.analysis.compounds:
                row = self.compound_table.rowCount()
                self.compound_table.insertRow(row)
                # Populate table cells using correct indices based on the 7 columns
                # CAS号(0), 组分RT(1), 化合物名称(2), 中文名(3), 匹配因子(4), 分子式(5), 相对含量mg/L(6)
                
                # CAS号 (Column 0)
                self.compound_table.setItem(row, 0, QTableWidgetItem(compound.cas or ""))
                
                # 组分RT (Column 1)
                rt_item = NumericTableWidgetItem(str(compound.rt or ""))
                rt_item.setData(Qt.ItemDataRole.UserRole, float(compound.rt) if compound.rt is not None else -1.0) # Store numeric value for sorting
                self.compound_table.setItem(row, 1, rt_item)
                
                # 化合物名称 (Column 2)
                self.compound_table.setItem(row, 2, QTableWidgetItem(compound.name_en or ""))
                
                # 中文名 (Column 3)
                self.compound_table.setItem(row, 3, QTableWidgetItem(compound.name_cn or ""))
                
                # 匹配因子 (Column 4)
                match_factor_item = NumericTableWidgetItem(str(compound.match_factor or ""))
                try:
                    match_factor_value = float(compound.match_factor) if compound.match_factor is not None else -1.0
                except ValueError:
                    match_factor_value = -1.0
                match_factor_item.setData(Qt.ItemDataRole.UserRole, match_factor_value)
                self.compound_table.setItem(row, 4, match_factor_item)
                
                # 分子式 (Column 5)
                self.compound_table.setItem(row, 5, QTableWidgetItem(compound.formula or ""))
                
                # 相对含量mg/L (Column 6)
                relative_content_item = NumericTableWidgetItem(str(compound.relative_content or ""))
                try:
                    relative_content_value = float(compound.relative_content) if compound.relative_content is not None else -1.0
                except ValueError:
                    relative_content_value = -1.0
                relative_content_item.setData(Qt.ItemDataRole.UserRole, relative_content_value)
                self.compound_table.setItem(row, 6, relative_content_item)

        self.compound_table.resizeColumnsToContents()
        self.compound_table.resizeRowsToContents()

    def import_compounds(self):
        """导入化合物信息到表格"""
        from PyQt6.QtWidgets import QFileDialog, QMessageBox
        from datetime import datetime # Import datetime for potential parsing

        # 检查pandas可用性
        if not PANDAS_AVAILABLE:
            QMessageBox.warning(self, "功能受限", 
                              "pandas不可用，无法导入Excel/CSV文件。\n\n"
                              "建议：\n"
                              "1. 安装pandas: pip install pandas\n"
                              "2. 或手动输入化合物数据\n"
                              "3. 或查看'依赖包问题解决方案.md'")
            return

        try:
            from table_io import read_table_any, SUPPORTED_FILE_FILTER
        except ImportError as e:
            QMessageBox.critical(self, "导入失败", f"表格读取模块加载失败：{e}")
            return

        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择化合物数据文件", "", SUPPORTED_FILE_FILTER)
        if not file_path:
            return

        try:
            df = read_table_any(file_path)
            if df.empty:
                QMessageBox.warning(self, "警告", "文件为空或格式不正确")
                return

            # 标准列名及其允许的别名（匹配时忽略大小写与首尾空白）
            std_columns = ['CAS号', '组分RT', '化合物名称', '中文名',
                           '匹配因子', '分子式', '相对含量mg/L']
            aliases = {
                'CAS号': ['cas号', 'cas', 'cas no', 'cas number'],
                '组分RT': ['组分rt', 'rt', 'rt(min)', '保留时间',
                          'retention time'],
                '化合物名称': ['化合物名称', '英文名称', '英文名', 'compound',
                              'compound name', 'name_en'],
                '中文名': ['中文名', '中文名称', 'name_cn', 'chinese name'],
                '匹配因子': ['匹配因子', 'match', 'match factor', '相似度'],
                '分子式': ['分子式', 'formula', 'molecular formula'],
                '相对含量mg/L': ['相对含量mg/l', '相对含量(mg/l)', '相对含量',
                                '含量', '浓度', 'content', 'relative content'],
            }
            lower_map = {str(c).strip().lower(): c for c in df.columns}
            col_of, missing_cols = {}, []
            for std in std_columns:
                hit = next((lower_map[a] for a in aliases[std] if a in lower_map), None)
                if hit is None:
                    missing_cols.append(std)
                else:
                    col_of[std] = hit

            if missing_cols:
                QMessageBox.warning(
                    self, "表头错误",
                    f"导入文件缺少以下列：{'、'.join(missing_cols)}\n"
                    f"请检查文件表头是否包含：{'、'.join(std_columns)}\n\n"
                    f"当前文件列：{'、'.join(str(c) for c in df.columns)}")
                return

            def _cell(row, std):
                """取单元格文本，空值统一返回空字符串"""
                val = row.get(col_of[std], '')
                return '' if val is None else str(val).strip()

            # Clear existing compounds from the table for fresh import
            self.compound_table.setRowCount(0)

            imported_count = 0
            for index, row in df.iterrows():
                # Create a QTableWidgetItem for each cell and set its value
                cas_number = _cell(row, 'CAS号')
                rt_str = _cell(row, '组分RT')
                name_en = _cell(row, '化合物名称')
                name_cn = _cell(row, '中文名')
                match_factor_str = _cell(row, '匹配因子')
                molecular_formula = _cell(row, '分子式')
                relative_content_str = _cell(row, '相对含量mg/L')

                # 整行全空则跳过
                if not any([cas_number, rt_str, name_en, name_cn,
                            match_factor_str, molecular_formula,
                            relative_content_str]):
                    continue

                current_row = self.compound_table.rowCount()
                self.compound_table.insertRow(current_row)

                # Populate table cells with data from DataFrame columns using correct indices
                # Indices correspond to the table headers:
                # CAS号(0), 组分RT(1), 化合物名称(2), 中文名(3), 匹配因子(4), 分子式(5), 相对含量mg/L(6)
                
                # CAS号 (Column 0)
                self.compound_table.setItem(current_row, 0, QTableWidgetItem(cas_number))
                
                # 组分RT (Column 1) - Use NumericTableWidgetItem for sorting
                rt_item = NumericTableWidgetItem(rt_str)
                try:
                    rt_value = float(rt_str) if rt_str else -1.0
                except ValueError:
                    rt_value = -1.0
                rt_item.setData(Qt.ItemDataRole.UserRole, rt_value)
                self.compound_table.setItem(current_row, 1, rt_item)

                # 化合物名称 (Column 2)
                self.compound_table.setItem(current_row, 2, QTableWidgetItem(name_en))

                # 中文名 (Column 3)
                self.compound_table.setItem(current_row, 3, QTableWidgetItem(name_cn))
                
                # 匹配因子 (Column 4) - Use NumericTableWidgetItem for sorting
                match_factor_item = NumericTableWidgetItem(match_factor_str)
                try:
                    match_factor_value = float(match_factor_str) if match_factor_str else -1.0
                except ValueError:
                     match_factor_value = -1.0
                match_factor_item.setData(Qt.ItemDataRole.UserRole, match_factor_value)
                self.compound_table.setItem(current_row, 4, match_factor_item)

                # 分子式 (Column 5)
                self.compound_table.setItem(current_row, 5, QTableWidgetItem(molecular_formula))

                # 相对含量mg/L (Column 6) - Use NumericTableWidgetItem for sorting
                relative_content_item = NumericTableWidgetItem(relative_content_str)
                try:
                    relative_content_value = float(relative_content_str) if relative_content_str else -1.0
                except ValueError:
                    relative_content_value = -1.0
                relative_content_item.setData(Qt.ItemDataRole.UserRole, relative_content_value)
                self.compound_table.setItem(current_row, 6, relative_content_item)

                imported_count += 1

            self.compound_table.resizeColumnsToContents()
            self.compound_table.resizeRowsToContents()
            QMessageBox.information(self, "导入完成", f"成功导入 {imported_count} 条化合物记录到表格！\n请点击保存将数据写入数据库。")

        except Exception as e:
            QMessageBox.critical(self, "导入失败", f"导入过程中发生错误：{str(e)}")
            print(f"Compounds import to table failed: {e}")

    def filter_compounds_table(self, text):
        """根据输入的文本筛选化合物表格"""
        keyword = text.strip().lower()
        for row in range(self.compound_table.rowCount()):
            # Get item texts from relevant columns (CAS, Name EN, Name CN)
            cas_item = self.compound_table.item(row, 0)
            name_en_item = self.compound_table.item(row, 2)
            name_cn_item = self.compound_table.item(row, 3)
            
            # Check if keyword is in any of the relevant columns
            is_match = False
            if cas_item and keyword in cas_item.text().lower():
                is_match = True
            elif name_en_item and keyword in name_en_item.text().lower():
                is_match = True
            elif name_cn_item and keyword in name_cn_item.text().lower():
                is_match = True
            
            # Hide or show the row based on match
            self.compound_table.setRowHidden(row, not is_match)


class GCMSResultDialog(QDialog):
    def __init__(self, parent=None, analysis=None, session=None):
        super().__init__(parent)
        self.analysis = analysis
        self.session = session
        self.setWindowTitle("GC-MS 分析结果")
        self.setMinimumSize(800, 600)
        
        # 创建主布局
        layout = QVBoxLayout()
        
        # 创建信息显示区域
        info_text = f"分析编号: {analysis.number}\n"
        info_text += f"分析名称: {analysis.name}\n"
        info_text += f"仪器参数: {analysis.instrument_params or '未记录'}\n"
        info_text += f"供应商: {analysis.supplier or '未知'}\n"
        info_text += f"香精构思: {analysis.perfume_idea or '未记录'}\n"
        info_text += f"分析时间: {analysis.analysis_time.strftime('%Y-%m-%d %H:%M:%S') if analysis.analysis_time else '未记录'}\n"
        
        info_label = QLabel(info_text)
        info_label.setStyleSheet("QLabel { padding: 10px; }")
        layout.addWidget(info_label)

        # --- Compounds Section ---
        compounds_group = QWidget()
        compounds_layout = QVBoxLayout(compounds_group)
        compounds_layout.addWidget(QLabel("化合物信息"))

        # Compound table to display compounds
        self.compound_table = QTableWidget()
        self.compound_table.setColumnCount(8)
        self.compound_table.setHorizontalHeaderLabels([
            "ID", "CAS号", "组分RT", "化合物名称", "中文名", "匹配因子", "分子式", "相对含量mg/L"
        ])
        self.compound_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        compounds_layout.addWidget(self.compound_table)

        # Load compounds data
        self._load_compounds_data()

        main_layout.addWidget(compounds_group)

        # --- Close Button ---
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.accept)
        main_layout.addWidget(close_btn)

        self.setLayout(main_layout)

    def _load_compounds_data(self):
        """加载化合物数据到表格"""
        self.compound_table.setRowCount(0)  # 清空表格
        if self.analysis and self.analysis.compounds:
            for compound in self.analysis.compounds:
                row = self.compound_table.rowCount()
                self.compound_table.insertRow(row)
                self.compound_table.setItem(row, 0, QTableWidgetItem(str(compound.rt or "")))
                self.compound_table.setItem(row, 1, QTableWidgetItem(compound.cas or ""))
                self.compound_table.setItem(row, 2, QTableWidgetItem(compound.name_en or ""))
                self.compound_table.setItem(row, 3, QTableWidgetItem(compound.name_cn or ""))
                self.compound_table.setItem(row, 4, QTableWidgetItem(str(compound.match_factor or "")))
                self.compound_table.setItem(row, 5, QTableWidgetItem(compound.formula or ""))
                self.compound_table.setItem(row, 6, QTableWidgetItem(str(compound.relative_content or "")))

    def accept(self):
        super().accept()

    def reject(self):
        super().reject()


class GCMSAnalysisFuncDialog(QDialog):
    def __init__(self, parent, session):
        super().__init__(parent)
        self.setWindowTitle("GC-MS分析功能")
        self.resize(800, 600)
        self.session = session
        layout = QVBoxLayout()

        # Add analysis selection area with checkboxes
        self.analysis_list_label = QLabel("请勾选要分析的GC-MS记录：")
        layout.addWidget(self.analysis_list_label)

        self.analysis_list_widget = QListWidget()
        layout.addWidget(self.analysis_list_widget)

        # Load available analyses with checkboxes
        self.load_analyses()

        tabs = QTabWidget()

        # --- 特征物质筛查Tab ---
        self.feature_tab = QWidget()
        feature_layout = QVBoxLayout()
        feature_layout.addWidget(QLabel("特征物质筛查功能（可自定义筛查规则和展示结果）"))
        self.feature_btn = QPushButton("执行筛查并可视化")
        self.feature_btn.clicked.connect(self.run_feature_screening)
        feature_layout.addWidget(self.feature_btn)

        # Export buttons for feature screening
        feature_export_layout = QHBoxLayout()
        self.feature_export_img_btn = QPushButton("导出图片")
        self.feature_export_data_btn = QPushButton("导出数据")
        feature_export_layout.addWidget(self.feature_export_img_btn)
        feature_export_layout.addWidget(self.feature_export_data_btn)
        feature_layout.addLayout(feature_export_layout)

        self.feature_fig_label = QLabel()
        feature_layout.addWidget(self.feature_fig_label)

        self.feature_table = QTableWidget()
        self.feature_table.setMinimumHeight(150)
        self.feature_table.setAlternatingRowColors(True)
        feature_layout.addWidget(self.feature_table)
        self.feature_tab.setLayout(feature_layout)
        tabs.addTab(self.feature_tab, "特征物质筛查")

        # --- 差异物质分析Tab ---
        self.diff_tab = QWidget()
        diff_layout = QVBoxLayout()
        diff_layout.addWidget(QLabel("差异物质分析功能（可选择两组分析结果进行差异分析）"))
        self.diff_btn = QPushButton("执行差异分析并可视化")
        self.diff_btn.clicked.connect(self.run_diff_analysis)
        diff_layout.addWidget(self.diff_btn)

        # Export buttons for differential analysis
        diff_export_layout = QHBoxLayout()
        self.diff_export_img_btn = QPushButton("导出图片")
        self.diff_export_data_btn = QPushButton("导出数据")
        diff_export_layout.addWidget(self.diff_export_img_btn)
        diff_export_layout.addWidget(self.diff_export_data_btn)
        diff_layout.addLayout(diff_export_layout)

        self.diff_fig_label = QLabel()
        diff_layout.addWidget(self.diff_fig_label)

        self.diff_table = QTableWidget()
        self.diff_table.setMinimumHeight(150)
        self.diff_table.setAlternatingRowColors(True)
        diff_layout.addWidget(self.diff_table)
        self.diff_tab.setLayout(diff_layout)
        tabs.addTab(self.diff_tab, "差异物质分析")

        # --- 可视化Tab ---
        self.vis_tab = QWidget()
        vis_layout = QVBoxLayout()
        vis_layout.addWidget(QLabel("可视化功能（热图 / PCA 降维）"))

        vis_btn_layout = QHBoxLayout()
        self.vis_btn = QPushButton("显示热图")
        self.vis_btn.clicked.connect(self.show_heatmap)
        self.pca_btn = QPushButton("PCA 分析")
        self.pca_btn.clicked.connect(self.run_pca_analysis)
        vis_btn_layout.addWidget(self.vis_btn)
        vis_btn_layout.addWidget(self.pca_btn)
        vis_layout.addLayout(vis_btn_layout)

        # Export buttons for visualization
        vis_export_layout = QHBoxLayout()
        self.vis_export_img_btn = QPushButton("导出图片")
        self.vis_export_data_btn = QPushButton("导出数据")
        vis_export_layout.addWidget(self.vis_export_img_btn)
        vis_export_layout.addWidget(self.vis_export_data_btn)
        vis_layout.addLayout(vis_export_layout)

        self.vis_fig_label = QLabel()
        vis_layout.addWidget(self.vis_fig_label)

        self.vis_table = QTableWidget()
        self.vis_table.setMinimumHeight(150)
        self.vis_table.setAlternatingRowColors(True)
        vis_layout.addWidget(self.vis_table)
        self.vis_tab.setLayout(vis_layout)
        tabs.addTab(self.vis_tab, "可视化")

        # Corrected layout setting for tabs
        tab_layout = QVBoxLayout()
        tab_layout.addWidget(tabs)
        layout.addLayout(tab_layout)


        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn)
        self.setLayout(layout)

        # Connect export button signals
        self.feature_export_img_btn.clicked.connect(lambda: self.export_analysis_image(self.feature_fig_label.pixmap(), "特征筛查图片"))
        self.feature_export_data_btn.clicked.connect(lambda: self.export_analysis_data("feature"))

        self.diff_export_img_btn.clicked.connect(lambda: self.export_analysis_image(self.diff_fig_label.pixmap(), "差异分析图片"))
        self.diff_export_data_btn.clicked.connect(lambda: self.export_analysis_data("differential"))

        self.vis_export_img_btn.clicked.connect(lambda: self.export_analysis_image(self.vis_fig_label.pixmap(), "可视化图片"))
        self.vis_export_data_btn.clicked.connect(lambda: self.export_analysis_data("visualization"))


        # Variables to store data for export
        self._feature_result_data = None
        self._differential_result_data = None
        self._visualization_result_data = None
        self._differential_matrix = None


    def load_analyses(self):
        """Load available GC-MS analyses into the list widget with checkboxes"""
        self.analysis_list_widget.clear()
        analyses = self.session.query(GCMSAnalysis).all()
        for ana in analyses:
            item = QListWidgetItem(f"{ana.number} - {ana.name}")
            item.setData(Qt.ItemDataRole.UserRole, ana.id) # Store analysis ID in item data
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable) # Make item checkable
            item.setCheckState(Qt.CheckState.Unchecked) # Set initial state to unchecked
            self.analysis_list_widget.addItem(item)

    def get_selected_analysis_ids(self):
        """Get the IDs of the checked analyses"""
        selected_ids = []
        for i in range(self.analysis_list_widget.count()):
            item = self.analysis_list_widget.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                analysis_id = item.data(Qt.ItemDataRole.UserRole)
                if analysis_id is not None:
                    selected_ids.append(analysis_id)
        return selected_ids

    # --- Framework for More Complex Analyses ---

    def _analysis_label_map(self, analysis_ids):
        """分析ID -> 易读标签（编号 - 样品名称），避免图上只显示数字ID"""
        labels = {}
        for aid in analysis_ids:
            ana = self.session.query(GCMSAnalysis).get(aid)
            if ana is None:
                labels[aid] = f"#{aid}"
                continue
            num = (ana.number or '').strip()
            name = (ana.name or '').strip()
            labels[aid] = f"{num} {name}".strip() or f"#{aid}"
        return labels

    def _prepare_compound_data(self, selected_analysis_ids):
        """Query compound data for selected analyses and return as a pandas DataFrame."""
        if not PANDAS_AVAILABLE:
            QMessageBox.warning(self, "功能受限", "pandas不可用，无法进行数据分析。")
            return pd.DataFrame()

        if not selected_analysis_ids:
            return pd.DataFrame()
        compounds = self.session.query(GCMSCompound).filter(
            GCMSCompound.analysis_id.in_(selected_analysis_ids)).all()

        labels = self._analysis_label_map(selected_analysis_ids)
        data = []
        for c in compounds:
            # 优先用相对含量，缺失时回退到绝对含量，保证矩阵有数值
            value = c.relative_content
            if value is None:
                value = c.content
            try:
                value = float(value) if value is not None else 0.0
            except (TypeError, ValueError):
                value = 0.0
            data.append({
                'analysis_id': c.analysis_id,
                'analysis_name': labels.get(c.analysis_id, f"#{c.analysis_id}"),
                'compound_name': (c.name_cn or c.name_en or '').strip() or '未命名',
                'relative_content': value
            })
        df = pd.DataFrame(data)
        return df

    def _perform_feature_analysis(self, df):
        """特征物质筛查：按总相对含量排序，并给出出现样品数与均值

        返回 DataFrame，列为：化合物 / 总相对含量 / 出现样品数 / 平均相对含量
        """
        print("Performing feature analysis...")
        if df.empty:
            print("DataFrame is empty for feature analysis.")
            self._feature_result_data = pd.DataFrame()
            return pd.DataFrame()

        grouped = df.groupby('compound_name')['relative_content'].agg(
            total='sum', mean='mean', samples='count')
        grouped = grouped.sort_values('total', ascending=False)
        result = grouped.reset_index().rename(columns={
            'compound_name': '化合物',
            'total': '总相对含量',
            'mean': '平均相对含量',
            'samples': '出现样品数',
        })
        result['平均相对含量'] = result['平均相对含量'].round(4)
        result['总相对含量'] = result['总相对含量'].round(4)

        self._feature_result_data = result
        return result

    def _perform_differential_analysis(self, df):
        """构建 化合物 x 样品 的相对含量矩阵，并计算差异指标"""
        print("Performing differential analysis...")
        if df.empty:
            print("DataFrame is empty for differential analysis.")
            self._differential_result_data = pd.DataFrame()
            return pd.DataFrame()

        matrix = df.pivot_table(index='analysis_name', columns='compound_name',
                                values='relative_content', aggfunc='sum',
                                fill_value=0)
        # 化合物为行、样品为列，更符合调香分析阅读习惯
        matrix = matrix.T
        matrix = matrix.loc[:, [c for c in matrix.columns]]
        self._differential_matrix = matrix
        self._differential_result_data = matrix
        return matrix

    def _differential_metrics(self, matrix):
        """基于 化合物 x 样品 矩阵计算差异指标

        两组样品时给出 B/A 差异倍数与 log2FC；
        多组时给出极差与变异系数 CV。
        """
        if matrix is None or matrix.empty:
            return pd.DataFrame()

        cols = list(matrix.columns)
        stats = pd.DataFrame({
            '均值': matrix.mean(axis=1),
            '最大值': matrix.max(axis=1),
            '最小值': matrix.min(axis=1),
        })
        stats['极差'] = stats['最大值'] - stats['最小值']

        # 用 where() 把 0 置为 NaN 而非 pd.NA：
        # pd.NA 无法再 astype(float)，会导致导出与计算报错
        mean_nonzero = stats['均值'].abs().where(stats['均值'] != 0)
        stats['变异系数CV'] = matrix.std(axis=1) / mean_nonzero
        min_nonzero = stats['最小值'].where(stats['最小值'] != 0)
        stats['最大值/最小值'] = stats['最大值'] / min_nonzero

        if len(cols) == 2:
            a, b = cols[0], cols[1]
            stats[a] = matrix[a]
            stats[b] = matrix[b]
            denom = matrix[a].where(matrix[a] != 0)
            ratio = matrix[b] / denom
            stats['差异倍数(B/A)'] = ratio
            stats['log2FC'] = [self._safe_log2(v) for v in ratio]

        result = stats.round(4).sort_values('极差', ascending=False)
        result.index.name = '化合物'
        return result.reset_index()

    @staticmethod
    def _safe_log2(value):
        """安全的 log2 计算，非正数或缺失返回 NaN"""
        try:
            if value is None or pd.isna(value) or float(value) <= 0:
                return float('nan')
            return math.log2(float(value))
        except Exception:
            return float('nan')


    def _perform_pca_analysis(self, df):
        """Framework for PCA analysis."""
        print("Performing PCA analysis...")
        if df.empty:
            print("DataFrame is empty for PCA analysis.")
            self._visualization_result_data = pd.DataFrame() # Store empty for export
            return pd.DataFrame() # Return empty DataFrame instead of None

        # 样品为行、化合物为列
        pivot_df = df.pivot_table(index='analysis_name', columns='compound_name',
                                  values='relative_content', aggfunc='sum',
                                  fill_value=0)

        if pivot_df.empty:
             print("Pivoted DataFrame is empty for PCA analysis.")
             return pd.DataFrame()

        # 至少需要 2 个样品与 2 种化合物才有降维意义
        if pivot_df.shape[0] < 2 or pivot_df.shape[1] < 2:
             print("Not enough data for meaningful PCA.")
             return pd.DataFrame()

        try:
            from sklearn.decomposition import PCA
            # 主成分数受样本数与特征数共同限制，避免越界报错
            n_comp = min(2, pivot_df.shape[0], pivot_df.shape[1])
            pca = PCA(n_components=n_comp)
            components = pca.fit_transform(pivot_df)

            cols = ['principal_component_1', 'principal_component_2'][:n_comp]
            pca_df = pd.DataFrame(data=components, columns=cols,
                                  index=pivot_df.index)
            if n_comp == 1:
                pca_df['principal_component_2'] = 0.0
            # 样品名放在索引上并命名，reset_index() 后即为 analysis_name 列，
            # 避免索引名与列名重复导致 insert 冲突
            pca_df.index.name = 'analysis_name'

            self._visualization_result_data = pca_df
            return pca_df

        except ImportError:
            print("sklearn not installed. Cannot perform PCA.")
            return pd.DataFrame()
        except Exception as e:
            print(f"Error during PCA analysis: {e}")
            return pd.DataFrame()


    # ---------- 绘图与展示辅助 ----------
    def _fig_to_pixmap(self, fig):
        """matplotlib figure -> QPixmap（失败返回 None）"""
        from PyQt6.QtGui import QPixmap
        try:
            buf = BytesIO()
            fig.savefig(buf, format='png', dpi=150, bbox_inches='tight')
            plt.close(fig)
            buf.seek(0)
            pix = QPixmap()
            pix.loadFromData(buf.read())
            return pix if not pix.isNull() else None
        except Exception as e:
            print(f"图表转换失败: {e}")
            try:
                plt.close(fig)
            except Exception:
                pass
            return None

    @staticmethod
    def _set_scaled_pixmap(label, pix, max_w, max_h):
        """按最大边缩放后显示在 QLabel 上，保持宽高比"""
        label.setPixmap(pix.scaled(
            max_w, max_h,
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation))
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)

    def _draw_heatmap(self, matrix, ax, title):
        """绘制热图；seaborn 不可用时回退到 matplotlib imshow"""
        if sns is not None:
            sns.heatmap(matrix, ax=ax, cmap='YlGnBu',
                        cbar_kws={'label': '相对含量'})
        else:
            im = ax.imshow(matrix.values, aspect='auto', cmap='YlGnBu')
            ax.set_xticks(range(len(matrix.columns)))
            ax.set_xticklabels(matrix.columns, rotation=45, ha='right')
            ax.set_yticks(range(len(matrix.index)))
            ax.set_yticklabels(matrix.index)
            plt.colorbar(im, ax=ax, label='相对含量')
        ax.set_title(title)
        ax.set_xlabel('样品')
        ax.set_ylabel('化合物名称')

    @staticmethod
    def _fill_table(table, df):
        """将 DataFrame 填充到 QTableWidget（含表头）"""
        table.clear()
        if df is None or df.empty:
            table.setRowCount(0)
            table.setColumnCount(0)
            return
        table.setColumnCount(len(df.columns))
        table.setHorizontalHeaderLabels([str(c) for c in df.columns])
        table.setRowCount(len(df))
        for r in range(len(df)):
            for c in range(len(df.columns)):
                value = df.iat[r, c]
                if value is None or (isinstance(value, float) and pd.isna(value)):
                    text = ''
                elif isinstance(value, float):
                    text = f"{value:,.4f}"
                else:
                    text = str(value)
                table.setItem(r, c, QTableWidgetItem(text))
        table.resizeColumnsToContents()

    # ---------- 特征物质筛查 ----------
    def run_feature_screening(self):
        selected_ids = self.get_selected_analysis_ids()
        if not selected_ids:
            QMessageBox.warning(self, "提示", "请先勾选要分析的GC-MS记录！")
            return
        df = self._prepare_compound_data(selected_ids)
        if df.empty:
            self.feature_fig_label.setText("选定的记录无化合物数据")
            self._feature_result_data = pd.DataFrame()
            return

        feature_results = self._perform_feature_analysis(df)
        if feature_results.empty:
            self.feature_fig_label.setText("特征筛查无结果")
            self._feature_result_data = pd.DataFrame()
            return

        self._fill_table(self.feature_table, feature_results)

        if not MATPLOTLIB_AVAILABLE:
            self.feature_fig_label.setText("matplotlib不可用，无法生成图表")
            return

        try:
            top = feature_results.head(15).iloc[::-1]  # 反转让最大值在最上方
            fig, ax = plt.subplots(figsize=(7, max(3.0, len(top) * 0.35)))
            ax.barh(top['化合物'], top['总相对含量'], color='#2E86AB')
            ax.set_xlabel('总相对含量（所选样品合计）')
            ax.set_title('特征物质 Top15（按总相对含量）')
            plt.tight_layout()
            pix = self._fig_to_pixmap(fig)
            if pix is None:
                self.feature_fig_label.setText("图表生成失败")
            else:
                self._set_scaled_pixmap(self.feature_fig_label, pix, 720, 500)
        except Exception as e:
            self.feature_fig_label.setText(f"可视化失败：{e}")
            print(f"Feature screening visualization failed: {e}")

    # ---------- 差异物质分析 ----------
    def run_diff_analysis(self):
        selected_ids = self.get_selected_analysis_ids()
        if len(selected_ids) < 2:
            QMessageBox.warning(self, "提示", "差异分析至少需要勾选两组GC-MS记录！")
            return
        df = self._prepare_compound_data(selected_ids)
        if df.empty:
            self.diff_fig_label.setText("选定的记录无化合物数据")
            self._differential_result_data = pd.DataFrame()
            self._differential_matrix = None
            return
        matrix = self._perform_differential_analysis(df)

        if matrix.empty:
            self.diff_fig_label.setText("差异分析无结果")
            self._differential_result_data = pd.DataFrame()
            self._differential_matrix = None
            return

        # 差异指标表：两组时给出差异倍数与 log2FC
        metrics = self._differential_metrics(matrix)
        self._fill_table(self.diff_table, metrics)
        self._differential_result_data = metrics if not metrics.empty else matrix

        if not MATPLOTLIB_AVAILABLE:
            self.diff_fig_label.setText("matplotlib不可用，无法生成图表")
            return

        try:
            n_rows = max(3, len(matrix) * 0.3 + 2)
            fig, ax = plt.subplots(figsize=(8, n_rows))
            self._draw_heatmap(
                matrix, ax,
                '化合物 x 样品 相对含量热图（按极差排序前20）')
            plt.tight_layout()
            pix = self._fig_to_pixmap(fig)
            if pix is None:
                self.diff_fig_label.setText("图表生成失败")
            else:
                self._set_scaled_pixmap(self.diff_fig_label, pix, 720, 520)
        except Exception as e:
            self.diff_fig_label.setText(f"可视化失败：{e}")
            print(f"Differential analysis visualization failed: {e}")

    # ---------- 可视化：热图 ----------
    def show_heatmap(self):
        selected_ids = self.get_selected_analysis_ids()
        if not selected_ids:
            QMessageBox.warning(self, "提示", "请先勾选要分析的GC-MS记录！")
            return
        df = self._prepare_compound_data(selected_ids)
        if df.empty:
            self.vis_fig_label.setText("选定的记录无化合物数据")
            self._visualization_result_data = pd.DataFrame()
            return

        matrix = self._perform_differential_analysis(df)
        if matrix.empty:
            self.vis_fig_label.setText("可视化数据准备失败")
            self._visualization_result_data = pd.DataFrame()
            return

        self._visualization_result_data = matrix
        self._fill_table(self.vis_table, matrix.reset_index())

        if not MATPLOTLIB_AVAILABLE:
            self.vis_fig_label.setText("matplotlib不可用，无法生成图表")
            return

        try:
            # 化合物过多时只画差异最大的前 25 个，保证可读性
            plot_df = matrix
            if len(matrix) > 25:
                spread = (matrix.max(axis=1) - matrix.min(axis=1))
                plot_df = matrix.loc[spread.sort_values(ascending=False).head(25).index]
            fig, ax = plt.subplots(figsize=(8, max(3, len(plot_df) * 0.3 + 2)))
            self._draw_heatmap(plot_df, ax, '样品-化合物相对含量热图')
            plt.tight_layout()
            pix = self._fig_to_pixmap(fig)
            if pix is None:
                self.vis_fig_label.setText("图表生成失败")
            else:
                self._set_scaled_pixmap(self.vis_fig_label, pix, 720, 500)
        except Exception as e:
            self.vis_fig_label.setText(f"可视化失败：{e}")
            print(f"Visualization (Heatmap) failed: {e}")

    # ---------- 可视化：PCA ----------
    def run_pca_analysis(self):
        """对勾选的GC-MS记录做PCA降维并绘制散点图"""
        selected_ids = self.get_selected_analysis_ids()
        if len(selected_ids) < 2:
            QMessageBox.warning(self, "提示", "PCA分析至少需要勾选两组GC-MS记录！")
            return
        df = self._prepare_compound_data(selected_ids)
        if df.empty:
            self.vis_fig_label.setText("选定的记录无化合物数据")
            return

        pca_df = self._perform_pca_analysis(df)
        if pca_df is None or pca_df.empty:
            self.vis_fig_label.setText(
                "PCA分析无结果：需要至少2组样品与2种化合物，且需安装 scikit-learn")
            return

        self._visualization_result_data = pca_df

        try:
            from sklearn.decomposition import PCA
            pivot_df = df.pivot_table(index='analysis_name',
                                      columns='compound_name',
                                      values='relative_content',
                                      aggfunc='sum', fill_value=0)
            pca = PCA(n_components=2)
            pca.fit(pivot_df)
            explained = pca.explained_variance_ratio_ * 100
        except Exception:
            explained = None

        self._fill_table(self.vis_table, pca_df.reset_index())

        if not MATPLOTLIB_AVAILABLE:
            self.vis_fig_label.setText("matplotlib不可用，无法生成图表")
            return

        try:
            fig, ax = plt.subplots(figsize=(6.5, 5))
            ax.scatter(pca_df['principal_component_1'],
                       pca_df['principal_component_2'],
                       s=90, color='#2E86AB', edgecolors='white', zorder=3)
            for label, r in pca_df.iterrows():
                ax.annotate(str(label),
                            (r['principal_component_1'],
                             r['principal_component_2']),
                            fontsize=9, xytext=(6, 4),
                            textcoords='offset points')
            ax.axhline(0, color='#cccccc', linewidth=0.8, zorder=1)
            ax.axvline(0, color='#cccccc', linewidth=0.8, zorder=1)
            xlabel = '主成分 1' + (f" ({explained[0]:.1f}%)" if explained is not None else '')
            ylabel = '主成分 2' + (f" ({explained[1]:.1f}%)" if explained is not None else '')
            ax.set_xlabel(xlabel)
            ax.set_ylabel(ylabel)
            ax.set_title('GC-MS 样品 PCA 降维')
            ax.grid(True, linestyle='--', alpha=0.3)
            plt.tight_layout()
            pix = self._fig_to_pixmap(fig)
            if pix is None:
                self.vis_fig_label.setText("图表生成失败")
            else:
                self._set_scaled_pixmap(self.vis_fig_label, pix, 720, 500)
        except Exception as e:
            self.vis_fig_label.setText(f"可视化失败：{e}")
            print(f"PCA visualization failed: {e}")


    # Keep export_analysis_image and export_analysis_data methods as they are useful for analysis functions
    def export_analysis_image(self, pixmap, base_filename):
        """Export the analysis result image."""
        if pixmap is None or pixmap.isNull():
            QMessageBox.warning(self, "导出失败", "没有可导出的图片。")
            return

        from PyQt6.QtWidgets import QFileDialog
        file_path, _ = QFileDialog.getSaveFileName(self, "导出图片", f"{base_filename}.png", "PNG 图片 (*.png);;JPEG 图片 (*.jpg)")
        if file_path:
            if pixmap.save(file_path):
                QMessageBox.information(self, "导出成功", f"图片已导出到：\n{file_path}")
            else:
                QMessageBox.critical(self, "导出失败", "保存图片失败。")

    def export_analysis_data(self, analysis_type):
        """Export the analysis result data to Excel or CSV file."""
        data_to_export = None
        filename = ""
        if analysis_type == "feature":
            data_to_export = self._feature_result_data
            filename = "特征筛查结果数据"
        elif analysis_type == "differential":
            data_to_export = self._differential_result_data
            filename = "差异分析结果数据"
        elif analysis_type == "visualization":
            data_to_export = self._visualization_result_data
            filename = "可视化结果数据"

        if data_to_export is None or data_to_export.empty:
            QMessageBox.warning(self, "导出失败", "没有可导出的数据。")
            return

        from PyQt6.QtWidgets import QFileDialog

        file_path, file_type = QFileDialog.getSaveFileName(self, "导出数据", f"{filename}.xlsx", "Excel 文件 (*.xlsx);;CSV 文件 (*.csv)") # Fixed file type string
        if not file_path:
            return
        try:
            if file_path.endswith('.csv'):
                # Adjust index based on how the data is stored (True if index is meaningful, False otherwise)
                data_to_export.to_csv(file_path, index=True if analysis_type != "feature" else False, encoding='utf-8-sig')
            else:
                # Adjust index based on how the data is stored
                data_to_export.to_excel(file_path, index=True if analysis_type != "feature" else False)
            QMessageBox.information(self, "导出成功", f"数据已导出到：\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", f"导出过程中发生错误：{str(e)}")

# Removed LoginDialog and RegisterDialog classes


class SupplierDialog(QDialog):
    """供应商添加/编辑对话框"""
    def __init__(self, parent=None, supplier=None):
        super().__init__(parent)
        self.setWindowTitle("编辑供应商" if supplier else "添加供应商")
        self.setMinimumSize(400, 300)
        self.supplier = supplier
        
        layout = QVBoxLayout()
        
        # 创建表单
        form = QFormLayout()
        
        self.name_edit = QLineEdit()
        self.contact_person_edit = QLineEdit()
        self.phone_edit = QLineEdit()
        self.email_edit = QLineEdit()
        self.address_edit = QLineEdit()
        
        form.addRow("供应商名称*", self.name_edit)
        form.addRow("联系人", self.contact_person_edit)
        form.addRow("电话", self.phone_edit)
        form.addRow("邮箱", self.email_edit)
        form.addRow("地址", self.address_edit)
        
        # 如果是编辑模式，填充现有数据
        if supplier:
            self.name_edit.setText(supplier.name or "")
            self.contact_person_edit.setText(supplier.contact_person or "")
            self.phone_edit.setText(supplier.phone or "")
            self.email_edit.setText(supplier.email or "")
            self.address_edit.setText(supplier.address or "")
        
        layout.addLayout(form)
        
        # 按钮
        button_layout = QHBoxLayout()
        save_btn = QPushButton("保存")
        save_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        
        button_layout.addWidget(save_btn)
        button_layout.addWidget(cancel_btn)
        
        layout.addLayout(button_layout)
        self.setLayout(layout)
    
    def get_data(self):
        """获取表单数据"""
        return {
            'name': self.name_edit.text().strip(),
            'contact_person': self.contact_person_edit.text().strip(),
            'phone': self.phone_edit.text().strip(),
            'email': self.email_edit.text().strip(),
            'address': self.address_edit.text().strip()
        }
    
    def accept(self):
        """验证并保存"""
        if not self.name_edit.text().strip():
            QMessageBox.warning(self, "警告", "供应商名称不能为空！")
            return
        super().accept()


class FormulaUsageDialog(QDialog):
    """配方使用对话框 - 完整版本"""
    def __init__(self, parent, formula, session):
        super().__init__(parent)
        self.formula = formula
        self.session = session
        self.setWindowTitle(f"使用配方 - {formula.name}")
        self.setMinimumSize(800, 600)
        self.resize(900, 700)
        
        # 主布局
        main_layout = QVBoxLayout()
        
        # 配方基本信息区域
        info_group = QGroupBox("📋 配方信息")
        info_layout = QFormLayout()
        info_layout.addRow("配方名称:", QLabel(formula.name))
        info_layout.addRow("编号:", QLabel(formula.number or ""))
        info_layout.addRow("版本:", QLabel(formula.version or ""))
        info_layout.addRow("创建人:", QLabel(formula.creator or ""))
        info_layout.addRow("描述:", QLabel(formula.description or ""))
        info_group.setLayout(info_layout)
        main_layout.addWidget(info_group)
        
        # 使用参数设置区域
        params_group = QGroupBox("⚙️ 使用参数")
        params_layout = QFormLayout()
        
        # 生产倍数（核心参数）
        self.multiplier_spin = QDoubleSpinBox()
        self.multiplier_spin.setRange(0.1, 100.0)
        self.multiplier_spin.setDecimals(2)
        self.multiplier_spin.setValue(1.0)
        self.multiplier_spin.valueChanged.connect(self.update_stock_check)
        
        # 添加说明标签
        multiplier_help = QLabel("💡 生产倍数：1.0=按配方标准用量，2.0=双倍生产，0.5=半量生产")
        multiplier_help.setStyleSheet("color: #666; font-size: 11px; padding: 2px;")
        params_layout.addRow("生产倍数:", self.multiplier_spin)
        params_layout.addRow("", multiplier_help)
        
        # 批次号（必填）
        self.batch_edit = QLineEdit()
        self.batch_edit.setPlaceholderText("如：20241201-001...")
        params_layout.addRow("批次号:", self.batch_edit)
        
        # 操作人（必填）
        self.operator_edit = QLineEdit()
        self.operator_edit.setPlaceholderText("请输入操作人姓名...")
        params_layout.addRow("操作人:", self.operator_edit)
        
        # 调制时间（自动填充当前时间，可修改）
        self.production_time_edit = QLineEdit()
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.production_time_edit.setText(current_time)
        params_layout.addRow("调制时间:", self.production_time_edit)
        
        # 预期用途（新增）
        self.purpose_combo = QComboBox()
        self.purpose_combo.addItems([
            "产品生产", "样品试制", "质量检验", "研发测试", "客户样品", "库存补充", "其他"
        ])
        params_layout.addRow("使用用途:", self.purpose_combo)
        
        # 质量标准（新增）
        self.quality_combo = QComboBox()
        self.quality_combo.addItems([
            "标准品质", "高品质", "测试品质", "样品级别"
        ])
        params_layout.addRow("质量标准:", self.quality_combo)
        
        # 备注
        self.notes_edit = QLineEdit()
        self.notes_edit.setPlaceholderText("可选备注信息（如特殊要求、注意事项等）...")
        params_layout.addRow("备注:", self.notes_edit)
        
        params_group.setLayout(params_layout)
        main_layout.addWidget(params_group)
        
        # 原料需求与库存状态表格
        stock_group = QGroupBox("📦 原料需求与库存状态")
        stock_layout = QVBoxLayout()
        
        # 状态标签
        self.status_label = QLabel("💡 请设置使用参数后查看库存状态")
        self.status_label.setStyleSheet("""
            QLabel {
                background-color: #e3f2fd;
                border: 1px solid #2196f3;
                border-radius: 4px;
                padding: 8px;
                color: #1976d2;
            }
        """)
        stock_layout.addWidget(self.status_label)
        
        # 库存状态表格
        self.stock_table = QTableWidget()
        self.stock_table.setColumnCount(8)
        self.stock_table.setHorizontalHeaderLabels([
            "原料名称", "配方比例", "配方用量(g)", "稀释倍数", "实际消耗(g)", "当前库存(g)", "库存状态", "剩余库存(g)"
        ])
        
        # 设置表格样式
        self.stock_table.setAlternatingRowColors(True)
        self.stock_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.stock_table.horizontalHeader().setStretchLastSection(True)
        
        # 设置列宽
        self.stock_table.setColumnWidth(0, 120)  # 原料名称
        self.stock_table.setColumnWidth(1, 80)   # 配方比例
        self.stock_table.setColumnWidth(2, 90)   # 配方用量
        self.stock_table.setColumnWidth(3, 70)   # 稀释倍数
        self.stock_table.setColumnWidth(4, 90)   # 实际消耗
        self.stock_table.setColumnWidth(5, 90)   # 当前库存
        self.stock_table.setColumnWidth(6, 80)   # 库存状态
        self.stock_table.setColumnWidth(7, 90)   # 剩余库存
        
        stock_layout.addWidget(self.stock_table)
        stock_group.setLayout(stock_layout)
        main_layout.addWidget(stock_group)
        
        # 按钮区域
        button_layout = QHBoxLayout()
        
        # 刷新库存按钮
        refresh_btn = QPushButton("🔄 刷新库存")
        refresh_btn.clicked.connect(self.update_stock_check)
        button_layout.addWidget(refresh_btn)
        
        # 帮助说明按钮
        help_btn = QPushButton("❓ 稀释说明")
        help_btn.clicked.connect(self.show_dilution_help)
        button_layout.addWidget(help_btn)
        
        button_layout.addStretch()
        
        # 确定和取消按钮
        self.ok_btn = QPushButton("✅ 确认使用")
        self.ok_btn.setEnabled(False)  # 初始禁用
        self.ok_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton("❌ 取消")
        cancel_btn.clicked.connect(self.reject)
        
        button_layout.addWidget(self.ok_btn)
        button_layout.addWidget(cancel_btn)
        
        main_layout.addLayout(button_layout)
        self.setLayout(main_layout)
        
        # 初始化库存检查
        self.update_stock_check()
    
    def show_dilution_help(self):
        """显示稀释倍数说明"""
        help_text = """
🧪 稀释倍数与原料消耗计算说明

📌 基本概念：
• 配方用量：配方中记录的稀释后用量
• 稀释倍数：原料被稀释的倍数（如10表示稀释10倍）
• 实际消耗：从库存中真正扣除的原料量

📊 计算公式：
实际消耗量 = 配方用量 ÷ 稀释倍数 × 生产倍数

💡 举例说明：
假设乙基麦芽酚的配方数据为：
• 配方用量：5.0g（稀释后的用量）
• 稀释倍数：10（稀释了10倍）
• 生产倍数：1（按标准用量生产）

那么实际消耗的原料 = 5.0g ÷ 10 × 1 = 0.5g

🎯 为什么这样计算？
调香师通常使用稀释过的原料来调制配方：
1. 强烈香料（如乙基麦芽酚）需要稀释后使用
2. 配方中记录的是稀释后的使用量
3. 但库存扣减需要按原料的实际消耗量计算

📈 表格说明：
• 配方用量：配方中设定的稀释后用量
• 稀释倍数：该原料的稀释倍数（>1表示已稀释）
• 实际消耗：★ 关键数据 - 真正从库存扣除的量
• 剩余库存：扣除实际消耗后的剩余量

这样可以确保库存管理的准确性！
        """
        
        dialog = QDialog(self)
        dialog.setWindowTitle("稀释倍数计算说明")
        dialog.setMinimumSize(500, 600)
        
        layout = QVBoxLayout()
        
        text_edit = QTextEdit()
        text_edit.setPlainText(help_text.strip())
        text_edit.setReadOnly(True)
        text_edit.setStyleSheet("""
            QTextEdit {
                font-family: '微软雅黑', Arial;
                font-size: 12px;
                line-height: 1.5;
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 4px;
                padding: 10px;
            }
        """)
        
        layout.addWidget(text_edit)
        
        close_btn = QPushButton("知道了")
        close_btn.clicked.connect(dialog.accept)
        layout.addWidget(close_btn)
        
        dialog.setLayout(layout)
        dialog.exec()
    
    def update_stock_check(self):
        """更新库存检查"""
        try:
            multiplier = self.multiplier_spin.value()
            
            if multiplier <= 0:
                self.status_label.setText("⚠️ 请输入有效的生产倍数")
                self.status_label.setStyleSheet("""
                    QLabel {
                        background-color: #fff3e0;
                        border: 1px solid #ff9800;
                        border-radius: 4px;
                        padding: 8px;
                        color: #f57c00;
                    }
                """)
                self.ok_btn.setEnabled(False)
                return
            
            # 解析配方内容
            ingredients_data = []
            if not self.formula.content:
                self.status_label.setText("❌ 配方内容为空")
                self.ok_btn.setEnabled(False)
                return
            
            all_sufficient = True
            
            for item in self.formula.content.split(','):
                if not item.strip():
                    continue
                    
                try:
                    parts = item.strip().split(':')
                    # 支持新旧两种格式：旧格式3个部分，新格式5个部分
                    if len(parts) < 3:
                        continue
                        
                    name = parts[0].strip()
                    percentage = float(parts[1].rstrip('%'))
                    
                    # 解析稀释倍数（从配方内容中）
                    dilution_ratio = 1.0  # 默认无稀释
                    if len(parts) > 3:
                        try:
                            dilution_ratio = float(parts[3])
                        except:
                            dilution_ratio = 1.0
                    
                    # 使用绝对用量而不是百分比计算
                    if len(parts) >= 3 and parts[2].endswith('g'):
                        # 新格式：直接使用绝对用量
                        formula_amount = float(parts[2].rstrip('g'))  # 配方中的稀释后用量
                        # 计算实际需要的原料量（考虑稀释倍数）
                        # 如果乙基麦芽酚稀释了10倍，配方用量5g，实际消耗原料 = 5g / 10 = 0.5g
                        actual_ingredient_consumption = (formula_amount / dilution_ratio) * multiplier
                    else:
                        # 旧格式：不支持，因为没有总重量参考
                        print(f"警告：配方项目 {item} 使用旧格式，无法计算")
                        continue
                    
                    # 查询当前库存 - 与库存汇总界面保持完全一致
                    ingredient = self.session.query(Ingredient).filter_by(name=name).first()
                    if ingredient:
                        # 计算当前库存：入库 - 出库
                        in_stock = self.session.query(func.sum(StockRecord.quantity)).filter(
                            StockRecord.ingredient_id == ingredient.id,
                            StockRecord.operation_type == 'in',
                            StockRecord.is_deleted == False
                        ).scalar() or 0.0
                        
                        out_stock = self.session.query(func.sum(StockRecord.quantity)).filter(
                            StockRecord.ingredient_id == ingredient.id,
                            StockRecord.operation_type == 'out',
                            StockRecord.is_deleted == False
                        ).scalar() or 0.0
                        
                        current_stock = in_stock - out_stock
                    else:
                        current_stock = 0.0
                    
                    # 判断库存状态（基于实际消耗量）
                    is_sufficient = current_stock >= actual_ingredient_consumption
                    if not is_sufficient:
                        all_sufficient = False
                    
                    remaining_stock = current_stock - actual_ingredient_consumption
                    
                    ingredients_data.append({
                        'name': name,
                        'percentage': percentage,
                        'formula_amount': formula_amount,
                        'dilution_ratio': dilution_ratio,
                        'actual_consumption': actual_ingredient_consumption,
                        'current': current_stock,
                        'sufficient': is_sufficient,
                        'remaining': remaining_stock
                    })
                    
                except (ValueError, IndexError) as e:
                    print(f"解析配方项目失败: {item}, 错误: {e}")
                    continue
            
            # 更新表格
            self.stock_table.setRowCount(len(ingredients_data))
            
            for row, data in enumerate(ingredients_data):
                # 原料名称
                self.stock_table.setItem(row, 0, QTableWidgetItem(data['name']))
                
                # 配方比例
                self.stock_table.setItem(row, 1, QTableWidgetItem(f"{data['percentage']:.1f}%"))
                
                # 配方用量（稀释后用量）
                formula_item = QTableWidgetItem(f"{data['formula_amount']:.2f}")
                formula_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                self.stock_table.setItem(row, 2, formula_item)
                
                # 稀释倍数
                dilution_item = QTableWidgetItem(f"{data['dilution_ratio']:.1f}")
                dilution_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                if data['dilution_ratio'] > 1:
                    dilution_item.setBackground(QColor(255, 248, 220))  # 浅黄色标识稀释
                self.stock_table.setItem(row, 3, dilution_item)
                
                # 实际消耗（原料消耗量）
                consumption_item = QTableWidgetItem(f"{data['actual_consumption']:.3f}")
                consumption_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                consumption_item.setBackground(QColor(230, 255, 230))  # 浅绿色强调这是关键数据
                self.stock_table.setItem(row, 4, consumption_item)
                
                # 当前库存
                current_item = QTableWidgetItem(f"{data['current']:.2f}")
                current_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                self.stock_table.setItem(row, 5, current_item)
                
                # 库存状态
                status_item = QTableWidgetItem("✅ 充足" if data['sufficient'] else "❌ 不足")
                if data['sufficient']:
                    status_item.setBackground(QColor(200, 255, 200))  # 浅绿色
                else:
                    status_item.setBackground(QColor(255, 200, 200))  # 浅红色
                self.stock_table.setItem(row, 6, status_item)
                
                # 剩余库存
                remaining_item = QTableWidgetItem(f"{data['remaining']:.3f}")
                remaining_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                if data['remaining'] < 0:
                    remaining_item.setBackground(QColor(255, 200, 200))  # 浅红色
                self.stock_table.setItem(row, 7, remaining_item)
            
            # 更新状态标签
            if all_sufficient:
                total_consumption = sum(data['actual_consumption'] for data in ingredients_data)
                self.status_label.setText(f"✅ 库存充足，可以使用配方（生产倍数: {multiplier}x，总原料消耗: {total_consumption:.3f}g）")
                self.status_label.setStyleSheet("""
                    QLabel {
                        background-color: #e8f5e8;
                        border: 1px solid #4caf50;
                        border-radius: 4px;
                        padding: 8px;
                        color: #2e7d32;
                    }
                """)
                self.ok_btn.setEnabled(True)
            else:
                insufficient_count = sum(1 for data in ingredients_data if not data['sufficient'])
                self.status_label.setText(f"❌ 有 {insufficient_count} 种原料库存不足，无法使用配方")
                self.status_label.setStyleSheet("""
                    QLabel {
                        background-color: #ffebee;
                        border: 1px solid #f44336;
                        border-radius: 4px;
                        padding: 8px;
                        color: #c62828;
                    }
                """)
                self.ok_btn.setEnabled(False)
                
        except Exception as e:
            self.status_label.setText(f"❌ 检查库存时出错: {str(e)}")
            self.ok_btn.setEnabled(False)
    
    def accept(self):
        """确认使用配方"""
        # 验证必填字段
        if not self.batch_edit.text().strip():
            QMessageBox.warning(self, "警告", "请输入批次号！")
            return
            
        if not self.operator_edit.text().strip():
            QMessageBox.warning(self, "警告", "请输入操作人！")
            return
        
        try:
            multiplier = self.multiplier_spin.value()
            
            # 计算配方总量（所有原料用量之和）
            total_formula_amount = 0
            for item in self.formula.content.split(','):
                if item.strip():
                    parts = item.strip().split(':')
                    if len(parts) >= 3 and parts[2].endswith('g'):
                        total_formula_amount += float(parts[2].rstrip('g'))
            
            actual_amount = total_formula_amount * multiplier
            
            # 创建配方使用记录（添加新字段）
            usage = FormulaUsage(
                formula_id=self.formula.id,
                formula_name=self.formula.name,
                batch_number=self.batch_edit.text().strip(),
                operator=self.operator_edit.text().strip(),
                total_amount=actual_amount,
                multiplier=multiplier,
                notes=f"用途:{self.purpose_combo.currentText()}, 质量:{self.quality_combo.currentText()}, 时间:{self.production_time_edit.text()}, 备注:{self.notes_edit.text().strip()}"
            )
            self.session.add(usage)
            self.session.flush()  # 获取usage.id
            
            # 扣减库存 - 创建出库记录
            for item in self.formula.content.split(','):
                if not item.strip():
                    continue
                    
                try:
                    parts = item.strip().split(':')
                    # 支持新旧两种格式：旧格式3个部分，新格式5个部分
                    if len(parts) < 3:
                        continue
                        
                    name = parts[0].strip()
                    percentage = float(parts[1].rstrip('%'))
                    
                    # 解析稀释倍数（从配方内容中）
                    dilution_ratio = 1.0  # 默认无稀释
                    if len(parts) > 3:
                        try:
                            dilution_ratio = float(parts[3])
                        except:
                            dilution_ratio = 1.0
                    
                    # 使用绝对用量而不是百分比计算
                    if len(parts) >= 3 and parts[2].endswith('g'):
                        # 新格式：直接使用绝对用量
                        formula_amount = float(parts[2].rstrip('g'))
                        # 计算实际需要的原料量（考虑稀释倍数）
                        actual_ingredient_consumption = (formula_amount / dilution_ratio) * multiplier
                    else:
                        # 旧格式：不支持，因为没有总重量参考
                        print(f"警告：配方项目 {item} 使用旧格式，跳过")
                        continue
                    
                    # 查找原料
                    ingredient = self.session.query(Ingredient).filter_by(name=name).first()
                    if ingredient:
                        # 创建出库记录 - 使用实际消耗量而不是配方用量
                        stock_record = StockRecord(
                            ingredient_id=ingredient.id,
                            ingredient_number=ingredient.number,
                            ingredient_name=ingredient.name,
                            quantity=-actual_ingredient_consumption,  # 负数表示出库，使用实际消耗量
                            supplier="配方使用",
                            batch_number=self.batch_edit.text().strip(),
                            operation_type='out',
                            operator=self.operator_edit.text().strip(),
                            formula_usage_id=usage.id
                        )
                        self.session.add(stock_record)
                        
                except (ValueError, IndexError) as e:
                    print(f"处理配方项目失败: {item}, 错误: {e}")
                    continue
            
            self.session.commit()
            
            # 刷新父窗口的库存界面
            if hasattr(self.parent, 'refresh_stock_summary'):
                self.parent.refresh_stock_summary()
            if hasattr(self.parent, 'refresh_stock_table'):
                self.parent.refresh_stock_table()
            
            QMessageBox.information(
                self, 
                "成功", 
                f"配方使用记录已创建！\n"
                f"批次号: {self.batch_edit.text()}\n"
                f"使用总量: {actual_amount:.2f}g\n"
                f"库存已同步更新"
            )
            super().accept()
            
        except Exception as e:
            self.session.rollback()
            QMessageBox.critical(self, "错误", f"使用配方失败：{str(e)}")


class StockRecordDialog(QDialog):
    """库存记录对话框"""
    def __init__(self, parent, session):
        super().__init__(parent)
        self.setWindowTitle("添加库存记录")
        self.session = session
        self.resize(500, 400)
        
        # 初始化映射字典
        self.ingredient_map = {}
        self.supplier_map = {}
        
        # 主布局
        main_layout = QVBoxLayout()
        
        # 表单区域
        form = QFormLayout()
        
        # 操作类型
        self.type_combo = QComboBox()
        self.type_combo.addItems(["in", "out"])
        form.addRow("操作类型", self.type_combo)
        
        # 数量
        self.quantity_edit = QDoubleSpinBox()
        self.quantity_edit.setDecimals(2)
        self.quantity_edit.setMaximum(999999.99)
        form.addRow("数量(g)", self.quantity_edit)
        
        # 原料选择区域
        ingredient_layout = QVBoxLayout()
        
        # 原料搜索框
        ingredient_search_layout = QHBoxLayout()
        ingredient_search_layout.addWidget(QLabel("🔍 搜索原料:"))
        self.ingredient_search = QLineEdit()
        self.ingredient_search.setPlaceholderText("输入原料名称、编号或CAS号进行筛选...")
        self.ingredient_search.textChanged.connect(self.filter_ingredients)
        self.ingredient_search.setStyleSheet("""
            QLineEdit {
                padding: 6px;
                border: 2px solid #ddd;
                border-radius: 4px;
                font-size: 14px;
            }
            QLineEdit:focus {
                border-color: #4CAF50;
            }
        """)
        ingredient_search_layout.addWidget(self.ingredient_search)
        
        # 清除搜索按钮
        clear_search_btn = QPushButton("×")
        clear_search_btn.setFixedSize(24, 24)
        clear_search_btn.clicked.connect(self.clear_ingredient_search)
        clear_search_btn.setStyleSheet("""
            QPushButton {
                background-color: #ff6b6b;
                color: white;
                border: none;
                border-radius: 12px;
                font-weight: bold;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #ff5252;
                transform: scale(1.1);
            }
            QPushButton:pressed {
                background-color: #e53e3e;
            }
        """)
        clear_search_btn.setToolTip("清除搜索")
        ingredient_search_layout.addWidget(clear_search_btn)
        
        ingredient_layout.addLayout(ingredient_search_layout)
        
        # 原料下拉框
        self.ingredient_combo = QComboBox()
        self.ingredient_combo.setEditable(True)
        self.ingredient_combo.currentTextChanged.connect(self.update_ingredient_info)
        self.ingredient_combo.setStyleSheet("""
            QComboBox {
                padding: 6px;
                border: 2px solid #ddd;
                border-radius: 4px;
                font-size: 14px;
                min-height: 20px;
            }
            QComboBox:focus {
                border-color: #4CAF50;
            }
            QComboBox::drop-down {
                border: none;
                width: 20px;
            }
            QComboBox::down-arrow {
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 5px solid #666;
            }
        """)
        self.all_ingredients = self.session.query(Ingredient).all()
        self.populate_ingredient_combo()
        ingredient_layout.addWidget(self.ingredient_combo)
        
        # 原料信息显示
        self.ingredient_info_label = QLabel("💡 选择原料后将显示详细信息")
        self.ingredient_info_label.setStyleSheet("""
            QLabel {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 4px;
                padding: 8px;
                color: #6c757d;
                font-size: 12px;
            }
        """)
        ingredient_layout.addWidget(self.ingredient_info_label)
        
        # 将原料选择区域添加到表单
        ingredient_widget = QWidget()
        ingredient_widget.setLayout(ingredient_layout)
        form.addRow("原料", ingredient_widget)
        
        # 供应商下拉选择
        self.supplier_combo = QComboBox()
        self.supplier_combo.setEditable(True)
        self.populate_supplier_combo()
        form.addRow("供应商", self.supplier_combo)
        
        # 批次号
        self.batch_edit = QLineEdit()
        form.addRow("批次号", self.batch_edit)
        
        # 操作员
        self.operator_edit = QLineEdit()
        form.addRow("操作员", self.operator_edit)
        
        # 过期日期
        self.expiration_edit = QDateEdit()
        self.expiration_edit.setDate(QDate.currentDate().addDays(365))
        self.expiration_edit.setCalendarPopup(True)
        form.addRow("过期日期", self.expiration_edit)
        
        # 添加表单到主布局
        form_widget = QWidget()
        form_widget.setLayout(form)
        main_layout.addWidget(form_widget)
        
        # 按钮区域
        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        main_layout.addWidget(btns)
        
        self.setLayout(main_layout)

    def populate_supplier_combo(self, filter_text=""):
        """填充供应商下拉框"""
        self.supplier_combo.clear()
        self.supplier_map.clear()
        
        # 添加默认选项
        self.supplier_combo.addItem("-- 请选择供应商 --")
        self.supplier_map["-- 请选择供应商 --"] = None
        
        # 查询所有供应商
        from models import Supplier
        suppliers = self.session.query(Supplier).order_by(Supplier.name).all()
        
        for supplier in suppliers:
            # 创建显示文本：供应商名称 - 联系人 - 电话
            display_text = supplier.name
            if supplier.contact_person:
                display_text += f" - {supplier.contact_person}"
            if supplier.phone:
                display_text += f" - {supplier.phone}"
            
            # 如果有过滤文本，检查是否匹配
            if not filter_text or any(filter_text.lower() in str(field).lower() 
                                    for field in [supplier.name, supplier.contact_person, supplier.phone] 
                                    if field):
                self.supplier_combo.addItem(display_text)
                self.supplier_map[display_text] = supplier

    def filter_ingredients(self):
        """根据搜索文本过滤原料"""
        filter_text = self.ingredient_search.text().strip()
        print(f"搜索文本: '{filter_text}'")  # 调试信息
        
        # 保存当前选择
        current_selection = self.ingredient_combo.currentText()
        
        # 重新填充下拉框
        self.populate_ingredient_combo(filter_text)
        
        # 如果有搜索结果且没有当前选择，自动展开下拉框
        if filter_text and self.ingredient_combo.count() > 1:  # 大于1是因为有默认选项
            self.ingredient_combo.showPopup()  # 显示下拉框

    def clear_ingredient_search(self):
        """清除原料搜索"""
        self.ingredient_search.clear()
        self.populate_ingredient_combo()  # 重新加载所有原料
        self.ingredient_combo.setCurrentIndex(0)  # 选择默认选项

    def update_ingredient_info(self):
        """更新原料信息显示"""
        # 检查 ingredient_info_label 是否存在，如果不存在就跳过
        if not hasattr(self, 'ingredient_info_label'):
            return
            
        ingredient_display = self.ingredient_combo.currentText()
        selected_ingredient = self.ingredient_map.get(ingredient_display)
        
        if selected_ingredient:
            # 获取当前库存
            from sqlalchemy import func
            current_stock = self.session.query(func.sum(StockRecord.quantity)).filter_by(ingredient_id=selected_ingredient.id).scalar() or 0
            
            info_parts = [f"📊 原料信息：{selected_ingredient.name}"]
            
            if selected_ingredient.number:
                info_parts.append(f"🔢 编号：{selected_ingredient.number}")
            
            if selected_ingredient.cas_number:
                info_parts.append(f"🧪 CAS号：{selected_ingredient.cas_number}")
                
            if selected_ingredient.english_name:
                info_parts.append(f"🌐 英文名：{selected_ingredient.english_name}")
                
            info_parts.append(f"📦 当前库存：{current_stock:.2f}g")
            
            self.ingredient_info_label.setText("\n".join(info_parts))
            self.ingredient_info_label.setStyleSheet("""
                QLabel {
                    background-color: #e8f5e8;
                    border: 1px solid #28a745;
                    border-radius: 4px;
                    padding: 8px;
                    color: #155724;
                    font-size: 12px;
                }
            """)
        else:
            self.ingredient_info_label.setText("💡 选择原料后将显示详细信息")
            self.ingredient_info_label.setStyleSheet("""
                QLabel {
                    background-color: #f8f9fa;
                    border: 1px solid #dee2e6;
                    border-radius: 4px;
                    padding: 8px;
                    color: #6c757d;
                    font-size: 12px;
                }
            """)

    def populate_ingredient_combo(self, filter_text=""):
        """填充原料下拉框"""
        self.ingredient_combo.clear()
        self.ingredient_map.clear()
        
        # 添加默认选项
        self.ingredient_combo.addItem("-- 请选择原料 --")
        self.ingredient_map["-- 请选择原料 --"] = None
        
        matched_count = 0  # 调试信息：匹配数量
        
        for ingredient in self.all_ingredients:
            # 创建显示文本：编号 - 名称 - CAS
            display_text = f"{ingredient.number} - {ingredient.name}"
            if ingredient.cas_number:
                display_text += f" - {ingredient.cas_number}"
            
            # 如果有过滤文本，检查是否匹配
            if not filter_text:
                # 没有过滤文本，显示所有
                self.ingredient_combo.addItem(display_text)
                self.ingredient_map[display_text] = ingredient
                matched_count += 1
            else:
                # 有过滤文本，检查匹配
                fields_to_check = [ingredient.number, ingredient.name, ingredient.cas_number]
                match_found = False
                
                for field in fields_to_check:
                    if field and filter_text.lower() in str(field).lower():
                        match_found = True
                        print(f"匹配找到: '{field}' 包含 '{filter_text}'")  # 调试信息
                        break
                
                if match_found:
                    self.ingredient_combo.addItem(display_text)
                    self.ingredient_map[display_text] = ingredient
                    matched_count += 1
        
        print(f"总原料数: {len(self.all_ingredients)}, 匹配数量: {matched_count}")  # 调试信息

    def get_data(self):
        """获取对话框数据"""
        ing_display = self.ingredient_combo.currentText()
        ing = self.ingredient_map.get(ing_display)
        
        if not ing:
            return None
            
        # 获取供应商信息
        supplier_display = self.supplier_combo.currentText()
        selected_supplier = self.supplier_map.get(supplier_display)
        supplier_name = selected_supplier.name if selected_supplier else supplier_display
        
        expiration_date = self.expiration_edit.date().toPyDate()
        quantity = self.quantity_edit.value()
        if self.type_combo.currentText() == 'out':
            quantity = -abs(quantity)
        else:
            quantity = abs(quantity)
            
        return {
            'ingredient_id': ing.id,
            'ingredient_number': ing.number,
            'ingredient_name': ing.name,
            'quantity': quantity,
            'supplier': supplier_name,
            'batch_number': self.batch_edit.text().strip(),
            'operator': self.operator_edit.text().strip(),
            'operation_type': self.type_combo.currentText(),
            'expiration_date': expiration_date,
            'created_at': datetime.now(),
        }


if __name__ == '__main__':
    app = QApplication(sys.argv)
    
    # Set application style
    app.setStyle('Fusion')
    
    window = MainWindow()
    window.show()
    sys.exit(app.exec())