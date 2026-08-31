#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
数据模型定义
从 alembic 迁移与字节码反编译重建
"""
from datetime import datetime

from sqlalchemy import (create_engine, Column, Integer, String, Float,
                        ForeignKey, DateTime, Table, Text, Boolean)
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import relationship

Base = declarative_base()

# 原料-配方 多对多关联表
ingredient_formula = Table(
    'ingredient_formula', Base.metadata,
    Column('ingredient_id', Integer, ForeignKey('ingredients.id')),
    Column('formula_id', Integer, ForeignKey('formulas.id')),
    Column('percentage', Float, nullable=False)
)


class Ingredient(Base):
    """原料表"""
    __tablename__ = 'ingredients'

    id = Column(Integer, primary_key=True)
    number = Column(String(50), unique=True, nullable=False)
    cas_number = Column(String(50))
    name = Column(String(100), nullable=False)
    english_name = Column(String(100))
    molecular_formula = Column(String(100))
    chemical_structure = Column(String(500))
    molecular_weight = Column(String(50))
    boiling_point = Column(String(50))
    solubility = Column(String(100))
    max_limit_gb = Column(String(100))
    natural_occurrence = Column(String(200))
    perfume_usage = Column(String(200))
    aroma_character = Column(String(200))
    aroma_change = Column(String(200))
    sniff_aroma = Column(String(200))
    sensory_evaluation = Column(String(200))
    aroma_composition = Column(String(200))
    price = Column(Float, default=0.0)
    min_stock_threshold = Column(Float, default=0.0)
    created_at = Column(DateTime, default=datetime.now)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)

    formulas = relationship('Formula', secondary=ingredient_formula,
                            back_populates='ingredients')
    stock_records = relationship('StockRecord', back_populates='ingredient')


class Formula(Base):
    """配方表"""
    __tablename__ = 'formulas'

    id = Column(Integer, primary_key=True)
    number = Column(String(50))
    name = Column(String(100), nullable=False)
    creator = Column(String(50))
    version = Column(String(20))
    description = Column(String(500))
    content = Column(String(1000))
    evaluation = Column(String(500))
    total_cost = Column(Float)
    created_at = Column(DateTime, default=datetime.now)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)

    ingredients = relationship('Ingredient', secondary=ingredient_formula,
                               back_populates='formulas')
    usages = relationship('FormulaUsage', back_populates='formula')


class StockRecord(Base):
    """库存记录表"""
    __tablename__ = 'stock_records'

    id = Column(Integer, primary_key=True)
    ingredient_id = Column(Integer, ForeignKey('ingredients.id'))
    ingredient_number = Column(String(50))
    ingredient_name = Column(String(100))
    quantity = Column(Float)
    supplier = Column(String(100))
    batch_number = Column(String(50))
    operation_type = Column(String(20))
    operation_time = Column(DateTime, default=datetime.now)
    expiration_date = Column(DateTime)
    operator = Column(String(50))
    formula_usage_id = Column(Integer, ForeignKey('formula_usage.id'))
    created_at = Column(DateTime, default=datetime.now)
    is_deleted = Column(Boolean, default=False)

    ingredient = relationship('Ingredient', back_populates='stock_records')
    formula_usage = relationship('FormulaUsage', back_populates='stock_records')


class Supplier(Base):
    """供应商表"""
    __tablename__ = 'suppliers'

    id = Column(Integer, primary_key=True)
    name = Column(String(100), nullable=False)
    contact_person = Column(String(50))
    phone = Column(String(20))
    email = Column(String(100))
    address = Column(String(200))
    created_at = Column(DateTime, default=datetime.now)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)


class GCMSAnalysis(Base):
    """GC-MS分析表"""
    __tablename__ = 'gcms_analyses'

    id = Column(Integer, primary_key=True)
    number = Column(String(50), unique=True)
    name = Column(String(100))
    supplier = Column(String(100))
    instrument_params = Column(String(500))
    perfume_idea = Column(String(1000))
    analysis_time = Column(DateTime, default=datetime.now)
    spectrum_image = Column(String(500))

    compounds = relationship('GCMSCompound', back_populates='analysis')


class GCMSCompound(Base):
    """GC-MS化合物表"""
    __tablename__ = 'gcms_compounds'

    id = Column(Integer, primary_key=True)
    analysis_id = Column(Integer, ForeignKey('gcms_analyses.id'))
    cas = Column(String(50))
    rt = Column(String(50))
    name_en = Column(String(100))
    name_cn = Column(String(100))
    match_factor = Column(String(50))
    formula = Column(String(100))
    relative_content = Column(Float)
    content = Column(Float)
    unit = Column(String(20))

    analysis = relationship('GCMSAnalysis', back_populates='compounds')


class FormulaUsage(Base):
    """配方使用记录表"""
    __tablename__ = 'formula_usage'

    id = Column(Integer, primary_key=True)
    formula_id = Column(Integer, ForeignKey('formulas.id'), nullable=False)
    formula_name = Column(String(100), nullable=False)
    batch_number = Column(String(100), nullable=False)
    operator = Column(String(50), nullable=False)
    total_amount = Column(Float, nullable=False)
    multiplier = Column(Float, default=1.0)
    usage_time = Column(DateTime, default=datetime.now)
    notes = Column(Text)
    created_at = Column(DateTime, default=datetime.now)

    formula = relationship('Formula', back_populates='usages')
    stock_records = relationship('StockRecord', back_populates='formula_usage')


class AppSetting(Base):
    """应用设置表（键值对，用于标记一次性操作等）"""
    __tablename__ = 'app_settings'

    key = Column(String(50), primary_key=True)
    value = Column(Text)
