from sqlalchemy import create_engine, Column, DateTime
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
from datetime import datetime
import sqlite3

# 创建数据库连接
engine = create_engine('sqlite:///fragrance_management.db')
Base = declarative_base()
Session = sessionmaker(bind=engine)
session = Session()

def upgrade():
    try:
        # 直接使用 sqlite3 连接
        conn = sqlite3.connect('fragrance_management.db')
        cursor = conn.cursor()
        
        # 检查列是否已存在
        cursor.execute("PRAGMA table_info(gcms_analyses)")
        columns = [column[1] for column in cursor.fetchall()]
        
        if 'analysis_time' not in columns:
            print("添加 analysis_time 列...")
            cursor.execute('ALTER TABLE gcms_analyses ADD COLUMN analysis_time DATETIME')
            # 为现有记录设置默认值
            cursor.execute('UPDATE gcms_analyses SET analysis_time = ? WHERE analysis_time IS NULL', 
                        (datetime.now(),))
            conn.commit()
            print("analysis_time 列添加成功！")
        else:
            print("analysis_time 列已存在，无需添加。")
            
    except Exception as e:
        print(f"迁移过程中发生错误：{str(e)}")
        raise
    finally:
        conn.close()

def downgrade():
    try:
        # 直接使用 sqlite3 连接
        conn = sqlite3.connect('fragrance_management.db')
        cursor = conn.cursor()
        
        # 检查列是否存在
        cursor.execute("PRAGMA table_info(gcms_analyses)")
        columns = [column[1] for column in cursor.fetchall()]
        
        if 'analysis_time' in columns:
            print("删除 analysis_time 列...")
            # SQLite 不支持直接删除列，需要创建新表并复制数据
            cursor.execute('''
                CREATE TABLE gcms_analyses_new (
                    id INTEGER PRIMARY KEY,
                    number VARCHAR(50) UNIQUE,
                    name VARCHAR(100),
                    supplier VARCHAR(100),
                    instrument_params VARCHAR(500),
                    perfume_idea VARCHAR(1000)
                )
            ''')
            cursor.execute('''
                INSERT INTO gcms_analyses_new (id, number, name, supplier, instrument_params, perfume_idea)
                SELECT id, number, name, supplier, instrument_params, perfume_idea
                FROM gcms_analyses
            ''')
            cursor.execute('DROP TABLE gcms_analyses')
            cursor.execute('ALTER TABLE gcms_analyses_new RENAME TO gcms_analyses')
            conn.commit()
            print("analysis_time 列删除成功！")
        else:
            print("analysis_time 列不存在，无需删除。")
            
    except Exception as e:
        print(f"回滚过程中发生错误：{str(e)}")
        raise
    finally:
        conn.close()

if __name__ == '__main__':
    upgrade() 