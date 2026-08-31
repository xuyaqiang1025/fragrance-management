#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AI 调香专家模块（香韵大师）
将 flavor-fragrance-expert skill 内置进香料管理系统：
- 专家人设与方法论作为系统提示词（GC-MS解析/香韵重组/双场景合规/雾化适配/单体库）
- 接入 OpenAI 兼容大模型 API（可配置：智谱GLM/DeepSeek/Kimi/OpenAI/自定义）
- 记忆学习系统：候选库（默认）/生效库（用户确认升格）闸门，防止污染
- 上下文注入：可附上系统内原料库/配方库/GC-MS分析数据
- 报告交付：对话与配方报告导出为 MD + HTML（香韵大师视觉风格）
"""
import json
import os
import re
import sys
import html as html_mod
from datetime import datetime

from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel,
                             QPushButton, QTextEdit, QComboBox, QGroupBox,
                             QMessageBox, QDialog, QFormLayout, QLineEdit,
                             QDoubleSpinBox, QCheckBox, QTabWidget,
                             QPlainTextEdit, QListWidget, QListWidgetItem,
                             QInputDialog, QFileDialog, QSpinBox, QSplitter)
from PyQt6.QtWidgets import QDialogButtonBox
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt6.QtGui import QFont


# ---------------------------------------------------------------- 路径与配置

def get_app_dir():
    """应用目录：打包后为 exe 所在目录，开发时为 src 目录"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = get_app_dir()
CONFIG_PATH = os.path.join(APP_DIR, 'ai_expert_config.json')
MEMORY_DIR = os.path.join(APP_DIR, '.flavor-memory')
DELIVERY_DIR = os.path.join(APP_DIR, 'deliveries')

PROVIDER_PRESETS = {
    '智谱GLM': {
        'base_url': 'https://open.bigmodel.cn/api/paas/v4/',
        'model': 'glm-4-flash'},
    'DeepSeek': {
        'base_url': 'https://api.deepseek.com/v1',
        'model': 'deepseek-chat'},
    'Kimi(月之暗面)': {
        'base_url': 'https://api.moonshot.cn/v1',
        'model': 'moonshot-v1-32k'},
    'OpenAI': {
        'base_url': 'https://api.openai.com/v1',
        'model': 'gpt-4o-mini'},
    '自定义/OpenAI兼容': {
        'base_url': 'http://localhost:11434/v1',
        'model': ''},
}


def load_config():
    cfg = dict(base_url='', api_key='', model='', temperature=0.7)
    try:
        if os.path.exists(CONFIG_PATH):
            with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                cfg.update(json.load(f))
    except Exception:
        pass
    return cfg


def save_config(cfg):
    try:
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ---------------------------------------------------------------- 系统提示词

SYSTEM_PROMPT = """你是「香韵大师」——一位深耕风味化学与香精配方研发8年的资深调香师，聚焦**电子雾化香精**与**食用香精**两大核心赛道，覆盖竞品逆向解析、实验室配方开发到量产落地优化的全研发链路。你精通GC-MS全谱解析、香韵结构重组、双场景合规管控、雾化工程化适配与合成单体全库认知。

## 核心能力
1. **GC-MS全谱解析**：溶剂峰扣除、杂峰过滤、共流出峰拆分、组分定性定量、OT值感官校正（突破峰面积局限还原真实感官贡献）、关键特征组分缺口补全（热敏物质、痕量单体）
2. **香韵结构重组**：基于「五层骨架（头香/体香/基香/特征指纹层/修饰层）+香型专属分路」方法论；分路推导引擎：特征路识别→基底路构建→修饰路补充→平衡校验；分路贡献量化（主导≥20%、支撑5%-20%、修饰<5%）；简单香型3-5路、中等5-7路、高复杂度7-10路
3. **双场景合规管控**：GB 2760（食用香精）、GB 41700+欧盟TPD+美国PMTA（电子雾化）；高风险物质（α-二酮类、呋喃衍生物、香豆素类）管控边界与等效替代方案库
4. **雾化工程化适配**：PG/VG体系溶解度校验（防析晶分层）、200-300°C热稳定性评估、积碳风险管控、萜烯类刺激性调控、击喉感/口腔饱满度/嗅抽一致性优化
5. **合成单体全库**：吡嗪类（坚果/焙烤）、噻唑类（肉香）、酯类（果香）、内酯类（奶香）、萜烯类（柑橘）、醛酮类（青香）、含硫含氮杂环（肉香/咖啡）等上千种单体的感官特征、嗅觉阈值、理化参数与毒理属性

## 工作流程
- **需求确认**：明确需求类型（逆向仿香/新香型开发/配方合规改造/雾化问题诊断/单体咨询）、目标场景（电子雾化/食用）、目标香型、输出版本（极致还原版/合规落地版/双版本）
- **GC-MS解析**（如有数据）：预处理→定性→共流出拆分→定量→OT值校正→特征指纹识别→缺口分析
- **配方设计**：五层骨架+香韵分路拆分双维度分析；OT值导向确定用量；补全GC未检出关键组分；调控酸甜比、烘焙度、鲜爽感等风格维度
- **合规审查**：受限单体识别→等效替代（最小感官损失）→双版本对比说明
- **雾化适配**（电子雾化场景）：溶解度→热稳定性→积碳→体感→量产风险提示

## 输出规范
- 使用 Markdown 输出；配方报告采用九大章节结构：推荐使用参数 / 一、极致还原版配方（组分表）/ 二、香韵结构分析（五层骨架）/ 三、香韵分路拆分分析 / 四、风格参数 / 五、合规审查报告 / 六、合规落地版配方 / 七、雾化工程化适配报告（电子雾化场景）/ 八、操作安全提示 / 九、配方总结
- 组分表列：序号｜组分名称｜CAS号｜化学类别｜用量(‰)｜功能层｜感官描述｜OT值(ppb)｜合规状态（✓/⚠️/❌）
- 单位规范：配方用量统一千分比（‰），GC-MS数据用百分比（%）或ppm
- **双版本原则**：完整配方默认同时输出「极致还原版」与「合规落地版」
- **数据真实性**：不臆造未检出的组分；推测组分标注"推测"；不确定的CAS号/OT值标注"待核实"
- 高致敏、高毒性单体（某些醛类、含硫单体）须标注安全操作提示
- 配方表用 Markdown 表格呈现，关键数据加粗

## 运行环境
你运行于「香料管理系统」桌面软件中。用户消息可能附带系统内真实数据（原料库/配方库/GC-MS分析记录），请优先结合这些数据回答；引用具体原料时使用其编号与名称；用户选择的目标场景和香型信息以用户消息为准。

## 记忆系统（候选/生效闸门）
若上下文提供了「用户偏好档案」「生效知识 [K0xx]」，请优先应用并在正文中以 [K0xx] 标注引用；候选知识不可直接当作既定事实使用。
每次实质性回答（配方设计、分析报告、经验总结）的最后，输出一节：

### 候选知识（供确认）
- [可复用规律的一句话描述] ｜ tags: 相关标签 ｜ scene: 电子雾化/食用/通用
（提炼0-3条本次回答中可复用的调香规律，供用户确认后升格为生效知识；若本次无可沉淀内容，输出"无"。）"""


# ---------------------------------------------------------------- 记忆库

class MemoryStore:
    """候选/生效双库记忆系统（闸门：默认只进候选，用户确认才升格）"""

    FILES = {
        'profile': ('PROFILE.md', '# 用户偏好档案\n\n（用户确认过的偏好将记录在此）\n'),
        'knowledge': ('KNOWLEDGE.md',
                      '# 生效知识库\n\n<!-- 仅存放用户确认升格的知识条目 [K0xx] -->\n\n'),
        'candidates': ('KNOWLEDGE-CANDIDATES.md',
                       '# 候选知识库\n\n<!-- 默认入库区，经确认后升格为生效知识 -->\n\n'),
        'feedback': ('FEEDBACK-LOG.md', '# 反馈日志\n\n'),
        'index': ('INDEX.md',
                  '# 交付索引\n\n| 时间 | 类型 | 香型 | 模型 | 路径 |\n|---|---|---|---|---|\n'),
    }

    def __init__(self, base_dir=MEMORY_DIR):
        self.base_dir = base_dir
        self.ensure_structure()

    def _path(self, key):
        return os.path.join(self.base_dir, self.FILES[key][0])

    def ensure_structure(self):
        try:
            os.makedirs(self.base_dir, exist_ok=True)
            os.makedirs(os.path.join(self.base_dir, 'formulas'), exist_ok=True)
            os.makedirs(DELIVERY_DIR, exist_ok=True)
            for key, (_, default_content) in self.FILES.items():
                p = self._path(key)
                if not os.path.exists(p):
                    with open(p, 'w', encoding='utf-8') as f:
                        f.write(default_content)
        except Exception:
            pass

    def _read(self, key):
        try:
            with open(self._path(key), 'r', encoding='utf-8') as f:
                return f.read()
        except Exception:
            return ''

    def _write(self, key, content):
        try:
            with open(self._path(key), 'w', encoding='utf-8') as f:
                f.write(content)
        except Exception:
            pass

    # -------- 上下文装载（Phase 0：只装生效库，候选不参与指导设计）--------
    def load_context(self):
        parts = []
        profile = self._read('profile').strip()
        if profile and '（用户确认过的偏好' not in profile:
            parts.append('### 用户偏好档案\n' + profile)
        knowledge = self._read('knowledge').strip()
        lines = [l for l in knowledge.splitlines()
                 if l.strip().startswith(('- ', '[K'))]
        if lines:
            parts.append('### 用户生效知识（可引用）\n' + '\n'.join(lines[:40]))
        index = self._read('index').strip().splitlines()
        idx_rows = [l for l in index if l.startswith('|') and '---' not in l]
        if len(idx_rows) > 1:
            parts.append('### 近期配方交付索引（最近5条）\n'
                         + '\n'.join(idx_rows[-5:]))
        return ('\n\n## 用户记忆库\n' + '\n\n'.join(parts)) if parts else ''

    # -------- 候选沉淀（Phase 6）--------
    def _next_id(self, content, prefix):
        ids = re.findall(r'\[(%s\d{3})\]' % prefix, content)
        if not ids:
            return f'{prefix}001'
        last = max(int(i[len(prefix):]) for i in ids)
        return f'{prefix}{last + 1:03d}'

    def add_candidates(self, entries):
        """entries: [str] 提炼出的候选知识"""
        if not entries:
            return []
        content = self._read('candidates')
        added = []
        for text in entries:
            text = text.strip()
            if not text or text == '无':
                continue
            kid = self._next_id(content, 'KC')
            ts = datetime.now().strftime('%Y-%m-%d %H:%M')
            line = (f'- [{kid}] {text} ｜ status: candidate ｜ '
                    f'evidence: model-only ｜ date: {ts}\n')
            content += line
            added.append((kid, text))
        if added:
            self._write('candidates', content)
        return added

    def get_candidates(self):
        """解析候选条目 -> [(id, text, raw_line)]"""
        result = []
        for line in self._read('candidates').splitlines():
            m = re.match(r'-\s*\[(KC\d{3})\]\s*(.+)$', line.strip())
            if m:
                text = re.split(r'\s+｜\s+status:', m.group(2))[0]
                result.append((m.group(1), text.strip(), line))
        return result

    def promote(self, kc_id):
        """候选升格为生效知识"""
        candidates = self._read('candidates')
        target = None
        kept = []
        for line in candidates.splitlines():
            if f'[{kc_id}]' in line and line.strip().startswith('- '):
                target = line
            else:
                kept.append(line)
        if not target:
            return False
        self._write('candidates', '\n'.join(kept) + '\n')
        knowledge = self._read('knowledge')
        new_id = self._next_id(knowledge, 'K')
        text = re.sub(r'^-\s*\[KC\d{3}\]\s*', '', target.strip())
        text = re.sub(r'\s+｜\s+status:\s*candidate.*$', '', text)
        ts = datetime.now().strftime('%Y-%m-%d')
        knowledge += (f'- [{new_id}] {text} ｜ status: active ｜ '
                      f'evidence: user-confirmed ｜ promoted: {ts}\n')
        self._write('knowledge', knowledge)
        self.append_feedback(f'{kc_id} 升格为 {new_id}（用户确认）', verified=True)
        return True

    def delete_candidate(self, kc_id):
        candidates = self._read('candidates')
        kept = [l for l in candidates.splitlines()
                if not (f'[{kc_id}]' in l and l.strip().startswith('- '))]
        self._write('candidates', '\n'.join(kept) + '\n')

    def append_profile(self, text):
        content = self._read('profile')
        if text.strip() and text.strip() not in content:
            lines = content.rstrip('\n').splitlines()
            lines.append(f'- {text.strip()}（{datetime.now():%Y-%m-%d}）')
            # 最多保留最近20条
            body = [l for l in lines if l.startswith('- ')]
            if len(body) > 20:
                drop = set(id(x) for x in body[:len(body) - 20])
                lines = [l for l in lines if id(l) not in drop]
            self._write('profile', '\n'.join(lines) + '\n')

    def append_feedback(self, text, verified=False):
        content = self._read('feedback')
        ts = datetime.now().strftime('%Y-%m-%d %H:%M')
        content += f'- [{ts}] {text} ｜ verified: {"yes" if verified else "no"}\n'
        self._write('feedback', content)

    def append_index(self, dtype, aroma, model_slug, rel_path):
        content = self._read('index')
        ts = datetime.now().strftime('%Y-%m-%d %H:%M')
        content += f'| {ts} | {dtype} | {aroma} | {model_slug} | {rel_path} |\n'
        self._write('index', content)


# ---------------------------------------------------------------- 解析候选知识

def parse_candidates(answer):
    """从回答中解析「候选知识（供确认）」小节"""
    entries = []
    idx = answer.rfind('候选知识')
    if idx < 0:
        return entries
    section = answer[idx:]
    # 截断到下一个标题
    m = re.search(r'\n#{1,3}\s', section[1:])
    if m:
        section = section[:m.start() + 1]
    for line in section.splitlines():
        line = line.strip()
        if line.startswith(('- ', '· ', '• ')):
            item = line.lstrip('-·• ').strip()
            item = re.sub(r'^\[?KC?\d{3}\]?\s*', '', item)
            if item and item != '无':
                entries.append(item)
    return entries[:3]


# ---------------------------------------------------------------- LLM 线程

class ExpertWorker(QThread):
    """流式对话工作线程（OpenAI 兼容接口）"""
    chunk = pyqtSignal(str)
    finished_ok = pyqtSignal(str)
    failed = pyqtSignal(str)

    def __init__(self, base_url, api_key, model, temperature, messages,
                 parent=None):
        super().__init__(parent)
        self.base_url = base_url.rstrip('/') + '/'
        self.api_key = api_key
        self.model = model
        self.temperature = temperature
        self.messages = messages
        self.cancelled = False
        self._raw = None

    def run(self):
        try:
            from openai import OpenAI
            client = OpenAI(base_url=self.base_url, api_key=self.api_key,
                            timeout=180, max_retries=1)
            collected = []
            stream = client.chat.completions.create(
                model=self.model,
                messages=self.messages,
                temperature=self.temperature,
                stream=True,
            )
            self._raw = stream
            for event in stream:
                if self.cancelled:
                    break
                if event.choices:
                    delta = event.choices[0].delta
                    piece = getattr(delta, 'content', None)
                    if piece:
                        collected.append(piece)
                        self.chunk.emit(piece)
            text = ''.join(collected)
            if self.cancelled and text:
                text += '\n\n*(已停止生成)*'
            if text:
                self.finished_ok.emit(text)
            else:
                self.failed.emit('模型未返回内容，请检查 API 配置或稍后重试。')
        except Exception as e:
            if not self.cancelled:
                self.failed.emit(f'{type(e).__name__}: {e}')

    def stop(self):
        self.cancelled = True
        try:
            if self._raw is not None:
                self._raw.close()
        except Exception:
            pass


class ConnectionTestWorker(QThread):
    """API 连接测试线程"""
    ok = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, base_url, api_key, model, parent=None):
        super().__init__(parent)
        self.base_url = base_url.rstrip('/') + '/'
        self.api_key = api_key
        self.model = model

    def run(self):
        try:
            from openai import OpenAI
            client = OpenAI(base_url=self.base_url, api_key=self.api_key,
                            timeout=15, max_retries=0)
            resp = client.chat.completions.create(
                model=self.model,
                messages=[{'role': 'user', 'content': '请回复：连接成功'}],
                max_tokens=20)
            text = resp.choices[0].message.content or ''
            self.ok.emit(text.strip() or '(空回复)')
        except Exception as e:
            self.error.emit(f'{type(e).__name__}: {e}')


# ---------------------------------------------------------------- Markdown→HTML

def md_to_html(md_text):
    """轻量 Markdown 转 HTML（标题/表格/列表/粗斜体/代码/引用/分隔线）"""
    def inline(s):
        s = html_mod.escape(s, quote=False)
        s = re.sub(r'`([^`]+)`', r'<code>\1</code>', s)
        s = re.sub(r'\*\*([^*]+)\*\*', r'<strong>\1</strong>', s)
        s = re.sub(r'(?<!\*)\*([^*\n]+)\*(?!\*)', r'<em>\1</em>', s)
        return s

    lines = md_text.splitlines()
    out, i = [], 0
    list_stack = []

    def close_lists():
        while list_stack:
            out.append('</%s>' % list_stack.pop())

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if not stripped:
            close_lists()
            i += 1
            continue
        if stripped.startswith('```'):
            close_lists()
            i += 1
            code = []
            while i < len(lines) and not lines[i].strip().startswith('```'):
                code.append(lines[i])
                i += 1
            i += 1
            out.append('<pre><code>' + html_mod.escape('\n'.join(code))
                       + '</code></pre>')
            continue
        m = re.match(r'^(#{1,6})\s+(.*)$', stripped)
        if m:
            close_lists()
            level = len(m.group(1))
            out.append(f'<h{level}>{inline(m.group(2))}</h{level}>')
            i += 1
            continue
        if stripped in ('---', '***', '___'):
            close_lists()
            out.append('<hr/>')
            i += 1
            continue
        if stripped.startswith('|') and i + 1 < len(lines) \
                and re.match(r'^\|[\s:|-]+\|?$', lines[i + 1].strip()):
            close_lists()
            header = [c.strip() for c in stripped.strip('|').split('|')]
            i += 2
            rows = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                rows.append([c.strip() for c in
                             lines[i].strip().strip('|').split('|')])
                i += 1
            out.append('<table><thead><tr>' + ''.join(
                f'<th>{inline(h)}</th>' for h in header) +
                '</tr></thead><tbody>')
            for r in rows:
                out.append('<tr>' + ''.join(
                    f'<td>{inline(c)}</td>' for c in r) + '</tr>')
            out.append('</tbody></table>')
            continue
        m = re.match(r'^[-*]\s+(.*)$', stripped)
        if m:
            if not list_stack or list_stack[-1] != 'ul':
                close_lists()
                out.append('<ul>')
                list_stack.append('ul')
            out.append(f'<li>{inline(m.group(1))}</li>')
            i += 1
            continue
        m = re.match(r'^\d+\.\s+(.*)$', stripped)
        if m:
            if not list_stack or list_stack[-1] != 'ol':
                close_lists()
                out.append('<ol>')
                list_stack.append('ol')
            out.append(f'<li>{inline(m.group(1))}</li>')
            i += 1
            continue
        if stripped.startswith('> '):
            close_lists()
            out.append(f'<blockquote>{inline(stripped[2:])}</blockquote>')
            i += 1
            continue
        # 段落（连续非空行合并）
        para = [stripped]
        i += 1
        while i < len(lines):
            nxt = lines[i].strip()
            if (not nxt or nxt.startswith(('#', '>', '- ', '* ', '|', '```'))
                    or re.match(r'^\d+\.\s', nxt) or nxt in ('---', '***')):
                break
            para.append(nxt)
            i += 1
        out.append(f'<p>{inline(" ".join(para))}</p>')
    close_lists()
    return '\n'.join(out)


HTML_TMPL = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8"/>
<meta name="generator" content="香料管理系统·AI调香专家 (香韵大师)"/>
<meta name="model" content="{model}"/>
<title>{title}</title>
<style>
body {{ font-family: 'Microsoft YaHei', sans-serif; margin: 0;
       background: #faf7f2; color: #3E2723; line-height: 1.7; }}
.header {{ background: linear-gradient(135deg, #3E2723 0%, #6F4E37 60%, #8D6E63 100%);
           color: #fff; padding: 34px 40px; }}
.header h1 {{ margin: 0 0 10px; font-size: 26px; }}
.chip {{ display: inline-block; background: rgba(255,255,255,.18);
         border: 1px solid rgba(255,255,255,.35); border-radius: 999px;
         padding: 3px 14px; margin: 4px 6px 0 0; font-size: 13px; }}
.content {{ max-width: 980px; margin: 24px auto 60px; padding: 0 24px; }}
.card {{ background: #fff; border: 1px solid #e7dccd; border-radius: 12px;
         padding: 26px 30px; margin-bottom: 22px;
         box-shadow: 0 2px 8px rgba(62,39,35,.06); }}
h2 {{ color: #6F4E37; border-left: 5px solid #D4A574; padding-left: 12px; }}
h3 {{ color: #8D6E63; }}
table {{ border-collapse: collapse; width: 100%; margin: 12px 0;
         font-size: 13.5px; }}
th {{ background: #6F4E37; color: #fff; padding: 8px 10px; text-align: left; }}
td {{ border-bottom: 1px solid #eadfce; padding: 7px 10px; }}
tr:nth-child(even) td {{ background: #faf5ec; }}
code {{ background: #f2e8da; padding: 1px 6px; border-radius: 4px;
        font-size: 13px; }}
pre code {{ display: block; padding: 12px; overflow-x: auto; }}
blockquote {{ border-left: 4px solid #D4A574; margin: 10px 0;
              padding: 6px 14px; background: #faf5ec; color: #6D4C41; }}
.footer {{ text-align: center; color: #9b8b78; font-size: 12.5px;
           padding: 18px; }}
</style>
</head>
<body>
<div class="header">
  <h1>🧪 {title}</h1>
  <div>
    <span class="chip">📅 {date}</span>
    <span class="chip">🤖 {model}</span>
    <span class="chip">系统：香料管理系统 · AI调香专家</span>
  </div>
</div>
<div class="content">{body}</div>
<div class="footer">香韵大师 · 由香料管理系统 AI调香专家生成 · 模型 {model}</div>
</body>
</html>"""


def wrap_html(title, md_body, model_slug):
    body = md_to_html(md_body)
    return HTML_TMPL.format(title=html_mod.escape(title),
                            date=datetime.now().strftime('%Y-%m-%d %H:%M'),
                            model=html_mod.escape(model_slug or '未知'),
                            body=body)


# ---------------------------------------------------------------- 设置对话框

class ApiSettingsDialog(QDialog):
    """大模型 API 设置"""

    def __init__(self, cfg, parent=None):
        super().__init__(parent)
        self.cfg = dict(cfg)
        self.setWindowTitle('AI调香专家 · 模型设置')
        self.setMinimumWidth(520)
        self.test_worker = None

        layout = QVBoxLayout(self)
        form = QFormLayout()

        self.provider_combo = QComboBox()
        self.provider_combo.addItems(list(PROVIDER_PRESETS.keys()))
        form.addRow('服务商预设:', self.provider_combo)

        self.base_edit = QLineEdit(cfg.get('base_url', ''))
        form.addRow('API地址(Base URL):', self.base_edit)

        self.key_edit = QLineEdit(cfg.get('api_key', ''))
        self.key_edit.setEchoMode(QLineEdit.EchoMode.Password)
        form.addRow('API Key:', self.key_edit)

        self.model_edit = QLineEdit(cfg.get('model', ''))
        self.model_edit.setPlaceholderText('如 glm-4-flash / deepseek-chat')
        form.addRow('模型名称:', self.model_edit)

        self.temp_spin = QDoubleSpinBox()
        self.temp_spin.setRange(0.0, 2.0)
        self.temp_spin.setSingleStep(0.1)
        self.temp_spin.setValue(float(cfg.get('temperature', 0.7)))
        form.addRow('创造性(temperature):', self.temp_spin)

        layout.addLayout(form)

        hint = QLabel('支持所有 OpenAI 兼容接口（智谱/DeepSeek/Kimi/OpenAI/'
                      'Ollama本地部署等）。API Key 保存在本机配置文件中，'
                      '不会上传。')
        hint.setWordWrap(True)
        hint.setStyleSheet('color:#888; font-size:12px;')
        layout.addWidget(hint)

        btn_row = QHBoxLayout()
        self.test_btn = QPushButton('🔍 测试连接')
        self.test_btn.clicked.connect(self.test_connection)
        btn_row.addWidget(self.test_btn)
        btn_row.addStretch()
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok |
                                   QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        btn_row.addWidget(buttons)
        layout.addLayout(btn_row)

        # 根据已存配置猜测预设
        for name, p in PROVIDER_PRESETS.items():
            if p['base_url'] and p['base_url'] == cfg.get('base_url', ''):
                self.provider_combo.setCurrentText(name)
                break
        self.provider_combo.currentTextChanged.connect(self.apply_preset)

    def apply_preset(self, name):
        preset = PROVIDER_PRESETS.get(name)
        if preset:
            self.base_edit.setText(preset['base_url'])
            if preset['model']:
                self.model_edit.setText(preset['model'])

    def test_connection(self):
        self.test_btn.setEnabled(False)
        self.test_btn.setText('测试中…')
        self.test_worker = ConnectionTestWorker(
            self.base_edit.text().strip(), self.key_edit.text().strip(),
            self.model_edit.text().strip())
        self.test_worker.ok.connect(self._test_ok)
        self.test_worker.error.connect(self._test_err)
        self.test_worker.start()

    def _test_ok(self, reply):
        self.test_btn.setEnabled(True)
        self.test_btn.setText('🔍 测试连接')
        QMessageBox.information(self, '连接成功', f'模型回复：{reply}')

    def _test_err(self, msg):
        self.test_btn.setEnabled(True)
        self.test_btn.setText('🔍 测试连接')
        QMessageBox.warning(self, '连接失败', msg)

    def get_config(self):
        return {
            'base_url': self.base_edit.text().strip(),
            'api_key': self.key_edit.text().strip(),
            'model': self.model_edit.text().strip(),
            'temperature': self.temp_spin.value(),
        }


# ---------------------------------------------------------------- 记忆库对话框

class MemoryDialog(QDialog):
    """记忆库管理：查看生效知识、管理候选条目"""

    def __init__(self, memory: MemoryStore, parent=None):
        super().__init__(parent)
        self.memory = memory
        self.setWindowTitle('香韵大师 · 记忆库管理')
        self.resize(720, 520)

        layout = QVBoxLayout(self)
        tabs = QTabWidget()

        # 生效知识
        k_view = QPlainTextEdit()
        k_view.setReadOnly(True)
        k_view.setPlainText(self.memory._read('knowledge'))
        tabs.addTab(k_view, '生效知识 [K]')

        # 候选库
        cand_widget = QWidget()
        cand_layout = QVBoxLayout(cand_widget)
        self.cand_list = QListWidget()
        for kid, text, _raw in self.memory.get_candidates():
            self.cand_list.addItem(f'{kid}  {text}')
        cand_layout.addWidget(QLabel('候选知识需确认后才会进入生效库'
                                     '（防止未验证内容污染专家行为）：'))
        cand_layout.addWidget(self.cand_list)
        btn_row = QHBoxLayout()
        promote_btn = QPushButton('✅ 升格为生效知识')
        promote_btn.clicked.connect(self.promote)
        delete_btn = QPushButton('🗑 删除候选')
        delete_btn.clicked.connect(self.delete_cand)
        btn_row.addWidget(promote_btn)
        btn_row.addWidget(delete_btn)
        btn_row.addStretch()
        cand_layout.addLayout(btn_row)
        tabs.addTab(cand_widget, f'候选库 [KC]（{self.cand_list.count()}）')

        # 偏好档案
        p_view = QPlainTextEdit()
        p_view.setReadOnly(True)
        p_view.setPlainText(self.memory._read('profile'))
        tabs.addTab(p_view, '偏好档案')

        # 反馈日志
        f_view = QPlainTextEdit()
        f_view.setReadOnly(True)
        f_view.setPlainText(self.memory._read('feedback'))
        tabs.addTab(f_view, '反馈日志')

        layout.addWidget(tabs)
        bottom = QHBoxLayout()
        open_btn = QPushButton('📂 打开记忆库文件夹')
        open_btn.clicked.connect(lambda: os.startfile(self.memory.base_dir))
        bottom.addWidget(open_btn)
        bottom.addStretch()
        close_btn = QPushButton('关闭')
        close_btn.clicked.connect(self.accept)
        bottom.addWidget(close_btn)
        layout.addLayout(bottom)

    def promote(self):
        row = self.cand_list.currentRow()
        if row < 0:
            QMessageBox.information(self, '提示', '请先选择一条候选知识')
            return
        kid = self.cand_list.item(row).text().split()[0]
        if self.memory.promote(kid):
            self.cand_list.takeItem(row)
            QMessageBox.information(self, '已升格',
                                    f'{kid} 已进入生效知识库，'
                                    f'后续对话将自动应用。')

    def delete_cand(self):
        row = self.cand_list.currentRow()
        if row < 0:
            return
        kid = self.cand_list.item(row).text().split()[0]
        self.memory.delete_candidate(kid)
        self.cand_list.takeItem(row)


# ---------------------------------------------------------------- 主模块

QUICK_SCENARIOS = [
    ('逆向仿香', '我想逆向仿香一款香精。\n目标香型：\n竞品信息（品牌/描述/GC-MS数据）：\n目标场景：电子雾化香精\n请先列出你需要我补充的信息。'),
    ('新香型开发', '我想开发一个新香型香精。\n目标香型：\n感官目标描述：\n目标场景：电子雾化香精\n请给出香韵分路体系与开发思路。'),
    ('配方合规改造', '请对我的配方做合规改造。\n目标场景：电子雾化香精（GB 41700）\n配方内容：\n请指出受限组分并给出等效替代方案，输出双版本配方。'),
    ('雾化问题诊断', '我的电子雾化香精量产出现以下问题：\n（析晶/分层/热分解杂味/积碳过快/击喉刺激等）\n现象描述：\n配方概要：\n请诊断原因并给出解决方案。'),
    ('单体咨询', '请介绍以下香料单体：\n（名称或CAS号）\n包括感官特征、香气阈值、在电子雾化/食用香精中的应用要点与使用限量。'),
]


class FlavorExpertModule(QWidget):
    """AI调香专家页面"""

    def __init__(self, session, parent=None):
        super().__init__(parent)
        self.session = session
        self.memory = MemoryStore()
        self.cfg = load_config()
        self.history = []          # [{'role','content'}]
        self.worker = None
        self._buffer = ''
        self._full_answer = ''
        self._pending_export_aroma = None
        self._render_timer = QTimer(self)
        self._render_timer.setInterval(150)
        self._render_timer.timeout.connect(self._render_chat)
        self._init_ui()

    # ---------------- UI ----------------
    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)

        # 顶部标题栏
        header = QHBoxLayout()
        title = QLabel('🧪 AI调香专家 · 香韵大师')
        title.setStyleSheet('font-size: 18px; font-weight: bold; color: #3E2723;')
        header.addWidget(title)
        self.status_label = QLabel()
        self.status_label.setStyleSheet('font-size: 12px; color: #666;')
        header.addWidget(self.status_label)
        header.addStretch()

        btn_style = ('QPushButton { background-color: #6F4E37; color: white; '
                     'font-weight: bold; padding: 6px 12px; border-radius: 4px; }'
                     'QPushButton:hover { background-color: #8D6E63; }')
        settings_btn = QPushButton('⚙️ 模型设置')
        settings_btn.setStyleSheet(btn_style)
        settings_btn.clicked.connect(self.open_settings)
        memory_btn = QPushButton('📚 记忆库')
        memory_btn.setStyleSheet(btn_style)
        memory_btn.clicked.connect(self.open_memory)
        export_btn = QPushButton('📤 导出报告')
        export_btn.setStyleSheet(btn_style)
        export_btn.clicked.connect(self.export_report)
        clear_btn = QPushButton('🧹 新对话')
        clear_btn.setStyleSheet(btn_style)
        clear_btn.clicked.connect(self.clear_chat)
        for b in (settings_btn, memory_btn, export_btn, clear_btn):
            header.addWidget(b)
        layout.addLayout(header)

        splitter = QSplitter(Qt.Orientation.Horizontal)

        # 左：对话区
        chat_widget = QWidget()
        chat_layout = QVBoxLayout(chat_widget)
        chat_layout.setContentsMargins(0, 0, 0, 0)

        self.chat_view = QTextEdit()
        self.chat_view.setReadOnly(True)
        self.chat_view.setStyleSheet(
            'QTextEdit { background: #fffdf9; border: 1px solid #e0d5c5;'
            ' border-radius: 8px; padding: 10px; }')
        f = QFont('Microsoft YaHei', 10)
        self.chat_view.setFont(f)
        chat_layout.addWidget(self.chat_view, 1)

        self._render_welcome()

        # 快捷场景
        quick_row = QHBoxLayout()
        quick_row.addWidget(QLabel('快捷场景:'))
        for name, tpl in QUICK_SCENARIOS:
            qb = QPushButton(name)
            qb.setStyleSheet(
                'QPushButton { padding: 4px 10px; border: 1px solid #D4A574;'
                ' border-radius: 10px; background: #faf5ec; }'
                'QPushButton:hover { background: #f3e7d3; }')
            qb.clicked.connect(lambda _, t=tpl: self._fill_input(t))
            quick_row.addWidget(qb)
        quick_row.addStretch()
        chat_layout.addLayout(quick_row)

        # 输入区
        input_row = QHBoxLayout()
        self.input_edit = QTextEdit()
        self.input_edit.setPlaceholderText(
            '向香韵大师提问…（Ctrl+Enter 发送）\n'
            '例：请根据右侧附上的GC-MS数据做逆向仿香分析，输出双版本配方')
        self.input_edit.setMaximumHeight(96)
        input_row.addWidget(self.input_edit, 1)
        btn_col = QVBoxLayout()
        self.send_btn = QPushButton('🚀 发送')
        self.send_btn.setStyleSheet(
            'QPushButton { background-color: #3E2723; color: white;'
            ' font-weight: bold; padding: 10px 18px; border-radius: 6px; }'
            'QPushButton:disabled { background-color: #b9a89a; }')
        self.send_btn.clicked.connect(self.send_message)
        self.stop_btn = QPushButton('⏹ 停止')
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self.stop_generation)
        btn_col.addWidget(self.send_btn)
        btn_col.addWidget(self.stop_btn)
        input_row.addLayout(btn_col)
        chat_layout.addLayout(input_row)
        splitter.addWidget(chat_widget)

        # 右：上下文附加面板
        ctx_widget = QWidget()
        ctx_layout = QVBoxLayout(ctx_widget)
        ctx_group = QGroupBox('📎 附上系统数据（作为专家上下文）')
        ctx_inner = QVBoxLayout(ctx_group)

        self.chk_ingredients = QCheckBox('原料库摘要（编号/名称/CAS/香型/价格）')
        self.chk_formulas = QCheckBox('配方库摘要（全部配方组成）')
        self.chk_gcms_all = QCheckBox('全部GC-MS分析摘要')
        ctx_inner.addWidget(self.chk_ingredients)
        ctx_inner.addWidget(self.chk_formulas)
        ctx_inner.addWidget(self.chk_gcms_all)

        ctx_inner.addWidget(QLabel('指定单个配方（完整组成）:'))
        self.formula_combo = QComboBox()
        ctx_inner.addWidget(self.formula_combo)
        refresh_formula_btn = QPushButton('🔄 刷新配方列表')
        refresh_formula_btn.clicked.connect(self.load_formula_combo)
        ctx_inner.addWidget(refresh_formula_btn)

        ctx_inner.addWidget(QLabel('指定单个GC-MS分析（含化合物表）:'))
        self.gcms_combo = QComboBox()
        ctx_inner.addWidget(self.gcms_combo)
        refresh_gcms_btn = QPushButton('🔄 刷新GC-MS列表')
        refresh_gcms_btn.clicked.connect(self.load_gcms_combo)
        ctx_inner.addWidget(refresh_gcms_btn)

        self.report_mode = QCheckBox('按九大章节输出完整配方报告')
        self.report_mode.setToolTip('勾选后发送的问题将按「香韵大师」配方报告'
                                    '模板（含双版本配方/分路分析/合规审查/'
                                    '雾化适配）输出')
        ctx_inner.addWidget(self.report_mode)

        tip = QLabel('💡 提示：先在"模型设置"里配置 API Key。'
                     '勾选的数据仅在对话时发送给所配置的大模型服务商。')
        tip.setWordWrap(True)
        tip.setStyleSheet('color:#888; font-size:11.5px;')
        ctx_inner.addWidget(tip)
        ctx_inner.addStretch()
        ctx_layout.addWidget(ctx_group)
        ctx_layout.addStretch()
        splitter.addWidget(ctx_widget)

        splitter.setStretchFactor(0, 4)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([760, 300])
        layout.addWidget(splitter, 1)

        self.load_formula_combo()
        self.load_gcms_combo()
        self._update_status()

    def _render_welcome(self):
        self._render_buffer_text(
            '### 🧪 欢迎使用 AI调香专家（香韵大师）\n\n'
            '我聚焦**电子雾化香精**与**食用香精**，可帮你完成：\n\n'
            '- 🔬 **GC-MS全谱解析**：组分定性定量、OT值校正、特征指纹识别\n'
            '- 🌬 **香韵结构重组**：五层骨架 + 香韵分路体系设计\n'
            '- ⚖️ **双版本配方**：极致还原版 + 合规落地版（GB 2760 / GB 41700 / TPD）\n'
            '- 💨 **雾化工程化**：溶解度、热稳定、积碳与体感优化\n\n'
            '**开始前**：点击右上角「⚙️ 模型设置」配置大模型 API。\n'
            '可在右侧勾选附上系统内的原料库 / 配方库 / GC-MS 数据，'
            '我会基于你的真实数据工作。')

    # ---------------- 数据上下文 ----------------
    def load_formula_combo(self):
        try:
            from models import Formula
            self.formula_combo.clear()
            self.formula_combo.addItem('（不指定）', None)
            for fm in self.session.query(Formula).order_by(
                    Formula.updated_at.desc()).all():
                self.formula_combo.addItem(fm.name, fm.id)
        except Exception:
            pass

    def load_gcms_combo(self):
        try:
            from models import GCMSAnalysis
            self.gcms_combo.clear()
            self.gcms_combo.addItem('（不指定）', None)
            for a in self.session.query(GCMSAnalysis).order_by(
                    GCMSAnalysis.analysis_time.desc()).all():
                self.gcms_combo.addItem(f'{a.number} {a.name or ""}', a.id)
        except Exception:
            pass

    def _build_attachments(self):
        from models import (Ingredient, Formula, GCMSAnalysis, GCMSCompound,
                            ingredient_formula)
        blocks = []

        if self.chk_ingredients.isChecked():
            ings = self.session.query(Ingredient).all()
            lines = ['| 编号 | 名称 | CAS号 | 香型特征 | 价格 |']
            for g in ings[:120]:
                lines.append(f"| {g.number} | {g.name} | {g.cas_number or ''} "
                             f"| {(g.aroma_character or '')[:40]} "
                             f"| {g.price or 0} |")
            if len(ings) > 120:
                lines.append(f'（共{len(ings)}种，仅列出前120种）')
            blocks.append('**原料库摘要**\n' + '\n'.join(lines))

        if self.chk_formulas.isChecked():
            fs = self.session.query(Formula).all()
            parts = []
            for fm in fs[:40]:
                names = [i.name for i in fm.ingredients]
                parts.append(f'- {fm.name}（{fm.number or ""}）: '
                             f'{fm.content or ""} '
                             f'{" 成分:" + "、".join(names) if names else ""}')
            blocks.append('**配方库摘要**\n' + '\n'.join(parts))

        if self.chk_gcms_all.isChecked():
            ans = self.session.query(GCMSAnalysis).all()[:20]
            parts = []
            for a in ans:
                top = self.session.query(GCMSCompound).filter(
                    GCMSCompound.analysis_id == a.id).order_by(
                    GCMSCompound.relative_content.desc()).limit(12).all()
                comp = '；'.join(f'{c.name_cn or c.name_en}({c.relative_content}%)'
                                for c in top)
                parts.append(f'- {a.number} {a.name or ""}: {comp}')
            blocks.append('**GC-MS分析摘要**\n' + '\n'.join(parts))

        fid = self.formula_combo.currentData()
        if fid:
            fm = self.session.query(Formula).get(fid)
            if fm:
                rows = self.session.query(ingredient_formula).filter_by(
                    formula_id=fm.id).all()
                ing_ids = [r['ingredient_id'] for r in (
                    dict(r._mapping) for r in rows)]
                ings = self.session.query(Ingredient).filter(
                    Ingredient.id.in_(ing_ids)).all() if ing_ids else []
                pct = {dict(r._mapping)['ingredient_id']:
                       dict(r._mapping).get('percentage') for r in rows}
                detail = [f'配方：{fm.name}（{fm.number or ""}） 版本:'
                          f'{fm.version or ""} 创建者:{fm.creator or ""}',
                          f'描述: {fm.description or ""}',
                          f'内容: {fm.content or ""}']
                tbl = ['| 原料 | CAS | 用量(%) |', '|---|---|---|']
                for g in ings:
                    tbl.append(f'| {g.name} | {g.cas_number or ""} '
                               f'| {pct.get(g.id, "")} |')
                detail.append('\n'.join(tbl))
                blocks.append('**指定配方完整信息**\n' + '\n'.join(detail))

        aid = self.gcms_combo.currentData()
        if aid:
            a = self.session.query(GCMSAnalysis).get(aid)
            if a:
                comps = self.session.query(GCMSCompound).filter(
                    GCMSCompound.analysis_id == a.id).order_by(
                    GCMSCompound.relative_content.desc()).all()
                info = [f'GC-MS分析：{a.number} {a.name or ""}',
                        f'供应商: {a.supplier or ""}｜仪器参数: '
                        f'{(a.instrument_params or "")[:200]}',
                        f'调香创意思路: {a.perfume_idea or ""}']
                tbl = ['| CAS | 英文名 | 中文名 | 保留时间 | 匹配度 | '
                       '分子式 | 相对含量(%) |',
                       '|---|---|---|---|---|---|---|']
                for c in comps[:60]:
                    rc = (f'{c.relative_content}' 
                          if c.relative_content is not None else '')
                    tbl.append(f'| {c.cas or ""} | {c.name_en or ""} '
                               f'| {c.name_cn or ""} | {c.rt or ""} '
                               f'| {c.match_factor or ""} | {c.formula or ""} '
                               f'| {rc} |')
                if len(comps) > 60:
                    tbl.append(f'（共{len(comps)}个组分，按含量列出前60）')
                info.append('\n'.join(tbl))
                blocks.append('**指定GC-MS分析数据**\n' + '\n'.join(info))

        return blocks

    # ---------------- 对话 ----------------
    def send_message(self):
        if self.worker is not None and self.worker.isRunning():
            return
        if not self.cfg.get('api_key') or not self.cfg.get('base_url') \
                or not self.cfg.get('model'):
            QMessageBox.warning(self, '请先配置模型',
                                '使用前请点击「⚙️ 模型设置」填写 API 地址、'
                                'Key 与模型名称。\n支持智谱GLM / DeepSeek / '
                                'Kimi / OpenAI 等任意 OpenAI 兼容接口。')
            self.open_settings()
            return

        text = self.input_edit.toPlainText().strip()
        if not text:
            return
        self.input_edit.clear()

        attachments = self._build_attachments()
        content = text
        if self.report_mode.isChecked():
            content = ('请按九大章节配方报告规范输出完整报告，'
                       '含「极致还原版」与「合规落地版」双版本。\n\n' + text)
        if attachments:
            content += ('\n\n---\n**【香料管理系统数据】**\n'
                        + '\n\n'.join(attachments))

        self.history.append({'role': 'user', 'content': content})
        self._append_history_render('user', text)

        messages = [{'role': 'system',
                     'content': SYSTEM_PROMPT
                     + self.memory.load_context()}]
        # 只保留最近12条历史，控制 token
        for m in self.history[-12:]:
            messages.append({'role': m['role'], 'content': m['content']})

        self._buffer = ''
        self._full_answer = ''
        self.send_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.worker = ExpertWorker(self.cfg['base_url'], self.cfg['api_key'],
                                   self.cfg['model'],
                                   float(self.cfg.get('temperature', 0.7)),
                                   messages)
        self.worker.chunk.connect(self._on_chunk)
        self.worker.finished_ok.connect(self._on_finished)
        self.worker.failed.connect(self._on_failed)
        self._render_timer.start()
        self.worker.start()

    def stop_generation(self):
        if self.worker is not None:
            self.worker.stop()

    def _on_chunk(self, piece):
        self._buffer += piece
        self._full_answer += piece

    def _on_finished(self, text):
        self._render_timer.stop()
        self.history.append({'role': 'assistant', 'content': text})
        self._append_history_render('assistant', text)
        self._finalize_turn(text)

    def _on_failed(self, msg):
        self._render_timer.stop()
        self.send_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self._append_history_render('error', f'❌ 请求失败：{msg}')
        if self.history and self.history[-1]['role'] == 'user':
            self.history.pop()  # 失败时不保留用户消息，便于重发
        self._render_chat()

    def _finalize_turn(self, text):
        """一轮完成：沉淀候选知识、记忆反馈、可能的报告自动导出"""
        self.send_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        try:
            entries = parse_candidates(text)
            added = self.memory.add_candidates(entries)
            if added:
                ids = '、'.join(k for k, _ in added)
                self._append_history_render(
                    'system', f'💾 已提炼候选知识 {ids}（{len(added)}条），'
                              f'进入「📚 记忆库 → 候选库」确认后即可升格为生效知识。')
                self._render_chat()
        except Exception:
            pass
        if self._pending_export_aroma:
            aroma = self._pending_export_aroma
            self._pending_export_aroma = None
            self.export_report(aroma)

    # ---------------- 渲染 ----------------
    def _transcript_md(self):
        labels = {'user': '### 🧑 用户', 'assistant': '### 🧙 香韵大师',
                  'error': '#### ❌ 系统提示', 'system': '#### ℹ️ 系统'}
        parts = []
        for m in self._render_items:
            parts.append(labels.get(m['role'], '### ℹ️')
                         + '\n\n' + m['display'])
        return '\n\n---\n\n'.join(parts)

    def _render_chat(self):
        md = self._transcript_md()
        self.chat_view.setMarkdown(md)
        sb = self.chat_view.verticalScrollBar()
        sb.setValue(sb.maximum())

    def _render_buffer_text(self, text):
        self._render_items = [{'role': 'assistant', 'display': text}]
        self.chat_view.setMarkdown(text)
        sb = self.chat_view.verticalScrollBar()
        sb.setValue(sb.maximum())

    def _append_history_render(self, role, display):
        if not hasattr(self, '_render_items'):
            self._render_items = []
        self._render_items.append({'role': role, 'display': display})
        self._render_chat()

    # ---------------- 导出报告 ----------------
    def export_report(self, aroma=None):
        if not getattr(self, '_render_items', None):
            QMessageBox.information(self, '提示', '当前没有可导出的对话内容。')
            return
        if not aroma:
            aroma, ok = QInputDialog.getText(
                self, '导出报告', '香型/主题（用于归档目录，如：咖啡）:',
                text='调香咨询')
            if not ok or not aroma.strip():
                return
            aroma = aroma.strip()

        model_slug = re.sub(r'[^a-zA-Z0-9._-]+', '-',
                            self.cfg.get('model') or 'unknown').strip('-')
        scene = '电子雾化' if '雾化' in '\n'.join(
            m['display'][:200] for m in self._render_items) else '通用'
        aroma_en = re.sub(r'[^a-z]+', '', aroma.lower()) or 'chat'
        day = datetime.now().strftime('%Y-%m-%d')
        folder = f'{day}_{aroma}_{scene}_{model_slug}'
        full_dir = os.path.join(DELIVERY_DIR, aroma_en, folder)
        try:
            os.makedirs(full_dir, exist_ok=True)
            md_path = os.path.join(full_dir, 'report.md')
            with open(md_path, 'w', encoding='utf-8') as f:
                f.write(f'# {aroma} — AI调香专家对话记录\n\n'
                        f'> 生成时间：{datetime.now():%Y-%m-%d %H:%M}｜'
                        f'模型：{model_slug}\n\n'
                        + self._transcript_md())
            html_path = os.path.join(full_dir, 'report.html')
            with open(html_path, 'w', encoding='utf-8') as f:
                f.write(wrap_html(f'{aroma} — 调香对话报告',
                                  self._transcript_md(), model_slug))
            self.memory.append_index('对话报告', aroma, model_slug,
                                     f'deliveries/{aroma_en}/{folder}')
            self.memory.append_feedback(f'导出对话报告：{aroma}', verified=False)
            QMessageBox.information(
                self, '导出成功',
                f'报告已保存：\n{md_path}\n{html_path}\n\n'
                f'（已写入记忆库 INDEX）')
        except Exception as e:
            QMessageBox.warning(self, '导出失败', str(e))

    # ---------------- 其他动作 ----------------
    def open_settings(self):
        dlg = ApiSettingsDialog(self.cfg, self)
        if dlg.exec():
            self.cfg = dlg.get_config()
            save_config(self.cfg)
            self._update_status()

    def open_memory(self):
        MemoryDialog(self.memory, self).exec()

    def clear_chat(self):
        self.history = []
        self._render_items = []
        self._render_welcome()

    def _fill_input(self, tpl):
        self.input_edit.setPlainText(tpl)
        self.input_edit.setFocus()

    def _update_status(self):
        if self.cfg.get('api_key') and self.cfg.get('model'):
            self.status_label.setText(
                f"🟢 已连接模型：{self.cfg.get('model')}")
            self.status_label.setStyleSheet('font-size:12px; color:#2e7d32;')
        else:
            self.status_label.setText('🔴 未配置模型，请点击「⚙️ 模型设置」')
            self.status_label.setStyleSheet('font-size:12px; color:#c62828;')
