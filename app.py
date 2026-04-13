#!/usr/bin/env python3
"""
配电设备成本计算系统 - 多柜项目模式
基于Excel公式逻辑，自动计算铜排成本和项目总价
支持多台柜子独立计算和项目汇总
"""

import streamlit as st
import pandas as pd
import json
import os
import re
import io
import sqlite3
from contextlib import contextmanager
from pathlib import Path
from datetime import datetime
import openpyxl
import requests

# Telegram通知配置
TELEGRAM_BOT_TOKEN = "8626932429:AAE8Muk7WSPgqgJQlvyg5PcLPqaFoByyUHk"
TELEGRAM_CHAT_ID = "7473762677"  # 张东方
TELEGRAM_NOTIFY = True  # 设为False可关闭通知

# ─── 配置 ───────────────────────────────────────────────
st.set_page_config(
    page_title="配电设备成本计算系统",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"

CABINET_TYPES = ["进线柜", "出线柜", "联络柜", "计量柜", "无功补偿柜", "电容柜", "其他"]

# ─── 铜价获取 ─────────────────────────────────────────────
def fetch_copper_price():
    """从新浪财经获取沪铜主力合约卖出价"""
    import urllib.request
    url = "https://hq.sinajs.cn/list=nf_CU0"
    req = urllib.request.Request(url, headers={"Referer": "https://finance.sina.com.cn"})
    try:
        with urllib.request.urlopen(req, timeout=5) as resp:
            data = resp.read().decode('gbk')
            start = data.index('="') + 2
            end = data.rindex('"')
            fields = data[start:end].split(',')
            sell_price = float(fields[7]) if fields[7] else 0
            date_str = fields[16] if len(fields) > 16 else ''
            return {'price': sell_price, 'date': date_str}
    except Exception:
        return None

@st.cache_data(ttl=3600)
def get_cached_copper_price():
    """缓存1小时的铜价"""
    return fetch_copper_price()

# ─── SQLite 数据库 ────────────────────────────────────────
DB_PATH = DATA_DIR / "price.db"

@contextmanager
def get_db():
    conn = sqlite3.connect(str(DB_PATH), timeout=10)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA busy_timeout=5000")
    try:
        yield conn
    finally:
        conn.close()

def init_db():
    with get_db() as conn:
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS products (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                model TEXT NOT NULL,
                name TEXT DEFAULT '',
                unit_price REAL DEFAULT 0,
                retail_price REAL DEFAULT 0,
                brand TEXT DEFAULT '',
                created_at TEXT DEFAULT (datetime('now','localtime')),
                updated_at TEXT DEFAULT (datetime('now','localtime'))
            );
            CREATE TABLE IF NOT EXISTS price_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                action TEXT NOT NULL,
                model TEXT,
                old_value TEXT,
                new_value TEXT,
                operator TEXT DEFAULT '',
                created_at TEXT DEFAULT (datetime('now','localtime'))
            );
        """)

def migrate_from_json():
    json_path = DATA_DIR / "price_db.json"
    if not json_path.exists():
        return
    with open(json_path, 'r', encoding='utf-8') as f:
        items = json.load(f)
    with get_db() as conn:
        for item in items:
            model = str(item.get('型号', '')).strip()
            if not model:
                continue
            conn.execute(
                "INSERT OR IGNORE INTO products (model, name, unit_price, retail_price, brand) VALUES (?,?,?,?,?)",
                (model, str(item.get('名称', '')).strip(),
                 float(item.get('单价', 0) or 0),
                 float(item.get('面价', 0) or 0),
                 str(item.get('产地/厂家', '')).strip()))
        conn.commit()
    import shutil
    shutil.copy2(str(json_path), str(json_path) + ".bak")

def get_all_products(search=''):
    with get_db() as conn:
        if search:
            rows = conn.execute("SELECT * FROM products WHERE model LIKE ? OR name LIKE ? OR brand LIKE ?",
                              (f'%{search}%', f'%{search}%', f'%{search}%')).fetchall()
        else:
            rows = conn.execute("SELECT * FROM products").fetchall()
        return [dict(r) for r in rows]

def get_all_brands():
    with get_db() as conn:
        rows = conn.execute("SELECT DISTINCT brand FROM products WHERE brand != '' ORDER BY brand").fetchall()
        return [r['brand'] for r in rows]

def lookup_price_sqlite(model):
    """按型号查找价格，先精确匹配，再模糊匹配。"""
    with get_db() as conn:
        # 1. 精确匹配
        row = conn.execute("SELECT * FROM products WHERE model = ?", (model,)).fetchone()
        if row:
            d = dict(row)
            return {
                'name': d['name'],
                'unit_price': float(d['unit_price'] or 0),
                'retail_price': float(d['retail_price'] or 0),
                'brand': d['brand'],
                'type': d.get('type', ''),
            }
        # 2. 模糊匹配：去掉空格后匹配
        model_nospace = model.replace(' ', '').upper()
        rows = conn.execute("SELECT * FROM products WHERE REPLACE(UPPER(model), ' ', '') LIKE ?", (f"%{model_nospace}%",)).fetchall()
        if rows:
            d = dict(rows[0])
            return {
                'name': d['name'],
                'unit_price': float(d['unit_price'] or 0),
                'retail_price': float(d['retail_price'] or 0),
                'brand': d['brand'],
                'type': d.get('type', ''),
            }
    return None

def lookup_price_by_name_sqlite(name):
    """按名称查找价格，先精确匹配，再模糊匹配。"""
    with get_db() as conn:
        # 1. 精确匹配（名称或型号）
        row = conn.execute("SELECT * FROM products WHERE name = ? OR model = ?", (name, name)).fetchone()
        if row:
            d = dict(row)
            return {
                'name': d['name'] or d['model'],
                'unit_price': float(d['unit_price'] or 0),
                'retail_price': float(d['retail_price'] or 0),
                'brand': d['brand'],
                'type': d.get('type', ''),
            }
        # 2. 模糊匹配（名称或型号）
        rows = conn.execute("SELECT * FROM products WHERE name LIKE ? OR model LIKE ?", (f"%{name}%", f"%{name}%")).fetchall()
        if rows:
            d = dict(rows[0])
            return {
                'name': d['name'] or d['model'],
                'unit_price': float(d['unit_price'] or 0),
                'retail_price': float(d['retail_price'] or 0),
                'brand': d['brand'],
                'type': d.get('type', ''),
            }
    return None

def insert_product(model, name, unit_price, retail_price, brand):
    with get_db() as conn:
        conn.execute("INSERT INTO products (model, name, unit_price, retail_price, brand) VALUES (?,?,?,?,?)",
                    (model, name, unit_price, retail_price, brand))
        conn.execute("INSERT INTO price_log (action, model, new_value) VALUES ('insert', ?, ?)",
                    (model, json.dumps({'name':name,'unit_price':unit_price}, ensure_ascii=False)))
        conn.commit()

def update_product(pid, name, unit_price, retail_price, brand):
    with get_db() as conn:
        old = conn.execute("SELECT * FROM products WHERE id=?", (pid,)).fetchone()
        old_json = json.dumps(dict(old), ensure_ascii=False) if old else ''
        conn.execute("UPDATE products SET name=?, unit_price=?, retail_price=?, brand=?, updated_at=datetime('now','localtime') WHERE id=?",
                    (name, unit_price, retail_price, brand, pid))
        new = conn.execute("SELECT * FROM products WHERE id=?", (pid,)).fetchone()
        new_json = json.dumps(dict(new), ensure_ascii=False) if new else ''
        conn.execute("INSERT INTO price_log (action, model, old_value, new_value) VALUES ('update', ?, ?, ?)",
                    (old['model'] if old else '', old_json, new_json))
        conn.commit()

def delete_product(pid):
    with get_db() as conn:
        old = conn.execute("SELECT * FROM products WHERE id=?", (pid,)).fetchone()
        if old:
            conn.execute("INSERT INTO price_log (action, model, old_value) VALUES ('delete', ?, ?)",
                        (old['model'], json.dumps(dict(old), ensure_ascii=False)))
            conn.execute("DELETE FROM products WHERE id=?", (pid,))
            conn.commit()

def bulk_import_products(items):
    with get_db() as conn:
        for item in items:
            existing = conn.execute("SELECT id FROM products WHERE model=?", (item['model'],)).fetchone()
            if existing:
                conn.execute("UPDATE products SET name=?, unit_price=?, retail_price=?, brand=?, updated_at=datetime('now','localtime') WHERE id=?",
                           (item.get('name',''), item.get('unit_price',0), item.get('retail_price',0), item.get('brand',''), existing['id']))
            else:
                conn.execute("INSERT INTO products (model, name, unit_price, retail_price, brand) VALUES (?,?,?,?,?)",
                           (item['model'], item.get('name',''), item.get('unit_price',0), item.get('retail_price',0), item.get('brand','')))
        conn.commit()

def get_price_logs(limit=50):
    with get_db() as conn:
        rows = conn.execute("SELECT * FROM price_log ORDER BY id DESC LIMIT ?", (limit,)).fetchall()
        return [dict(r) for r in rows]

# ─── 启动时检查数据库 ─────────────────────────────────────
if not DB_PATH.exists():
    init_db()
    migrate_from_json()
else:
    # 确保表存在
    init_db()

COPPER_SPECS = [
    {'spec': '15×2',  'width': 15, 'thickness': 2,  'area_mm2': 30,  'area_cm2': 0.3, 'current': 125},
    {'spec': '15×3',  'width': 15, 'thickness': 3,  'area_mm2': 45,  'area_cm2': 0.45, 'current': 185},
    {'spec': '20×3',  'width': 20, 'thickness': 3,  'area_mm2': 60,  'area_cm2': 0.6, 'current': 245},
    {'spec': '20×4',  'width': 20, 'thickness': 4,  'area_mm2': 80,  'area_cm2': 0.8, 'current': 320},
    {'spec': '25×3',  'width': 25, 'thickness': 3,  'area_mm2': 75,  'area_cm2': 0.75, 'current': 305},
    {'spec': '25×4',  'width': 25, 'thickness': 4,  'area_mm2': 100, 'area_cm2': 1.0, 'current': 370},
    {'spec': '30×3',  'width': 30, 'thickness': 3,  'area_mm2': 90,  'area_cm2': 0.9, 'current': 355},
    {'spec': '30×4',  'width': 30, 'thickness': 4,  'area_mm2': 120, 'area_cm2': 1.2, 'current': 420},
    {'spec': '40×4',  'width': 40, 'thickness': 4,  'area_mm2': 160, 'area_cm2': 1.6, 'current': 560},
    {'spec': '40×5',  'width': 40, 'thickness': 5,  'area_mm2': 200, 'area_cm2': 2.0, 'current': 615},
    {'spec': '50×5',  'width': 50, 'thickness': 5,  'area_mm2': 250, 'area_cm2': 2.5, 'current': 755},
    {'spec': '60×6',  'width': 60, 'thickness': 6,  'area_mm2': 360, 'area_cm2': 3.6, 'current': 990},
    {'spec': '60×8',  'width': 60, 'thickness': 8,  'area_mm2': 480, 'area_cm2': 4.8, 'current': 1160},
    {'spec': '80×8',  'width': 80, 'thickness': 8,  'area_mm2': 640, 'area_cm2': 6.4, 'current': 1490},
    {'spec': '80×10', 'width': 80, 'thickness': 10, 'area_mm2': 800, 'area_cm2': 8.0, 'current': 1670},
    {'spec': '100×10', 'width': 100, 'thickness': 10, 'area_mm2': 1000, 'area_cm2': 10.0, 'current': 2030},
    {'spec': '120×10', 'width': 120, 'thickness': 10, 'area_mm2': 1200, 'area_cm2': 12.0, 'current': 2330},
]

@st.cache_data
def load_breaker_cable_params():
    """加载断路器电缆参数: 电流(A) → 标准电缆截面积(mm²)
    参考 GB/T 7251 标准及行业常用对照表
    """
    return {
        63: 16, 80: 25, 100: 35, 125: 50, 160: 70,
        200: 95, 225: 120, 250: 120,  # 250A及以上用铜排，此值仅备用
    }


def get_breaker_poles(model: str) -> int:
    """从型号判断断路器极数: 1P/2P → 1(单相), 3P/4P → 3(三相), 默认3"""
    import re
    m = re.search(r'(\d)\s*[Pp]', model.upper())
    if m:
        poles = int(m.group(1))
        return 1 if poles <= 2 else 3
    return 3  # 默认三相

# ─── 核心计算函数 ─────────────────────────────────────────

def lookup_price_by_name(name: str, db=None) -> dict:
    """根据名称查找价格（兼容旧接口，db参数忽略）"""
    return lookup_price_by_name_sqlite(str(name).strip())

def lookup_price(model: str, db=None) -> dict:
    """根据型号查找价格（兼容旧接口，db参数忽略）"""
    return lookup_price_sqlite(str(model).strip())

def extract_current_from_model(model: str) -> int:
    """从型号中提取额定电流（支持多品牌断路器）"""
    m = re.search(r'E1C\s*(\d+)', model, re.IGNORECASE)
    if m: return int(m.group(1))
    m = re.search(r'3WL\w*N(\d{2})', model, re.IGNORECASE)
    if m:
        wl_map = {'08': 800, '10': 1000, '12': 1250, '16': 1600, '20': 2000}
        return wl_map.get(m.group(1), 0)
    m = re.search(r'MT\s*(\d+)', model, re.IGNORECASE)
    if m: return int(m.group(1))
    m = re.search(r'AN-\d+D\d+-(\d+)A', model, re.IGNORECASE)
    if m: return int(m.group(1))
    m = re.search(r'TM[DA]\s*(\d+)', model, re.IGNORECASE)
    if m: return int(m.group(1))
    m = re.search(r'NSX\s*(\d+)', model, re.IGNORECASE)
    if m: return int(m.group(1))
    m = re.search(r'3VL\w*(\d{3})', model, re.IGNORECASE)
    if m: return int(m.group(1))
    m = re.search(r'CB\w*/(\d+)A', model, re.IGNORECASE)
    if m: return int(m.group(1))
    m = re.search(r'CM1[-/](\d+)', model, re.IGNORECASE)
    if m: return int(m.group(1))
    m = re.search(r'NM1[-/](\d+)', model, re.IGNORECASE)
    if m: return int(m.group(1))
    m = re.search(r'(\d{3,4})A', model)
    if m: return int(m.group(1))
    return 0

def get_copper_spec_by_current(total_current: float) -> dict:
    """根据总电流，找到载流量 >= 总电流的最小铜排规格，超过1670A返回80×10"""
    for spec in COPPER_SPECS:
        if spec['current'] >= total_current:
            return spec
    return COPPER_SPECS[-1]  # 80×10

def get_copper_area_by_current(total_current: float) -> float:
    """返回选定铜排的截面积(cm²)，用于成本计算"""
    spec = get_copper_spec_by_current(total_current)
    return spec['area_cm2']

def get_pe_area(area_mm2: float) -> float:
    """根据GB标准计算PE保护导体最小截面积(mm²)"""
    S = area_mm2
    if S <= 16:
        return S
    elif S <= 35:
        return 16
    elif S <= 400:
        return S / 2
    elif S <= 800:
        return 200
    else:
        return S / 4

def calc_copper_busbar_cost(extra_copper_cost: float, total_current: float,
                            copper_price: float, incoming: str = '下进线',
                            width: float = 0.8) -> dict:
    """铜排成本 = 三相母线 + 零线 + 地线 + 额外铜排费用(≥250A出线)
    
    进线方式: 下进线(默认)/上进线/侧进线
    侧进线时额外增加约一台柜宽的主母排量
    """
    spec = get_copper_spec_by_current(total_current)
    L = spec['area_cm2']  # 相线截面积cm²
    S_mm2 = spec['area_mm2']  # 相线截面积mm²
    pe_area_cm2 = get_pe_area(S_mm2) / 100  # PE截面积cm²
    K = copper_price
    density = 8.9
    
    # 侧进线: 三相各增加一台柜宽
    incoming_extra = 0
    if incoming == '侧进线':
        incoming_extra = 3 * width * L * K * density / 10  # 三相各加柜宽
    
    phase_cost = 7 * L * K * density / 10 + incoming_extra
    neutral_cost = 2 * (L / 2) * K * density / 10
    ground_cost = 2 * pe_area_cm2 * K * density / 10
    copper_cost = round(phase_cost + neutral_cost + ground_cost + extra_copper_cost, 0)
    return {
        'total_cost': copper_cost,
        'phase_cost': round(phase_cost, 2),
        'neutral_cost': round(neutral_cost, 2),
        'ground_cost': round(ground_cost, 2),
        'extra_copper_cost': round(extra_copper_cost, 2),
        'copper_spec': spec,
        'copper_area_cm2': L,
        'pe_area_cm2': pe_area_cm2,
        'total_current': round(total_current, 1),
        'incoming_extra': round(incoming_extra, 2),
    }

def calc_profit(amount: float) -> float:
    """阶梯利润率（与Excel P列公式一致）"""
    if amount <= 0:
        return 0.0
    if amount < 10000:
        return amount * 0.3
    elif amount > 49999.99:
        return amount * 0.1
    else:
        return (0.29 - 0.000003125 * amount) * amount

def calc_accessory_cost(outgoing_circuits: int) -> float:
    """辅助材料费用: 价格库未找到时的fallback，提示用户手动输入"""
    return 0  # 价格库未匹配时返回0，由用户手动输入

def calc_single_cabinet(cabinet: dict, copper_price: float) -> dict:
    """计算单台柜子的所有费用"""
    components = cabinet['components']
    cable_params = load_breaker_cable_params()
    width = cabinet['width']

    # 1. 元器件费用
    total_comp_cost = 0
    total_cable_cost = 0
    total_current = 0
    high_current_copper_cost = 0  # ≥250A铜排出线费用，归入铜排

    for comp in components:
        rounded_price = round(comp['unit_price'], 0)
        amount = rounded_price * comp['qty']
        total_comp_cost += amount

        if comp['current'] >= 250 and '断路器' in comp.get('type', ''):
            # ≥250A断路器: 铜排出线，费用归入铜排
            spec = get_copper_spec_by_current(comp['current'])
            cost = 2.5 * comp['qty'] * spec['area_cm2'] * copper_price * 8.9 * 3 / 10
            high_current_copper_cost += cost
        else:
            # ≤160A: 电缆
            cable_area = cable_params.get(comp['current'], 0)
            if cable_area > 0:
                poles = get_breaker_poles(comp.get('model', ''))
                cable = cable_area * comp['qty'] * (poles * 0.7) * copper_price * 8.9 / 1000
                total_cable_cost += cable
        if '断路器' in comp.get('type', ''):
            total_current += comp['current'] * comp['qty']

    # 2. 铜排成本（根据柜内总电流独立计算）
    # 出线路数：优先使用数显仪表数量，如无仪表则使用断路器数量
    meter_count = sum(c['qty'] for c in components if '仪表' in c.get('type', '') or '数显' in c.get('type', ''))
    breaker_count_for_circuits = sum(c['qty'] for c in components if '断路器' in c.get('type', ''))
    # 出线路数始终按断路器数量计算
    outgoing_circuits = breaker_count_for_circuits
    # 分散系数：根据出线路数（参考GB/T 7251.1标准）
    if outgoing_circuits >= 10:
        derate = 0.6
    elif outgoing_circuits >= 6:
        derate = 0.7
    elif outgoing_circuits >= 4:
        derate = 0.8
    elif outgoing_circuits >= 2:
        derate = 0.9
    else:
        derate = 1.0

    reduced_current = total_current * derate

    # 仪表不单独计算铜排（铜排不接入数显仪表）
    meter_copper_cost = 0
    meter_copper_cost = 0
    
    # 刀熔开关: 断路器上方带刀熔开关时，每相增加0.4m铜排
    fuse_switch_extra = 0
    has_fuse_switch = any('刀熔' in c.get('type', '') or '刀熔' in c.get('name', '') 
                         for c in components)
    if has_fuse_switch:
        spec_for_fuse = get_copper_spec_by_current(reduced_current)
        fuse_switch_extra = 3 * 0.4 * spec_for_fuse['area_cm2'] * copper_price * 8.9 / 10

    # 带计量: 有电度表时，计量室比普通仪表室高0.2m，每相增加0.4m铜排
    meter_extra = 0
    has_meter_room = any('电度表' in c.get('name', '') for c in components)
    if has_meter_room:
        spec_for_meter = get_copper_spec_by_current(reduced_current)
        meter_extra = 3 * 0.4 * spec_for_meter['area_cm2'] * copper_price * 8.9 / 10

    copper_detail = calc_copper_busbar_cost(high_current_copper_cost + fuse_switch_extra + meter_extra, reduced_current, copper_price,
                                                  incoming=cabinet.get('incoming', '下进线'),
                                                  width=width)
    copper_detail['high_current_copper_cost'] = high_current_copper_cost
    copper_detail['meter_copper_cost'] = 0
    copper_detail['fuse_switch_extra'] = fuse_switch_extra
    copper_detail['meter_extra'] = meter_extra

    # 3. 辅助材料：从价格库查找
    # 辅助材料：用模糊匹配查找（匹配包含"N路"或"N路出线"的辅助材料记录）
    accessory_match = None
    with get_db() as conn:
        rows = conn.execute(
            "SELECT * FROM products WHERE (model LIKE '%辅助材料%' OR name LIKE '%辅助材料%') AND (model LIKE ? OR name LIKE ?)",
            (f"%{outgoing_circuits}路%", f"%{outgoing_circuits}路%")
        ).fetchall()
        if rows:
            d = dict(rows[0])
            accessory_match = {
                'name': d['name'],
                'unit_price': float(d['unit_price'] or 0),
                'retail_price': float(d['retail_price'] or 0),
                'brand': d['brand'],
            }
    if not accessory_match:
        # 退而使用旧逻辑
        accessory_name = f"辅助材料（出线柜，{outgoing_circuits}路出线）"
        accessory_match = lookup_price_by_name(accessory_name)
    if accessory_match:
        accessory_cost = accessory_match['unit_price']
        accessory_matched = True
    else:
        accessory_cost = calc_accessory_cost(outgoing_circuits)  # fallback 0，用户手动输入
        accessory_matched = False

    # 4. 箱体费用: 2200 + 400×断路器数量（只算断路器）
    breaker_count = sum(c['qty'] for c in components if '断路器' in c.get('type', ''))
    cabinet_cost = 2200 + 400 * breaker_count

    # 5. 小计（不含利润）
    subtotal = total_comp_cost + copper_detail['total_cost'] + accessory_cost + cabinet_cost

    # 6. 成套费（各项分别算利润，汇总）
    total_cost = subtotal
    comp_profit = calc_profit(total_comp_cost)
    copper_profit = calc_profit(copper_detail['total_cost'])
    accessory_profit = calc_profit(accessory_cost)
    cabinet_profit = calc_profit(cabinet_cost)
    total_profit = comp_profit + copper_profit + accessory_profit + cabinet_profit
    grand_total = total_cost + total_profit

    return {
        'name': cabinet['name'],
        'type': cabinet['type'],
        'width': width,
        'comp_cost': total_comp_cost,
        'cable_cost': total_cable_cost,
        'copper_detail': copper_detail,
        'copper_cost': copper_detail['total_cost'],
        'accessory_cost': accessory_cost,
        'accessory_label': f"辅助材料（{cabinet['name']}，{outgoing_circuits}路出线）",
        'cabinet_cost': cabinet_cost,
        'total_cost': total_cost,
        'total_profit': total_profit,
        'grand_total': grand_total,
        'comp_profit': comp_profit,
        'copper_profit': copper_profit,
        'accessory_profit': accessory_profit,
        'cabinet_profit': cabinet_profit,
        'total_current': total_current,
        'reduced_current': reduced_current,
        'outgoing_circuits': outgoing_circuits,
        'derate': derate,
        'accessory_matched': accessory_matched,
        'no_meter_warning': None,
    }

# ─── Session State 初始化 ─────────────────────────────────
def init_session_state():
    if 'project_name' not in st.session_state:
        st.session_state.project_name = ''
    if 'cabinet_list' not in st.session_state:
        st.session_state.cabinet_list = []
    if 'active_cabinet_idx' not in st.session_state:
        st.session_state.active_cabinet_idx = 0
    if 'show_calc' not in st.session_state:
        st.session_state.show_calc = False

# ─── UI ─────────────────────────────────────────────────

def main():
    init_session_state()

    st.title("⚡ 配电设备成本计算系统")
    st.caption("多柜项目模式 · 自动计算铜排成本 · 智能匹配元器件价格 · 统一利润率计算")

    # ─── 侧边栏 ───
    with st.sidebar:
        st.header("⚙️ 项目参数")
        st.subheader("📊 实时铜价（沪铜主力）")
        copper_info = get_cached_copper_price()
        if copper_info and copper_info['price'] > 0:
            copper_price = copper_info['price'] / 1000
            st.metric("卖出价", f"¥{copper_info['price']:,.0f}/吨（¥{copper_price:.1f}）",
                      help=f"数据日期: {copper_info.get('date', 'N/A')}")
            st.caption(f"📅 数据日期: {copper_info.get('date', 'N/A')}")
        else:
            copper_price = 100
            st.warning(f"⚠️ 铜价获取失败，使用默认值: ¥{copper_price}/单位（即¥{copper_price*1000:,.0f}/吨）")

        st.divider()
        st.subheader("📋 价格库搜索")
        search_term = st.text_input("搜索型号或名称", placeholder="如: XT1N160")
        if search_term:
            results = get_all_products(search_term)
            if results:
                st.write(f"找到 {len(results)} 个结果（显示前15个）")
                for r in results[:15]:
                    col1, col2 = st.columns([3, 1])
                    with col1:
                        st.text(f"{r['model']} | {r['name']}")
                    with col2:
                        st.text(f"¥{r['unit_price']:.0f}")

    # ─── 主区域 3个Tab ───
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["📋 项目配置", "⚡ 元器件管理", "📊 成本分析报告", "📖 计算公式说明", "⚙️ 价格库管理"])

    # ==================== Tab1: 项目配置 ====================
    with tab1:
        st.subheader("📋 项目配置")

        st.session_state.project_name = st.text_input(
            "项目名称", value=st.session_state.project_name,
            placeholder="输入项目名称，如：XX配电工程")

        st.divider()
        st.subheader("🗄️ 柜子管理")

        # 添加柜子
        with st.expander("➕ 添加柜子", expanded=not st.session_state.cabinet_list):
            col_a, col_b, col_c, col_d = st.columns([2, 2, 1, 2])
            with col_a:
                new_name = st.text_input("柜号名称", placeholder="如: 1#出线柜", key="new_cab_name")
            with col_b:
                new_type = st.selectbox("柜类型", CABINET_TYPES, key="new_cab_type")
            with col_c:
                new_width = st.number_input("柜宽(米)", value=0.8, min_value=0.1, max_value=2.0,
                                            step=0.1, key="new_cab_width")
            with col_d:
                new_incoming = st.selectbox("进线方式", ["上进线", "侧进线", "下进线"],
                                            index=2, key="new_cab_incoming")  # 默认下进线

            if st.button("✅ 添加柜子", type="primary", use_container_width=True):
                if new_name:
                    st.session_state.cabinet_list.append({
                        'name': new_name,
                        'type': new_type,
                        'width': new_width,
                        'incoming': new_incoming,
                        'components': [],
                    })
                    st.session_state.active_cabinet_idx = len(st.session_state.cabinet_list) - 1
                    st.success(f"✅ 已添加: {new_name}")
                    st.rerun()
                else:
                    st.warning("请填写柜号名称")

        # 显示柜子列表
        if st.session_state.cabinet_list:
            st.divider()
            st.subheader("📋 柜子列表")

            for i, cab in enumerate(st.session_state.cabinet_list):
                is_active = (i == st.session_state.active_cabinet_idx)
                border = "🔴" if is_active else "⚪"

                col1, col2, col3, col4, col5, col6, col7 = st.columns([0.5, 2, 1.5, 1, 1, 0.8, 0.8])
                with col1:
                    st.markdown(border)
                with col2:
                    if st.button(f"{'▶ ' if is_active else ''}{cab['name']}", key=f"sel_cab_{i}"):
                        st.session_state.active_cabinet_idx = i
                        st.rerun()
                with col3:
                    st.text(cab['type'])
                with col4:
                    st.text(f"{cab['width']}m")
                with col5:
                    st.text(f"{len(cab['components'])}个元器件")
                with col6:
                    if st.button("📋", key=f"copy_cab_{i}"):
                        import copy
                        new_cabinet = copy.deepcopy(st.session_state.cabinet_list[i])
                        new_cabinet['name'] = st.session_state.cabinet_list[i]['name'] + '(副本)'
                        st.session_state.cabinet_list.append(new_cabinet)
                        st.session_state.active_cabinet_idx = len(st.session_state.cabinet_list) - 1
                        st.rerun()
                with col7:
                    if st.button("🗑️", key=f"del_cab_{i}"):
                        st.session_state.cabinet_list.pop(i)
                        if st.session_state.active_cabinet_idx >= len(st.session_state.cabinet_list):
                            st.session_state.active_cabinet_idx = max(0, len(st.session_state.cabinet_list) - 1)
                        st.rerun()

            # 编辑当前选中柜子
            st.divider()
            idx = st.session_state.active_cabinet_idx
            if idx < len(st.session_state.cabinet_list):
                cab = st.session_state.cabinet_list[idx]
                st.subheader(f"✏️ 编辑柜子: {cab['name']}")

                col_e1, col_e2, col_e3 = st.columns([2, 2, 1])
                with col_e1:
                    new_name = st.text_input("柜号名称", value=cab['name'], key="edit_cab_name")
                with col_e2:
                    new_type = st.selectbox("柜类型", CABINET_TYPES,
                                            index=CABINET_TYPES.index(cab['type']) if cab['type'] in CABINET_TYPES else 0,
                                            key="edit_cab_type")
                with col_e3:
                    new_width = st.number_input("柜宽(米)", value=cab['width'],
                                                min_value=0.1, max_value=2.0, step=0.1, key="edit_cab_width")

                if st.button("💾 保存修改", use_container_width=True):
                    st.session_state.cabinet_list[idx]['name'] = new_name
                    st.session_state.cabinet_list[idx]['type'] = new_type
                    st.session_state.cabinet_list[idx]['width'] = new_width
                    st.success("✅ 已保存")
                    st.rerun()
        else:
            st.info("请先添加柜子")

    # ==================== Tab2: 元器件管理 ====================
    with tab2:
        if not st.session_state.cabinet_list:
            st.info("👈 请先在'项目配置'中添加柜子")
        else:
            idx = st.session_state.active_cabinet_idx
            if idx >= len(st.session_state.cabinet_list):
                idx = 0
                st.session_state.active_cabinet_idx = 0

            cab = st.session_state.cabinet_list[idx]

            # 柜子信息头部
            col_h1, col_h2, col_h3 = st.columns([2, 1.5, 1])
            col_h1.markdown(f"### 🔧 {cab['name']}")
            col_h2.markdown(f"**类型:** {cab['type']}")
            col_h3.markdown(f"**柜宽:** {cab['width']}m")

            # 快速切换柜子
            cab_names = [f"{i+1}. {c['name']} ({c['type']})" for i, c in enumerate(st.session_state.cabinet_list)]
            selected = st.selectbox("快速切换柜子", range(len(cab_names)),
                                    format_func=lambda x: cab_names[x], index=idx,
                                    key="cab_switch")
            if selected != idx:
                st.session_state.active_cabinet_idx = selected
                st.rerun()

            st.divider()

            # 添加元器件表单
            with st.expander("➕ 添加元器件", expanded=True):
                # 搜索建议（在widget之前处理选择逻辑）
                selected_model = st.session_state.get('_selected_model', '')
                if selected_model:
                    st.session_state.pop('_selected_model', None)
                    st.session_state.pop('model_input', None)
                    st.session_state.pop('name_input', None)
                    st.session_state.pop('current_input', None)

                col1, col2, col3 = st.columns([2, 1, 1])

                with col1:
                    model_input = st.text_input("型号", key="model_input",
                                                placeholder="输入元器件型号",
                                                value=selected_model if selected_model else '')
                    if model_input:
                        suggestions = get_all_products(model_input)[:5]
                        if suggestions:
                            for s in suggestions:
                                if st.button(f"选择: {s['model']}", key=f"sel_{s['model']}"):
                                    st.session_state._selected_model = s['model']
                                    st.session_state._selected_name = s['name']
                                    st.rerun()

                with col2:
                    qty_input = st.number_input("数量", min_value=0, value=1, key="qty_input")

                with col3:
                    # placeholder for spacing
                    st.empty()

                col_a, col_b = st.columns([1, 1])
                with col_a:
                    auto_current = extract_current_from_model(model_input) if model_input else 0
                    if auto_current > 0:
                        st.session_state.current_input = auto_current
                    current_input = st.number_input("额定电流 (A)", min_value=0,
                                                    key="current_input")
                    if auto_current > 0:
                        st.caption(f"💡 自动识别: {auto_current}A（可手动修改）")

                with col_b:
                    # 从价格库匹配类型
                    auto_type = ''
                    if model_input:
                        match = lookup_price(model_input)
                        if match and match['name']:
                            auto_type = match['name']
                    
                    use_custom_type = st.checkbox("自定义类型", key="custom_type_cb", value=False)
                    if use_custom_type:
                        custom_type = st.text_input("类型", key="custom_type_input", placeholder="输入自定义类型")
                        preset_type = ''
                    else:
                        type_options = ["塑壳断路器", "框架断路器", "电流互感器", "数显仪表", "其他"]
                        if auto_type and auto_type in type_options:
                            default_type = auto_type
                        else:
                            # 无法识别时不默认任何类型，让用户手动选择
                            type_options = ["（请选择类型）", "塑壳断路器", "框架断路器", "电流互感器", "数显仪表", "其他"]
                            default_type = type_options[0]
                        # 动态key确保型号变化时selectbox重新初始化
                        type_key = f"type_input_{model_input or '_empty'}"
                        preset_type = st.selectbox("类型", type_options, index=type_options.index(default_type), key=type_key)
                        if auto_type and auto_type not in type_options:
                            st.caption(f"💡 价格库类型: {auto_type}，请选择最接近的类型或自定义")
                        elif auto_type and auto_type in type_options:
                            st.caption(f"💡 已自动匹配类型: {auto_type}")
                        elif not auto_type:
                            st.warning("⚠️ 未在价格库中找到该型号，请手动选择类型！")
                        custom_type = ''

                if st.button("✅ 添加到清单", type="primary", use_container_width=True):
                    if model_input:
                        comp_type = custom_type if use_custom_type else preset_type
                        if comp_type.startswith('（请选择'):
                            st.error("❌ 请先选择元器件类型！未识别的型号需要手动选择类型，否则成本计算会出错。")
                        else:
                            match = lookup_price(model_input)
                            current = current_input if current_input > 0 else extract_current_from_model(model_input)
                            component = {
                                'model': model_input,
                                'name': comp_type,
                                'qty': qty_input,
                                'current': current,
                                'type': comp_type,
                                'unit_price': match['unit_price'] if match else 0,
                                'retail_price': match['retail_price'] if match else 0,
                                'brand': match['brand'] if match else '未找到',
                                'matched': match is not None,
                            }
                            # 合并同型号
                            existing = None
                            for c in st.session_state.cabinet_list[idx]['components']:
                                if c['model'] == model_input:
                                    existing = c
                                    break
                            if existing:
                                existing['qty'] += qty_input
                                st.success(f"✅ 已合并 {cab['name']}: {model_input} 数量→{existing['qty']}")
                            else:
                                st.session_state.cabinet_list[idx]['components'].append(component)
                                st.success(f"✅ 已添加到 {cab['name']}: {model_input} × {qty_input}")
                            st.rerun()
                    else:
                        st.warning("请填写型号和数量")

            # 价格库浏览
            with st.expander("📚 价格库浏览", expanded=False):
                price_search = st.text_input("搜索型号或名称", key="price_db_search", placeholder="输入关键词...")
                products = get_all_products(price_search)
                price_df = pd.DataFrame(products)
                if not price_df.empty:
                    display_cols = [c for c in ['id', 'model', 'name', 'unit_price', 'retail_price', 'brand'] if c in price_df.columns]
                    rename = {'model': '型号', 'name': '名称', 'unit_price': '单价', 'retail_price': '面价', 'brand': '品牌'}
                    price_df = price_df[display_cols].rename(columns=rename)
                    st.dataframe(price_df.head(20), use_container_width=True, hide_index=True)

            # 批量导入
            with st.expander(f"📦 批量导入 → {cab['name']}", expanded=False):
                st.caption(f"导入到柜子: **{cab['name']}** ({cab['type']}, {cab['width']}m)")
                import_mode = st.radio("导入方式", ["📄 Excel文件", "📝 文本粘贴"], horizontal=True)

                preview_data = []

                if import_mode == "📄 Excel文件":
                    uploaded = st.file_uploader("上传文件", type=["xlsx", "xls", "csv"], key="batch_excel")
                    if uploaded:
                        try:
                            if uploaded.name.endswith('.csv'):
                                # CSV文件处理
                                import csv
                                import io
                                content = uploaded.read().decode('utf-8-sig')
                                reader = csv.reader(io.StringIO(content))
                                all_rows = list(reader)
                                header_text = all_rows[0] if all_rows else []

                                col_model = None
                                col_qty = None
                                col_name = None
                                has_header = False

                                for ci, h in enumerate(header_text):
                                    if '型号' in str(h):
                                        col_model = ci; has_header = True
                                    if '数量' in str(h):
                                        col_qty = ci; has_header = True
                                    if '名称' in str(h):
                                        col_name = ci; has_header = True

                                if not has_header:
                                    col_name = 1; col_model = 2; col_qty = 3
                                    data_rows = all_rows[:15]
                                else:
                                    data_rows = all_rows[1:16]

                                for row in data_rows:
                                    model = str(row[col_model] or '').strip() if col_model is not None and col_model < len(row) else ''
                                    qty = row[col_qty] if col_qty is not None and col_qty < len(row) else 0
                                    name = str(row[col_name] or '').strip() if col_name is not None and col_name < len(row) else ''
                                    if model:
                                        try:
                                            qty = int(float(qty))
                                        except (ValueError, TypeError):
                                            qty = 1
                                        match = lookup_price(model)
                                        preview_data.append({
                                            'name': name or (match['name'] if match else ''),
                                            'model': model,
                                            'qty': qty,
                                            'unit_price': match['unit_price'] if match else 0,
                                            'matched': match is not None,
                                        })
                            else:
                                # Excel文件处理
                                wb = openpyxl.load_workbook(uploaded, read_only=True, data_only=True)
                                ws = wb.active
                                rows = list(ws.iter_rows(max_row=16, values_only=False))
                                wb.close()

                                # 识别列：找含"型号"的列头作为型号列(C)，含"数量"的列头作为数量列(D)，含"名称"的列头作为名称列(B)
                                header = rows[0] if rows else []
                                header_text = [str(c.value or '') for c in header]

                                col_model = None
                                col_qty = None
                                col_name = None
                                has_header = False

                                for ci, h in enumerate(header_text):
                                    if '型号' in h:
                                        col_model = ci; has_header = True
                                    if '数量' in h:
                                        col_qty = ci; has_header = True
                                    if '名称' in h:
                                        col_name = ci; has_header = True

                                if not has_header:
                                    col_name = 1; col_model = 2; col_qty = 3  # B=1, C=2, D=3
                                    data_rows = rows[:15]
                                else:
                                    data_rows = rows[1:16]

                                for row in data_rows:
                                    vals = [c.value for c in row]
                                    model = str(vals[col_model] or '').strip() if col_model is not None and col_model < len(vals) else ''
                                    qty = vals[col_qty] if col_qty is not None and col_qty < len(vals) else 0
                                    name = str(vals[col_name] or '').strip() if col_name is not None and col_name < len(vals) else ''
                                    if model:
                                        try:
                                            qty = int(float(qty))
                                        except (ValueError, TypeError):
                                            qty = 1
                                        match = lookup_price(model)
                                        preview_data.append({
                                            'name': name or (match['name'] if match else ''),
                                            'model': model,
                                            'qty': qty,
                                            'unit_price': match['unit_price'] if match else 0,
                                            'matched': match is not None,
                                        })
                        except Exception as e:
                            st.error(f"读取文件失败: {e}")

                else:  # 文本粘贴
                    text = st.text_area("粘贴元器件清单（每行一个：型号 数量）", height=68,
                                        placeholder="XT1N160 TMD 100 3P FF 1\nMC7200 6", key="batch_text")
                    if text.strip():
                        for line in text.strip().split('\n'):
                            # 从行末尾提取数量（支持空格/Tab分隔）
                            line = line.strip()
                            if not line:
                                continue
                            # 尝试从末尾提取数字作为数量
                            qty_match = re.search(r'[\t ]+(\d+)\s*$', line)
                            if qty_match:
                                qty = int(qty_match.group(1))
                                model = line[:qty_match.start()].strip()
                                name = ''
                            else:
                                # 尝试Tab/多空格分隔
                                parts = re.split(r'[\t]+|[ \t]{2,}', line)
                                parts = [p.strip() for p in parts if p.strip()]
                                if len(parts) >= 3:
                                    name, model, qty_str = parts[0], parts[1], parts[2]
                                elif len(parts) == 2:
                                    name, model, qty_str = '', parts[0], parts[1]
                                else:
                                    continue
                                try:
                                    qty = int(float(qty_str))
                                except (ValueError, TypeError):
                                    continue
                            match = lookup_price(model)
                            preview_data.append({
                                'name': name or (match['name'] if match else ''),
                                'model': model,
                                'qty': qty,
                                'unit_price': match['unit_price'] if match else 0,
                                'matched': match is not None,
                            })

                if preview_data:
                    st.subheader(f"📋 预览导入 ({len(preview_data)} 个元器件)")
                    pdf = []
                    for i, item in enumerate(preview_data):
                        _current = extract_current_from_model(item['model'])
                        _match = lookup_price(item['model'])
                        _type = _match.get('type', '') if _match else ''
                        _name = item['name'] or (_match['name'] if _match else '未知')
                        pdf.append({
                            '序号': i + 1,
                            '名称': _name,
                            '型号': item['model'],
                            '电流(A)': _current,
                            '类型': _type or '⚠️未识别',
                            '数量': item['qty'],
                            '匹配单价': f"¥{item['unit_price']:.0f}" if item['matched'] else '-',
                            '状态': '✅已匹配' if item['matched'] else '⚠️未找到',
                        })
                    st.dataframe(pd.DataFrame(pdf), use_container_width=True, hide_index=True)

                    if st.button("✅ 确认导入", type="primary", use_container_width=True, key="batch_confirm"):
                        for item in preview_data:
                            current = extract_current_from_model(item['model'])
                            match = lookup_price(item['model'])
                            # 从价格库获取类型，未匹配则置空
                            comp_type = ''
                            if match and match.get('type'):
                                comp_type = match['type']
                            component = {
                                'model': item['model'],
                                'name': item['name'] or (match['name'] if match else '未知'),
                                'qty': item['qty'],
                                'current': current,
                                'type': comp_type,
                                'unit_price': item['unit_price'],
                                'retail_price': match['retail_price'] if match else 0,
                                'brand': match['brand'] if match else '未找到',
                                'matched': item['matched'],
                            }
                            st.session_state.cabinet_list[idx]['components'].append(component)
                        st.success(f"✅ 已导入 {len(preview_data)} 个元器件到 {cab['name']}")
                        st.rerun()

            # 显示元器件清单
            components = cab['components']
            if components:
                st.divider()
                st.subheader(f"📋 {cab['name']} - 元器件清单")

                display_data = []
                for i, comp in enumerate(components):
                    amount = round(comp['unit_price'], 0) * comp['qty']
                    status = '✅' if comp['matched'] else '⚠️未匹配'
                    if not comp.get('type') or comp['type'].startswith('（'):
                        status = '❌类型未设置'
                    display_data.append({
                        '序号': i + 1,
                        '名称': comp['name'],
                        '型号': comp['model'],
                        '电流(A)': comp['current'],
                        '数量': comp['qty'],
                        '单价': f"¥{comp['unit_price']:.0f}",
                        '金额': f"¥{amount:,.0f}",
                        '品牌': comp['brand'],
                        '状态': status,
                    })

                df_display = pd.DataFrame(display_data)
                st.dataframe(df_display, use_container_width=True, hide_index=True)

                # 类型未设置的警告
                unset = [c for c in components if not c.get('type') or c['type'].startswith('（')]
                if unset:
                    st.error(f"⚠️ 有 {len(unset)} 个元器件类型未设置，将无法正确计算成本！请返回编辑。")

                # 删除按钮（输入序号删除）
                if components:
                    del_num = st.number_input("删除第几行", min_value=1, max_value=len(components), value=1, step=1, key=f"del_row_{idx}", format="%d")
                    c1, c2 = st.columns(2)
                    with c1:
                        if st.button(f"🗑️ 删除第{del_num}行: {components[del_num-1]['model']}", key=f"del_btn_{idx}"):
                            st.session_state.cabinet_list[idx]['components'].pop(del_num - 1)
                            st.rerun()
                    with c2:
                        if st.button("🗑️ 清空清单", key=f"clear_{idx}"):
                            st.session_state.cabinet_list[idx]['components'] = []
                            st.rerun()

                col_clear, col_calc = st.columns([1, 2])
                with col_clear:
                    if st.button("🗑️ 清空清单"):
                        st.session_state.cabinet_list[idx]['components'] = []
                        st.rerun()

                with col_calc:
                    if st.button("📊 计算成本分析报告", type="primary", use_container_width=True):
                        st.session_state.show_calc = True
                        st.rerun()

    # ==================== Tab3: 成本分析报告 ====================
    with tab3:
        if not st.session_state.cabinet_list:
            st.info("👈 请先添加柜子和元器件")
        elif not any(cab['components'] for cab in st.session_state.cabinet_list):
            st.info("👈 请先在'元器件管理'中添加元器件并点击'计算成本分析报告'")
        elif st.session_state.get('show_calc'):
            run_project_report(st.session_state.cabinet_list, copper_price)
        else:
            st.info("👈 请在'元器件管理'中点击'计算成本分析报告'")

    # ==================== Tab4: 计算公式说明 ====================
    with tab4:
        st.header("📖 计算公式说明")
        st.caption("配电设备成本计算系统所使用的全部公式和参数")

        # 1. 铜排载流量对照表
        with st.expander("📏 1. 铜排载流量对照表", expanded=False):
            st.markdown("""
            根据降容后总电流，从下表选取满足载流量要求的最小铜排规格。
            """)
            copper_table = pd.DataFrame(COPPER_SPECS)[['spec', 'area_mm2', 'area_cm2', 'current']]
            copper_table.columns = ['规格', '截面积(mm²)', '截面积(cm²)', '载流量(A)']
            st.dataframe(copper_table, use_container_width=True, hide_index=True)

        # 2. 电缆费用公式
        with st.expander("🔌 2. 电缆费用公式", expanded=False):
            st.markdown("#### 塑壳断路器电缆（≤160A）")
            st.code("""电缆费(元) = 截面积(mm²) × 数量(个) × 长度(m) × 铜价(元/kg) × 8.9(g/cm³) / 1000

长度 = 极数 × 0.7m
  1P → 0.7m
  2P → 1.4m
  3P → 2.1m
  4P → 2.8m""", language="text")
            st.markdown("- 按标准电缆截面积计算，每P 0.7m")
            st.markdown("- 160A及以下用电缆，250A及以上用铜排")

            cable_params = load_breaker_cable_params()
            cable_table = pd.DataFrame([
                {'电流(A)': k, '截面积(mm²)': v} for k, v in sorted(cable_params.items()) if v > 0 and k <= 160
            ])
            st.dataframe(cable_table, use_container_width=True, hide_index=True)

            st.markdown("#### 断路器铜排出线（≥250A）")
            st.code("""出线费(元) = 2.5(m) × 数量(个) × 铜排截面积(cm²) × 铜价(元/kg) × 8.9(g/cm³) × 3(三相) / 10""", language="text")
            st.markdown("- 250A及以上（塑壳/框架）均用铜排出线，出线长度2.5m")
            st.markdown("- 费用归入铜排成本")

        # 3. 铜排成本公式
        with st.expander("🟫 3. 铜排成本公式", expanded=False):
            st.code("""铜排费(元) = ROUND(
  三相母线: 7(m) × 截面积(cm²) × 铜价(元/kg) × 8.9(g/cm³) / 10
  + 侧进线增加: 3 × 柜宽(m) × 截面积(cm²) × 铜价(元/kg) × 8.9(g/cm³) / 10  // 仅侧进线
  零线N:   2(m) × (截面积/2)(cm²) × 铜价(元/kg) × 8.9(g/cm³) / 10
  地线PE:  2(m) × PE截面积(cm²) × 铜价(元/kg) × 8.9(g/cm³) / 10  // PE按国标分段
  + ≥250A出线铜排费(元)
  + 刀熔开关增加: 3 × 0.4(m) × 截面积(cm²) × 铜价(元/kg) × 8.9(g/cm³) / 10  // 有刀熔开关时
  + 带计量增加: 3 × 0.4(m) × 截面积(cm²) × 铜价(元/kg) × 8.9(g/cm³) / 10  // 有电度表时
, 0)""", language="text")
            st.markdown("""
            - **7** = 水平母线总长度(m)，经验估算值
            - **2** = 零线/地线长度(m)
            - **8.9** = 铜密度(g/cm³)，物理常数
            - 铜排截面积根据降容后总电流(A)从载流量对照表选取
            - **进线方式**: 下进线(默认)不增加; 侧进线三相各加一台柜宽
            - **刀熔开关**: 有刀熔开关时每相增加0.4m
            - **带计量**: 有电度表时每相增加0.4m（计量室比仪表室高0.2m）
            - **PE截面积国标**: S≤16→S, 16<S≤35→16, 35<S≤400→S/2, 400<S≤800→200, S>800→S/4
            """)

        # 4. 分散系数
        with st.expander("📉 4. 分散系数", expanded=False):
            st.code("""降容后电流(A) = 总电流(A) × 分散系数

分散系数（GB/T 7251标准）:
  1路:     1.0
  2-3路:   0.9
  4-5路:   0.8
  6-9路:   0.7
  ≥10路:   0.6""", language="text")
            st.markdown("- 出线路数 = 断路器总数量（塑壳+框架）")

        # 5. 辅助材料费用
        with st.expander("🔧 5. 辅助材料费用", expanded=False):
            st.code("""辅材费(元) = 从价格库查找
名称格式: 辅助材料（出线柜，N路出线）""", language="text")
            st.markdown("- 未找到时默认0元，提示用户手动输入")

        # 6. 箱体费用
        with st.expander("📦 6. 箱体费用", expanded=False):
            st.code("""箱体费(元) = 2200(元) + 400(元/个) × 断路器总数量(个)""", language="text")
            st.markdown("- 2200元为基础柜体费用，每个断路器(抽屉)加400元")

        # 7. 成套费计算
        with st.expander("💰 7. 成套费计算（阶梯利润率）", expanded=False):
            st.code("""成套费(元) = 各项费用之和
  = 元器件(元) + 铜排(元) + 辅材(元) + 箱体(元)""", language="text")
            st.code("""利润率阶梯（按各项金额分别计算）:
  金额(元) < 10,000    → 金额 × 30%
  金额(元) 10,000~50,000  → (0.29 - 0.000003125 × 金额) × 金额
  金额(元) > 50,000    → 金额 × 10%""", language="text")
            st.markdown("- 元器件、铜排、辅材、箱体各自按上述阶梯计算利润后汇总")
            st.markdown("- 整体利润率 = 成套费(元) ÷ 总成本(元) × 100%")

        # 8. 最终报价
        with st.expander("🏷️ 8. 最终报价", expanded=False):
            st.code("""含税报价(元) = (总成本(元) + 成套费(元)) × 1.13""", language="text")

    # ==================== Tab5: 价格库管理 ====================
    with tab5:
        st.subheader("⚙️ 价格库管理")
        total_count = len(get_all_products())
        st.caption(f"共 {total_count} 个产品")

        # 搜索和筛选
        col_search, col_brand = st.columns([3, 1])
        with col_search:
            mgmt_search = st.text_input("搜索型号/名称", key="mgmt_search", placeholder="输入关键词...")
        with col_brand:
            brands = get_all_brands()
            brand_filter = st.selectbox("品牌筛选", ["全部"] + brands, key="mgmt_brand")

        # 构建查询
        products = get_all_products(mgmt_search)
        if brand_filter != "全部":
            products = [p for p in products if p['brand'] == brand_filter]

        if products:
            # 可编辑表格
            edit_df = pd.DataFrame(products)[['id', 'model', 'name', 'unit_price', 'retail_price', 'brand']]
            edit_df.columns = ['ID', '型号', '名称', '单价', '面价', '品牌']

            # 分页控件
            page_size = st.selectbox("每页显示", [50, 100, 200, 500, 9999], index=0,
                                     key="mgmt_page_size", format_func=lambda x: "全部" if x >= len(products) else str(x))
            page_size = min(page_size, len(products))

            edited = st.data_editor(
                edit_df.head(page_size),
                use_container_width=True,
                hide_index=True,
                disabled=['ID', '型号'],
                key="price_editor",
            )
            if page_size < len(products):
                st.caption(f"显示前 {page_size} 条，共 {len(products)} 条")

            # 检测编辑并保存
            if st.button("💾 保存编辑", type="primary"):
                changes = 0
                for _, row in edited.iterrows():
                    pid = int(row['ID'])
                    orig = next((p for p in products if p['id'] == pid), None)
                    if orig and (orig['name'] != row['名称'] or orig['unit_price'] != row['单价']
                                or orig['retail_price'] != row['面价'] or orig['brand'] != row['品牌']):
                        update_product(pid, row['名称'], row['单价'], row['面价'], row['品牌'])
                        changes += 1
                if changes:
                    st.success(f"✅ 已保存 {changes} 条修改")
                    st.rerun()
                else:
                    st.info("没有修改")

            # 删除
            del_col1, del_col2 = st.columns([3, 1])
            with del_col2:
                del_model = st.text_input("删除型号", key="del_model_input", placeholder="输入要删除的型号")
                if del_model:
                    p = lookup_price_sqlite(del_model)
                    if p:
                        if st.button(f"🗑️ 删除 {del_model}", type="secondary"):
                            # find id
                            with get_db() as conn:
                                row = conn.execute("SELECT id FROM products WHERE model=?", (del_model,)).fetchone()
                                if row:
                                    delete_product(row['id'])
                            st.success(f"✅ 已删除 {del_model}")
                            st.rerun()
                    else:
                        st.warning("型号不存在")
        else:
            st.info("没有找到匹配的产品")

        st.divider()

        # 新增产品
        with st.expander("➕ 新增产品", expanded=False):
            nc1, nc2, nc3 = st.columns([2, 2, 1])
            with nc1:
                new_model = st.text_input("型号", key="new_prod_model")
            with nc2:
                new_name = st.text_input("名称", key="new_prod_name")
            with nc3:
                new_brand = st.text_input("品牌", key="new_prod_brand")
            nc4, nc5 = st.columns([1, 1])
            with nc4:
                new_unit = st.number_input("单价", min_value=0.0, key="new_prod_unit", step=1.0)
            with nc5:
                new_retail = st.number_input("面价", min_value=0.0, key="new_prod_retail", step=1.0)

            if st.button("✅ 添加产品", key="add_product_btn"):
                if new_model:
                    existing = lookup_price_sqlite(new_model)
                    if existing:
                        st.warning(f"型号 {new_model} 已存在，请使用编辑功能修改")
                    else:
                        insert_product(new_model, new_name, new_unit, new_retail, new_brand)
                        st.success(f"✅ 已添加 {new_model}")
                        st.rerun()
                else:
                    st.warning("请填写型号")

        # 批量导入
        with st.expander("📦 批量导入（Excel/CSV）", expanded=False):
            import_file = st.file_uploader("上传文件", type=["xlsx", "xls", "csv"], key="price_import_file")
            if import_file:
                try:
                    import_df = pd.read_excel(import_file) if not import_file.name.endswith('.csv') else pd.read_csv(import_file, encoding='utf-8-sig')
                    # 尝试映射列
                    col_map = {}
                    for c in import_df.columns:
                        cl = str(c).strip()
                        if '型号' in cl: col_map['model'] = c
                        elif '名称' in cl: col_map['name'] = c
                        elif '单价' in cl: col_map['unit_price'] = c
                        elif '面价' in cl: col_map['retail_price'] = c
                        elif '品牌' in cl or '产地' in cl: col_map['brand'] = c

                    if 'model' not in col_map:
                        # fallback: use first 3 columns
                        cols = list(import_df.columns)
                        col_map = {'model': cols[0], 'name': cols[1] if len(cols)>1 else '', 
                                  'unit_price': cols[2] if len(cols)>2 else ''}

                    preview_items = []
                    for _, row in import_df.iterrows():
                        model = str(row.get(col_map.get('model',''), '')).strip()
                        if not model:
                            continue
                        preview_items.append({
                            'model': model,
                            'name': str(row.get(col_map.get('name',''), '')).strip(),
                            'unit_price': float(row.get(col_map.get('unit_price',0), 0) or 0),
                            'retail_price': float(row.get(col_map.get('retail_price',0), 0) or 0),
                            'brand': str(row.get(col_map.get('brand',''), '')).strip(),
                        })

                    if preview_items:
                        st.write(f"预览 {len(preview_items)} 个产品（前20个）：")
                        st.dataframe(pd.DataFrame(preview_items[:20]), use_container_width=True, hide_index=True)
                        if st.button("✅ 确认导入", type="primary", key="confirm_import"):
                            bulk_import_products(preview_items)
                            st.success(f"✅ 已导入 {len(preview_items)} 个产品")
                            st.rerun()
                    else:
                        st.warning("未识别到有效数据")
                except Exception as e:
                    st.error(f"读取文件失败: {e}")

        # 操作日志
        with st.expander("📜 操作日志", expanded=False):
            logs = get_price_logs(50)
            if logs:
                log_data = []
                for log in logs:
                    log_data.append({
                        '时间': log['created_at'],
                        '操作': log['action'],
                        '型号': log['model'],
                        '变更内容': log.get('new_value', '')[:80] if log['action'] == 'insert' 
                                   else f"{log.get('old_value','')[:40]}→{log.get('new_value','')[:40]}",
                    })
                st.dataframe(pd.DataFrame(log_data), use_container_width=True, hide_index=True)
            else:
                st.info("暂无操作日志")

# ─── 计算日志 ───────────────────────────────────────────────
LOG_FILE = Path(__file__).parent / "data" / "calc_logs.jsonl"
MAX_LOGS = 1500

def save_calc_log(copper_price, cabinets, results):
    """保存计算日志到JSONL文件，最多保留1500条"""
    LOG_FILE.parent.mkdir(exist_ok=True)
    # (removed, no longer used)
    # (no longer used, remove)
    entry = {
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "copper_price": copper_price,
        "cabinets": []
    }
    for cab, result in zip(cabinets, results):
        entry["cabinets"].append({
            "name": cab.get("name", ""),
            "type": cab.get("type", ""),
            "width": cab.get("width", 0),
            "components": cab.get("components", []),
            "result": result
        })

    # 读取现有日志
    logs = []
    if LOG_FILE.exists():
        with open(LOG_FILE, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line:
                    try:
                        logs.append(json.loads(line))
                    except:
                        pass

    logs.append(entry)
    if len(logs) > MAX_LOGS:
        logs = logs[-MAX_LOGS:]

    with open(LOG_FILE, 'w', encoding='utf-8') as f:
        for log in logs:
            f.write(json.dumps(log, ensure_ascii=False) + '\n')

    # 生成报价摘要并发送Telegram通知
    if TELEGRAM_NOTIFY:
        try:
            summary_lines = []
            total_all = 0
            for cab, result in zip(cabinets, results):
                summary_lines.append(f"📦 {cab.get('name','')} ({cab.get('type','')} {cab.get('width',0)}m)")
                summary_lines.append(f"   元器件: ¥{result['comp_cost']:.0f}  铜排: ¥{result['copper_cost']:.0f}  电缆: ¥{result['cable_cost']:.0f}")
                summary_lines.append(f"   辅材: ¥{result['accessory_cost']:.0f}  箱体: ¥{result['cabinet_cost']:.0f}  成套费: ¥{result['total_profit']:.0f}")
                summary_lines.append(f"   💰 含税: ¥{result['grand_total']:.0f}")
                total_all += result['grand_total']
            msg = f"🧾 配电设备成本报价\n⏰ {entry['timestamp']}\n🔩 铜价: ¥{copper_price:.2f}/kg\n\n" + "\n".join(summary_lines) + f"\n\n📊 项目合计含税: ¥{total_all:.0f}"
            
            # 元器件清单（第二条消息，避免太长）
            comp_lines = []
            for cab in cabinets:
                comp_lines.append(f"【{cab.get('name','')}】")
                for c in cab.get('components', []):
                    if c.get('qty', 0) > 0:
                        comp_lines.append(f"  {c.get('model','')} ×{c['qty']}  ¥{c.get('unit_price',0):.0f}  {c.get('brand','')}")
            comp_msg = "📋 元器件清单\n\n" + "\n".join(comp_lines)
            
            url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
            requests.post(url, json={"chat_id": TELEGRAM_CHAT_ID, "text": msg}, timeout=10)
            requests.post(url, json={"chat_id": TELEGRAM_CHAT_ID, "text": comp_msg}, timeout=10)
        except Exception as e:
            pass  # 发送失败不影响计算

    # 写入通知文件（心跳备用）
    notify_file = Path(__file__).parent / "data" / ".calc_notify"
    notify_file.write_text(json.dumps({
        "timestamp": entry["timestamp"],
        "copper_price": copper_price,
        "cabinets": [{"name": c["name"], "result": r} for c, r in zip(cabinets, results)]
    }, ensure_ascii=False))


def run_project_report(cabinet_list: list, copper_price: float):
    """生成项目成本分析报告"""
    project_name = st.session_state.project_name or "未命名项目"
    st.header(f"📊 成本分析报告 - {project_name}")

    # 计算每台柜子
    cabinet_results = []
    for cab in cabinet_list:
        if cab['components']:
            result = calc_single_cabinet(cab, copper_price)
            # 检查辅材是否匹配成功
            if not result.get('accessory_matched', False):
                st.warning(f"⚠️ 未在价格库中找到 {result['outgoing_circuits']}路出线辅助材料，请手动输入或更新价格库")
                result['accessory_cost'] = st.number_input(
                    f"🔧 {result['name']} 辅助材料价格（元）",
                    min_value=0.0, value=float(result['accessory_cost']),
                    step=100.0, format="%.0f", key=f"accessory_{result['name']}")
                # 重新计算
                result['total_cost'] = result['comp_cost'] + result['copper_cost'] + result['accessory_cost'] + result['cabinet_cost']
                result['total_profit'] = calc_profit(result['comp_cost']) + calc_profit(result['copper_cost']) + calc_profit(result['accessory_cost']) + calc_profit(result['cabinet_cost'])
                result['grand_total'] = result['total_cost'] + result['total_profit']
            cabinet_results.append(result)

    if cabinet_results:
        # 保存计算日志
        save_calc_log(copper_price,
                      [cab for cab in cabinet_list if cab['components']],
                      cabinet_results)

    if not cabinet_results:
        st.warning("没有包含元器件的柜子")
        return

    # ─── 每台柜子明细 ───
    st.subheader("🗄️ 各柜明细")

    project_total_cost = 0
    project_total_profit = 0

    for result in cabinet_results:
        # 找到对应的cabinet
        cab = next(c for c in cabinet_list if c['name'] == result['name'])
        components = cab['components']

        with st.expander(f"{'🔴 ' if result == cabinet_results[0] else ''}{result['name']} ({result['type']}, {result['width']}m)", expanded=(result == cabinet_results[0])):
            if result.get('no_meter_warning'):
                st.warning(result['no_meter_warning'], icon="⚠️")

            # 构建表格数据
            table_data = []
            idx = 1
            for comp in components:
                if comp['qty'] == 0:  # 报告中隐藏数量为0的
                    continue
                amount = round(comp['unit_price'], 0) * comp['qty']
                table_data.append({
                    '序号': idx,
                    '名称': comp['name'],
                    '型号': comp['model'],
                    '数量': comp['qty'],
                    '单价(元)': f"¥{comp['unit_price']:.0f}",
                    '金额(元)': f"¥{amount:,.0f}",
                    '品牌': comp['brand'],
                })
                idx += 1

            # 铜排行
            cd = result['copper_detail']
            copper_spec_str = f"TMY-{cd['copper_spec']['spec']}"
            table_data.append({
                '序号': idx,
                '名称': '铜排',
                '型号': copper_spec_str,
                '数量': '—',
                '单价(元)': '—',
                '金额(元)': f"¥{cd['total_cost']:,.0f}",
                '品牌': '江西/金来',
            })
            idx += 1

            # 电缆行
            if result['cable_cost'] > 0:
                table_data.append({
                    '序号': idx,
                    '名称': '电缆',
                    '型号': '—',
                    '数量': '—',
                    '单价(元)': '—',
                    '金额(元)': f"¥{result['cable_cost']:,.0f}",
                    '品牌': '',
                })
                idx += 1

            # 辅助材料行
            table_data.append({
                '序号': idx,
                '名称': result['accessory_label'],
                '型号': '—',
                '数量': '—',
                '单价(元)': '—',
                '金额(元)': f"¥{result['accessory_cost']:,.0f}",
                '品牌': '湖北/恒晟',
            })
            idx += 1

            # 成套费行
            table_data.append({
                '序号': idx,
                '名称': '成套费（组装费、企业管理费、税收、利润等）',
                '型号': f"阶梯利润率",
                '数量': '—',
                '单价(元)': '—',
                '金额(元)': f"¥{result['total_profit']:,.0f}",
                '品牌': '宁波/海越',
            })
            idx += 1

            # 箱体行
            breaker_count = sum(c['qty'] for c in components if '断路器' in c.get('type', ''))
            cabinet_label = f"箱体:GCS({result['name']}、{breaker_count}路)"
            table_data.append({
                '序号': idx,
                '名称': cabinet_label,
                '型号': '—',
                '数量': '—',
                '单价(元)': '—',
                '金额(元)': f"¥{result['cabinet_cost']:,.0f}",
                '品牌': '宁波/海越',
            })

            # 合计行
            table_data.append({
                '序号': '',
                '名称': '**合计**',
                '型号': '',
                '数量': '',
                '单价(元)': '',
                '金额(元)': f"**¥{result['grand_total']:,.0f}**",
                '品牌': '',
            })

            st.dataframe(pd.DataFrame(table_data), use_container_width=True, hide_index=True)

            # 铜排信息（折叠）
            with st.expander("📐 铜排详情"):
                st.write(f"""
                - 总电流: {result['total_current']:,.0f}A
                - 出线路数: {result['outgoing_circuits']}路
                - 降容系数: {result['derate']}
                - 降容后电流: {result['reduced_current']:,.0f}A
                - 铜排规格: {copper_spec_str}
                - 铜排截面积: {cd['copper_area_cm2']} cm²
                - 铜排费用合计: ¥{cd['total_cost']:,.0f}（含主母线+≥250A出线）
                - 电缆费用(≤160A): ¥{result['cable_cost']:,.0f}
                """)

            # 详细计算过程（折叠）
            with st.expander("📝 详细计算过程", expanded=False):
                st.markdown(f"**铜价:** {copper_price*1000:,.0f}元/吨 = {copper_price}元/kg")
                st.markdown(f"**柜宽:** {cab['width']}m")

                # 总电流
                st.markdown("#### ① 总电流")
                for c in components:
                    if '断路器' in c.get('type', ''):
                        st.code(f"{c['model']}  {c['current']}A × {c['qty']} = {c['current']*c['qty']}A")
                    else:
                        st.code(f"{c['model']}  ({c['type']}) → 不计入")
                st.code(f"总电流 = {result['total_current']}A")

                # 分散系数
                st.markdown("#### ② 分散系数")
                st.code(f"出线路数 = {result['outgoing_circuits']}路(断路器) → 分散系数 = {result['derate']}")
                st.code(f"降容后电流 = {result['total_current']}A × {result['derate']} = {result['reduced_current']}A")

                # 铜排选型
                st.markdown("#### ③ 铜排选型")
                st.code(f"降容后 {result['reduced_current']}A → {copper_spec_str}  截面积 {cd['copper_area_cm2']}cm²")

                # 主母线
                st.markdown("#### ④ 主母线")
                busbar_sub = cd['phase_cost'] + cd['neutral_cost'] + cd['ground_cost']
                st.code(f"""三相ABC: 7m × {cd['copper_area_cm2']}cm² × {copper_price} × 8.9 / 10 = {cd['phase_cost']:,.2f}元
零线N:   2m × {cd['copper_area_cm2']/2}cm² × {copper_price} × 8.9 / 10 = {cd['neutral_cost']:,.2f}元
地线PE:  2m × {cd['pe_area_cm2']}cm² × {copper_price} × 8.9 / 10 = {cd['ground_cost']:,.2f}元  (PE按国标)
主母线小计: {busbar_sub:,.2f}元""")

                # ≥250A出线
                st.markdown("#### ⑤ ≥250A铜排出线")
                for c in components:
                    if '断路器' in c.get('type', '') and c['current'] >= 250:
                        from app import get_copper_spec_by_current as _gspec
                        _s = _gspec(c['current'])
                        _cost = 2.5 * c['qty'] * _s['area_cm2'] * copper_price * 8.9 * 3 / 10
                        st.code(f"{_s['spec']}  {c['model']}  {c['current']}A × {c['qty']}: 2.5m × {_s['area_cm2']}cm² × {copper_price} × 8.9 × 3 / 10 = {_cost:,.2f}元")
                st.code(f"出线小计: {cd['high_current_copper_cost']:,.2f}元")

                # 电缆
                st.markdown("#### ⑥ 电缆(≤160A)")
                for c in components:
                    if '断路器' in c.get('type', '') and c['current'] < 250 and c['current'] > 0:
                        from app import get_breaker_poles as _gpoles, load_breaker_cable_params as _lcp
                        _cp = _lcp()
                        _area = _cp.get(c['current'], 0)
                        if _area > 0:
                            _poles = _gpoles(c.get('model', ''))
                            _cl = _area * c['qty'] * (_poles * 0.7) * copper_price * 8.9 / 1000
                            st.code(f"{c['model']}  {c['current']}A × {c['qty']}: {_area}mm² × {c['qty']} × {_poles*0.7}m × {copper_price} × 8.9 / 1000 = {_cl:,.2f}元")
                st.code(f"电缆小计: {result['cable_cost']:,.2f}元")

                # 汇总
                st.markdown("#### ⑧ 费用汇总")
                st.code(f"""元器件: ¥{result['comp_cost']:,.0f}
铜排:   ¥{cd['total_cost']:,.0f}（主母线{busbar_sub:,.0f} + 出线{cd['high_current_copper_cost']:,.0f}）
电缆:   ¥{result['cable_cost']:,.0f}
辅材:   ¥{result['accessory_cost']:,.0f}
箱体:   ¥{result['cabinet_cost']:,.0f}
总成本: ¥{result['total_cost']:,.0f}
成套费: ¥{result['total_profit']:,.0f}
含税价: ¥{result['grand_total']:,.0f}""")

            project_total_cost += result['total_cost']
            project_total_profit += result['total_profit']

    # ─── 项目汇总 ───
    st.divider()
    st.subheader("📈 项目汇总")

    effective_rate = (project_total_profit / project_total_cost * 100) if project_total_cost > 0 else 0
    final_price = project_total_cost + project_total_profit
    tax_price = round(final_price * 1.13, 0)

    # 建议利润率（阶梯公式算出的）
    suggested_rate = effective_rate

    # 可编辑利润率
    col_rate, _, col_diff = st.columns([2, 1, 2])
    with col_rate:
        manual_rate = st.number_input(
            "手动调整利润率(%)", min_value=0.0, max_value=100.0,
            value=round(suggested_rate, 1), step=0.5, format="%.1f",
            help="建议利润率为阶梯公式自动计算值，可手动调整"
        ) / 100

    # 计算手动利润率下的价格
    manual_profit = project_total_cost * manual_rate
    manual_final = project_total_cost + manual_profit
    manual_tax = round(manual_final * 1.13, 0)
    price_diff = manual_final - final_price

    with col_diff:
        if abs(manual_rate - suggested_rate/100) > 0.001:
            diff_color = "🟢" if price_diff > 0 else "🔴"
            st.metric(
                f"价格差距（手动 vs 建议）",
                f"{diff_color} ¥{price_diff:+,.0f}",
                delta=f"利润率差 {((manual_rate - suggested_rate/100)*100):+.1f}%"
            )

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("总成本", f"¥{project_total_cost:,.0f}")
    c2.metric("建议成套费", f"¥{project_total_profit:,.0f}", delta=f"利润率 {suggested_rate:.1f}%")
    c3.metric("手动成套费", f"¥{manual_profit:,.0f}", delta=f"利润率 {manual_rate*100:.1f}%")
    c4.metric("柜子数量", f"{len(cabinet_results)}台")

    # 两种报价对比
    col_s, col_m = st.columns(2)
    with col_s:
        st.metric("建议含税报价", f"¥{tax_price:,.0f}")
    with col_m:
        st.metric("手动含税报价", f"¥{manual_tax:,.0f}", delta=f"¥{manual_tax - tax_price:+,.0f}")

    # 汇总表
    with st.expander("📋 柜子费用汇总"):
        summary_data = []
        for r in cabinet_results:
            summary_data.append({
                '柜号': r['name'],
                '类型': r['type'],
                '柜宽(m)': r['width'],
                '元器件': f"¥{r['comp_cost']:,.0f}",
                '铜排': f"¥{r['copper_cost']:,.0f}",
                '辅助材料': f"¥{r['accessory_cost']:,.0f}",
                '箱体': f"¥{r['cabinet_cost']:,.0f}",
                '成本小计': f"¥{r['total_cost']:,.0f}",
                '成套费': f"¥{r['total_profit']:,.0f}",
                '合计': f"¥{r['grand_total']:,.0f}",
            })
        summary_data.append({
            '柜号': '合计', '类型': '', '柜宽(m)': '',
            '元器件': '', '铜排': '', '辅助材料': '', '箱体': '',
            '成本小计': f"¥{project_total_cost:,.0f}",
            '成套费': f"¥{project_total_profit:,.0f}",
            '合计': f"¥{final_price:,.0f}",
        })
        st.dataframe(pd.DataFrame(summary_data), use_container_width=True, hide_index=True)

    # CSV导出
    csv_data = generate_project_csv(cabinet_list, cabinet_results, project_total_cost,
                                    project_total_profit, final_price, tax_price)
    st.download_button(
        "📥 导出项目报价明细 (CSV)",
        data=csv_data,
        file_name=f"{project_name}_成本报价.csv",
        mime="text/csv",
    )

def generate_project_csv(cabinet_list, cabinet_results, total_cost, total_profit, final_price, tax_price):
    """生成整个项目的CSV导出"""
    profit_rate_display = f"{total_profit/total_cost*100:.1f}%" if total_cost > 0 else "0%"
    lines = [f"项目名称,{st.session_state.project_name or '未命名项目'}", f"整体利润率,{profit_rate_display}", ""]

    for cab, result in zip(cabinet_list, cabinet_results):
        if not cab['components']:
            continue
        lines.append(f"=== {result['name']} ({result['type']}, {result['width']}m) ===")
        lines.append("序号,名称,型号,数量,单价(元),金额(元),品牌")
        for i, c in enumerate(cab['components']):
            if c['qty'] == 0:  # 导出时隐藏数量为0的
                continue
            amount = round(c['unit_price'], 0) * c['qty']
            lines.append(f"{i+1},{c['name']},{c['model']},{c['qty']},{round(c['unit_price'],0)},{amount},{c['brand']}")
        idx = len(cab['components']) + 1
        cd = result['copper_detail']
        lines.append(f"{idx},铜排,TMY-{cd['copper_spec']['spec']},—,—,{result['copper_cost']},江西/金来")
        idx += 1
        lines.append(f"{idx},{result['accessory_label']},—,—,—,{result['accessory_cost']},湖北/恒晟")
        idx += 1
        lines.append(f"{idx},成套费（组装费、企业管理费、税收、利润等）,阶梯利润率,—,—,{result['total_profit']},宁波/海越")
        idx += 1
        breaker_count = sum(c['qty'] for c in cab['components'] if '断路器' in c.get('type', ''))
        lines.append(f"{idx},箱体:GCS({result['name']}、{breaker_count}路),—,—,—,{result['cabinet_cost']},宁波/海越")
        lines.append(f"合计,,,,,,{result['grand_total']}")
        lines.append("")

    lines.append("=== 项目汇总 ===")
    lines.append(f"总成本,{total_cost}")
    lines.append(f"成套费（利润）,{total_profit}")
    lines.append(f"不含税报价,{final_price}")
    lines.append(f"含税报价(13%),{tax_price}")

    return "\n".join(lines)


if __name__ == "__main__":
    main()
