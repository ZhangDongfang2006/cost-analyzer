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
    {'spec': '40×5',  'width': 40, 'thickness': 5,  'area_mm2': 200, 'area_cm2': 0.20, 'current': 615},
    {'spec': '50×5',  'width': 50, 'thickness': 5,  'area_mm2': 250, 'area_cm2': 0.25, 'current': 755},
    {'spec': '60×6',  'width': 60, 'thickness': 6,  'area_mm2': 360, 'area_cm2': 0.36, 'current': 990},
    {'spec': '60×8',  'width': 60, 'thickness': 8,  'area_mm2': 480, 'area_cm2': 0.48, 'current': 1160},
    {'spec': '80×8',  'width': 80, 'thickness': 8,  'area_mm2': 640, 'area_cm2': 0.64, 'current': 1490},
    {'spec': '80×10', 'width': 80, 'thickness': 10, 'area_mm2': 800, 'area_cm2': 0.80, 'current': 1670},
]

@st.cache_data
def load_breaker_cable_params():
    """加载断路器电缆参数: 电流(A) → 电缆宽度(mm)"""
    return {
        63: 80, 100: 120, 160: 120, 250: 192, 400: 192,
        500: 401, 630: 401, 800: 0, 1000: 0, 1250: 0,
    }

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

def calc_cable_cost(breaker_type: str, current: int, qty: int,
                    cable_params: dict, copper_price: float, cabinet_width: float) -> float:
    if qty <= 0 or current <= 0:
        return 0.0
    if breaker_type == 'frame':
        return 0.0  # 框架断路器用铜排，不计入电缆费用
    else:
        cable_width = cable_params.get(current, 0)
        if cable_width == 0:
            return 0.0
        factor = (cabinet_width / 105 - 1) / 2 + 1
        return cable_width * qty * factor

def calc_copper_busbar_cost(total_cable_cost: float, total_current: float,
                            copper_price: float) -> dict:
    L = get_copper_area_by_current(total_current)  # area_cm2
    K = copper_price
    density = 8.9
    phase_cost = 7 * L * K * density
    neutral_cost = 2 * (L / 2) * K * density
    ground_cost = 2 * (L / 4) * K * density
    copper_cost = round(phase_cost + neutral_cost + ground_cost + total_cable_cost, 0)
    copper_spec = get_copper_spec_by_current(total_current)
    return {
        'total_cost': copper_cost,
        'phase_cost': round(phase_cost, 2),
        'neutral_cost': round(neutral_cost, 2),
        'ground_cost': round(ground_cost, 2),
        'cable_cost': round(total_cable_cost, 2),
        'copper_spec': copper_spec,
        'copper_area_cm2': L,
        'total_current': round(total_current, 1),
    }

def calc_profit(amount: float, rate: float = 0.16) -> float:
    if amount <= 0:
        return 0.0
    return amount * rate

def calc_accessory_cost(outgoing_circuits: int) -> float:
    """辅助材料费用: 固定值 2266 = 63×2 + 100 + 250×2 + 400×3"""
    return 63 * 2 + 100 + 250 * 2 + 400 * 3  # 固定2266

def calc_single_cabinet(cabinet: dict, copper_price: float, profit_rate: float = 0.16) -> dict:
    """计算单台柜子的所有费用"""
    components = cabinet['components']
    cable_params = load_breaker_cable_params()
    width = cabinet['width']

    # 1. 元器件费用
    total_comp_cost = 0
    total_cable_cost = 0
    total_current = 0
    frame_copper_cost = 0  # 框架断路器铜排费用

    for comp in components:
        rounded_price = round(comp['unit_price'], 0)
        amount = rounded_price * comp['qty']
        total_comp_cost += amount

        breaker_type = 'frame' if '框架' in comp.get('type', '') else 'mccb'
        cable = calc_cable_cost(breaker_type, comp['current'], comp['qty'],
                               cable_params, copper_price, width)
        total_cable_cost += cable
        total_current += comp['current'] * comp['qty']

        # 框架断路器出线用铜排
        if breaker_type == 'frame' and comp['qty'] > 0 and comp['current'] > 0:
            spec = get_copper_spec_by_current(comp['current'])
            frame_copper_cost += 2.5 * comp['qty'] * spec['area_cm2'] * copper_price * 8.9 * 3

    # 2. 铜排成本（根据柜内总电流独立计算）
    # 出线路数：优先使用数显仪表数量，如无仪表则使用断路器数量
    meter_count = sum(c['qty'] for c in components if '仪表' in c.get('type', '') or '数显' in c.get('type', ''))
    breaker_count_for_circuits = sum(c['qty'] for c in components if '断路器' in c.get('type', ''))
    # 出线路数始终按断路器数量计算
    outgoing_circuits = breaker_count_for_circuits
    # 降容系数：根据出线路数（Excel: IF(D50>9,I50*0.7,IF(D50>5,I50*0.8,I50))）
    if outgoing_circuits > 9:
        derate = 0.7
    elif outgoing_circuits > 5:
        derate = 0.8
    else:
        derate = 1.0

    # 数显仪表电缆费用：柜宽×4×铜价×8.9 + 柜宽×0.8×铜价×8.9
    has_meter = meter_count > 0
    if has_meter:
        meter_cable_cost = width * 4 * copper_price * 8.9 + width * 0.8 * copper_price * 8.9
        total_cable_cost += meter_cable_cost

    reduced_current = total_current * derate
    copper_detail = calc_copper_busbar_cost(total_cable_cost, reduced_current, copper_price)
    # 框架断路器铜排费用加到铜排成本中
    copper_detail['total_cost'] = round(copper_detail['total_cost'] + frame_copper_cost, 0)
    copper_detail['frame_copper_cost'] = round(frame_copper_cost, 2)

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
        accessory_cost = calc_accessory_cost(outgoing_circuits)  # fallback 1926
        accessory_matched = False

    # 4. 箱体费用: 2200 + 400×断路器数量（只算断路器）
    breaker_count = sum(c['qty'] for c in components if '断路器' in c.get('type', ''))
    cabinet_cost = 2200 + 400 * breaker_count

    # 5. 小计（不含利润）
    subtotal = total_comp_cost + copper_detail['total_cost'] + accessory_cost + cabinet_cost

    # 6. 成套费（统一利润率）
    total_cost = subtotal  # 不含利润
    total_profit = calc_profit(total_cost, profit_rate)
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
        'profit_rate': profit_rate,
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
            col_a, col_b, col_c = st.columns([2, 2, 1])
            with col_a:
                new_name = st.text_input("柜号名称", placeholder="如: 1#出线柜", key="new_cab_name")
            with col_b:
                new_type = st.selectbox("柜类型", CABINET_TYPES, key="new_cab_type")
            with col_c:
                new_width = st.number_input("柜宽(米)", value=0.8, min_value=0.1, max_value=2.0,
                                            step=0.1, key="new_cab_width")

            if st.button("✅ 添加柜子", type="primary", use_container_width=True):
                if new_name:
                    st.session_state.cabinet_list.append({
                        'name': new_name,
                        'type': new_type,
                        'width': new_width,
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
            with st.expander("📦 批量导入", expanded=False):
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
                    text = st.text_area("每行一个元器件（Tab/空格分隔）", height=150,
                                        placeholder="名称\t型号\t数量\n型号\t数量", key="batch_text")
                    if text.strip():
                        for line in text.strip().split('\n'):
                            parts = re.split(r'[\t]+|[ \t]{2,}', line.strip())
                            parts = [p.strip() for p in parts if p.strip()]
                            if not parts:
                                continue
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
                        pdf.append({
                            '序号': i + 1,
                            '名称': item['name'],
                            '型号': item['model'],
                            '数量': item['qty'],
                                    '匹配单价': f"¥{item['unit_price']:.0f}" if item['matched'] else '-',
                            '状态': '✅已匹配' if item['matched'] else '⚠️未找到',
                        })
                    st.dataframe(pd.DataFrame(pdf), use_container_width=True, hide_index=True)

                    if st.button("✅ 确认导入", type="primary", use_container_width=True, key="batch_confirm"):
                        for item in preview_data:
                            current = extract_current_from_model(item['model'])
                            match = lookup_price(item['model'])
                            component = {
                                'model': item['model'],
                                'name': item['name'] or '未知',
                                'qty': item['qty'],
                                'current': current,
                                'type': '其他',
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
        with st.expander("🔌 2. 电缆费用公式（J列）", expanded=False):
            st.markdown("#### 塑壳断路器")
            st.code("电缆费 = 电缆宽度(mm) × 数量 × ((柜宽/105 - 1)/2 + 1)", language="text")
            st.markdown("""
            > **说明：** 缓冲系数 `((柜宽/105 - 1)/2 + 1)` 用于平衡铜价波动，铜价上涨时报价只涨涨幅的一半。
            """)
            cable_params = load_breaker_cable_params()
            cable_table = pd.DataFrame([
                {'电流(A)': k, '电缆宽度(mm)': v} for k, v in sorted(cable_params.items()) if v > 0
            ])
            st.dataframe(cable_table, use_container_width=True, hide_index=True)

            st.markdown("#### 框架断路器")
            st.code("电缆费 = 2.5 × 数量 × 铜排截面积(cm²) × 铜价 × 8.9 × 3", language="text")

            st.markdown("#### 数显仪表电缆费用")
            st.code("电缆费 = 柜宽 × 4 × 铜价 × 8.9 + 柜宽 × 0.8 × 铜价 × 8.9", language="text")
            st.markdown("> 仅当柜内包含数显仪表时计算此费用。")

        # 3. 铜排成本公式
        with st.expander("🟫 3. 铜排成本公式（H51）", expanded=False):
            st.code("""铜排总价 = ROUND(
  7 × 铜排截面积(cm²) × 铜价 × 8.9        // 三相主母线（7米）
  + 2 × (截面积/2) × 铜价 × 8.9            // 零线N（截面积为主母线一半）
  + 2 × (截面积/4) × 铜价 × 8.9            // 地线PE（截面积为主母线四分之一）
  + 总电缆费用, 0)""", language="text")
            st.markdown("""
            - **7** = 水平母线总长度（米），经验估算值
            - **8.9** = 铜密度 (g/cm³)
            - 铜排截面积根据降容后总电流从对照表选取
            """)

        # 4. 降容系数
        with st.expander("📉 4. 降容系数", expanded=False):
            st.code("降容后电流 = IF(出线路数>9, 总电流×0.7, IF(出线路数>5, 总电流×0.8, 总电流))", language="text")
            st.markdown("- 出线路数优先使用**数显仪表**数量")
            st.markdown("- 如无数显仪表，则使用**断路器**数量（塑壳+框架总数量）作为出线路数")

        # 5. 辅助材料费用
        with st.expander("🔧 5. 辅助材料费用", expanded=False):
            st.code("辅助材料 = 从价格库查找，名称格式：辅助材料（出线柜，N路出线）", language="text")
            st.markdown("- 未找到时使用默认值1926元并提示用户手动输入")

        # 6. 箱体费用
        with st.expander("📦 6. 箱体费用", expanded=False):
            st.code("箱体 = 2200 + 400 × 断路器总数量", language="text")

        # 7. 利润计算
        with st.expander("💰 7. 利润计算（统一利润率）", expanded=False):
            st.code("成套费 = (元器件 + 铜排 + 辅助材料 + 箱体) × 利润率", language="text")
            st.markdown("- 默认利润率: **16%**，可在成本分析报告中调整")
            st.markdown("- 成套费 = 组装费 + 企业管理费 + 税收 + 利润等")

        # 8. 最终报价
        with st.expander("🏷️ 8. 最终报价", expanded=False):
            st.code("""不含税报价 = 总成本 + 总利润
含税报价 = 不含税报价 × 1.13""", language="text")

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
    profit_rate = results[0].get('profit_rate', 0.16) if results else 0.16
    entry = {
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "copper_price": copper_price,
        "profit_rate": profit_rate,
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


def run_project_report(cabinet_list: list, copper_price: float):
    """生成项目成本分析报告"""
    project_name = st.session_state.project_name or "未命名项目"
    st.header(f"📊 成本分析报告 - {project_name}")

    # 利润率输入
    col_pr, col_empty = st.columns([1, 5])
    with col_pr:
        profit_rate = st.number_input("利润率(%)", min_value=0, max_value=100, value=16, step=1, format="%d") / 100

    # 计算每台柜子
    cabinet_results = []
    for cab in cabinet_list:
        if cab['components']:
            result = calc_single_cabinet(cab, copper_price, profit_rate)
            # 检查辅材是否匹配成功
            if not result.get('accessory_matched', False):
                st.warning(f"⚠️ 未在价格库中找到 {result['outgoing_circuits']}路出线辅助材料，请手动输入或更新价格库")
                result['accessory_cost'] = st.number_input(
                    f"🔧 {result['name']} 辅助材料价格（元）",
                    min_value=0.0, value=float(result['accessory_cost']),
                    step=100.0, format="%.0f", key=f"accessory_{result['name']}")
                # 重新计算
                result['total_cost'] = result['comp_cost'] + result['copper_cost'] + result['accessory_cost'] + result['cabinet_cost']
                result['total_profit'] = calc_profit(result['total_cost'], profit_rate)
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

    # ─── 每台柜子明细（匹配Excel格式） ───
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
                '金额(元)': f"¥{cd['total_cost'] - cd['cable_cost']:,.0f}",
                '品牌': '江西/金来',
            })
            idx += 1

            # 电缆行
            if cd['cable_cost'] > 0:
                table_data.append({
                    '序号': idx,
                    '名称': '电缆',
                    '型号': '—',
                    '数量': '—',
                    '单价(元)': '—',
                    '金额(元)': f"¥{cd['cable_cost']:,.0f}",
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
                '型号': f"利润率{profit_rate*100:.0f}%",
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
                - 三相铜排: ¥{cd['phase_cost']:,.0f}
                - 零线: ¥{cd['neutral_cost']:,.0f}
                - 地线: ¥{cd['ground_cost']:,.0f}
                - 电缆费用: ¥{cd['cable_cost']:,.0f}
                """)

            project_total_cost += result['total_cost']
            project_total_profit += result['total_profit']

    # ─── 项目汇总 ───
    st.divider()
    st.subheader("📈 项目汇总")

    effective_rate = (project_total_profit / project_total_cost * 100) if project_total_cost > 0 else 0
    final_price = project_total_cost + project_total_profit
    tax_price = round(final_price * 1.13, 0)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("总成本", f"¥{project_total_cost:,.0f}")
    c2.metric("成套费（利润）", f"¥{project_total_profit:,.0f}")
    c3.metric("利润率", f"{profit_rate*100:.0f}%")
    c4.metric("柜子数量", f"{len(cabinet_results)}台")

    c5, c6 = st.columns(2)
    c5.metric("不含税报价", f"¥{final_price:,.0f}")
    c6.metric("含税报价 (13%)", f"¥{tax_price:,.0f}")

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
                                    project_total_profit, final_price, tax_price, profit_rate)
    st.download_button(
        "📥 导出项目报价明细 (CSV)",
        data=csv_data,
        file_name=f"{project_name}_成本报价.csv",
        mime="text/csv",
    )

def generate_project_csv(cabinet_list, cabinet_results, total_cost, total_profit, final_price, tax_price, profit_rate=0.16):
    """生成整个项目的CSV导出"""
    lines = [f"项目名称,{st.session_state.project_name or '未命名项目'}", f"利润率,{profit_rate*100:.0f}%", ""]

    for cab, result in zip(cabinet_list, cabinet_results):
        if not cab['components']:
            continue
        lines.append(f"=== {result['name']} ({result['type']}, {result['width']}m) ===")
        lines.append("序号,名称,型号,数量,单价(元),金额(元),品牌")
        for i, c in enumerate(cab['components']):
            amount = round(c['unit_price'], 0) * c['qty']
            lines.append(f"{i+1},{c['name']},{c['model']},{c['qty']},{round(c['unit_price'],0)},{amount},{c['brand']}")
        idx = len(cab['components']) + 1
        cd = result['copper_detail']
        lines.append(f"{idx},铜排,TMY-{cd['copper_spec']['spec']},—,—,{result['copper_cost']},江西/金来")
        idx += 1
        lines.append(f"{idx},{result['accessory_label']},—,—,—,{result['accessory_cost']},湖北/恒晟")
        idx += 1
        lines.append(f"{idx},成套费（组装费、企业管理费、税收、利润等）,利润率{profit_rate*100:.0f}%,—,—,{result['total_profit']},宁波/海越")
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
