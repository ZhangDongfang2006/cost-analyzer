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
from pathlib import Path

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

# ─── 数据加载 ─────────────────────────────────────────────
@st.cache_data
def load_price_db():
    """加载产品价格数据库"""
    db_path = DATA_DIR / "price_db.json"
    if db_path.exists():
        with open(db_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    df = pd.read_excel(BASE_DIR / "成本.xlsx", sheet_name='库', usecols='A:F', nrows=2637)
    df = df.dropna(subset=['型号', '名称']).reset_index(drop=True)
    return df.to_dict(orient='records')

@st.cache_data
def load_copper_specs():
    """加载铜排规格映射表"""
    return [
        (0,    500,  0.18),
        (500,  755,  0.24),
        (755,  990,  0.36),
        (990,  1160, 0.48),
        (1160, 1500, 0.60),
        (1500, 99999, 0.72),
    ]

@st.cache_data
def load_breaker_cable_params():
    """加载断路器电缆参数: 电流(A) → 电缆宽度(mm)"""
    return {
        63: 80, 100: 120, 160: 120, 250: 192, 400: 192,
        500: 401, 630: 401, 800: 0, 1000: 0, 1250: 0,
    }

# ─── 核心计算函数 ─────────────────────────────────────────

def lookup_price(model: str, db: list) -> dict:
    """根据型号查找价格"""
    model = str(model).strip()
    for item in db:
        if str(item.get('型号', '')).strip() == model:
            return {
                'name': item.get('名称', ''),
                'unit_price': float(item.get('单价', 0) or 0),
                'retail_price': float(item.get('面价', 0) or 0),
                'brand': item.get('产地/厂家', ''),
            }
    return None

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

def get_copper_spec_by_current(total_current: float) -> float:
    specs = load_copper_specs()
    for low, high, spec in specs:
        if low <= total_current < high:
            return spec
    return 0.72

def get_copper_threshold_by_current(total_current: float) -> float:
    thresholds = [500, 755, 990, 1160, 1500]
    for i, t in enumerate(thresholds):
        if total_current < t:
            return thresholds[i-1] if i > 0 else thresholds[0]
    return 1500

def calc_cable_cost(breaker_type: str, current: int, qty: int,
                    cable_params: dict, copper_price: float, cabinet_width: float) -> float:
    if qty <= 0 or current <= 0:
        return 0.0
    if breaker_type == 'frame':
        copper_spec = get_copper_spec_by_current(current)
        return 2.5 * qty * copper_spec * copper_price * 8.9 * 3
    else:
        cable_width = cable_params.get(current, 0)
        if cable_width == 0:
            return 0.0
        factor = (cabinet_width / 105 - 1) / 2 + 1
        return cable_width * qty * factor

def calc_copper_busbar_cost(total_cable_cost: float, total_current: float,
                            copper_price: float) -> dict:
    L = get_copper_threshold_by_current(total_current)
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
        'copper_threshold': L,
        'total_current': round(total_current, 1),
    }

def calc_profit(amount: float) -> float:
    if amount <= 0:
        return 0.0
    if amount < 10000:
        return amount * 0.3
    elif amount > 49999.99:
        return amount * 0.1
    else:
        return (0.29 - 0.000003125 * amount) * amount

def calc_accessory_cost(outgoing_circuits: int) -> float:
    """辅助材料费用: 固定值 2266 = 63×2 + 100 + 250×2 + 400×3"""
    return 63 * 2 + 100 + 250 * 2 + 400 * 3  # 固定2266

def calc_single_cabinet(cabinet: dict, copper_price: float, db: list) -> dict:
    """计算单台柜子的所有费用"""
    components = cabinet['components']
    cable_params = load_breaker_cable_params()
    width = cabinet['width']

    # 1. 元器件费用
    total_comp_cost = 0
    total_cable_cost = 0
    total_current = 0

    for comp in components:
        rounded_price = round(comp['unit_price'], 0)
        amount = rounded_price * comp['qty']
        total_comp_cost += amount

        breaker_type = 'frame' if '框架' in comp.get('type', '') else 'mccb'
        cable = calc_cable_cost(breaker_type, comp['current'], comp['qty'],
                               cable_params, copper_price, width)
        total_cable_cost += cable
        total_current += comp['current'] * comp['qty']

    # 2. 铜排成本（根据柜内总电流独立计算）
    outgoing_circuits = sum(c['qty'] for c in components
                            if '断路器' in c.get('type', '') and '框架' not in c.get('type', ''))
    # 降容系数：根据出线路数
    if outgoing_circuits >= 6:
        derate = 0.7
    elif outgoing_circuits >= 4:
        derate = 0.75
    elif outgoing_circuits >= 2:
        derate = 0.8
    else:
        derate = 1.0

    reduced_current = total_current * derate
    copper_detail = calc_copper_busbar_cost(total_cable_cost, reduced_current, copper_price)

    # 3. 辅助材料
    accessory_cost = calc_accessory_cost(outgoing_circuits)

    # 4. 箱体费用: 2200 + 400×元器件总数量
    total_qty = sum(c['qty'] for c in components)
    cabinet_cost = 2200 + 400 * total_qty

    # 5. 小计（不含利润）
    subtotal = total_comp_cost + copper_detail['total_cost'] + accessory_cost + cabinet_cost

    # 6. 利润（各项分别算）
    comp_profit = calc_profit(total_comp_cost)
    copper_profit = calc_profit(copper_detail['total_cost'])
    accessory_profit = calc_profit(accessory_cost)
    cabinet_profit = calc_profit(cabinet_cost)
    total_profit = comp_profit + copper_profit + accessory_profit + cabinet_profit

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
        'subtotal': subtotal,
        'comp_profit': comp_profit,
        'copper_profit': copper_profit,
        'accessory_profit': accessory_profit,
        'cabinet_profit': cabinet_profit,
        'total_profit': total_profit,
        'total_current': total_current,
        'reduced_current': reduced_current,
        'outgoing_circuits': outgoing_circuits,
        'derate': derate,
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
    db = load_price_db()

    st.title("⚡ 配电设备成本计算系统")
    st.caption("多柜项目模式 · 自动计算铜排成本 · 智能匹配元器件价格 · 阶梯利润计算")

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
            results = [item for item in db
                       if search_term.upper() in str(item.get('型号', '')).upper()
                       or search_term in str(item.get('名称', ''))]
            if results:
                st.write(f"找到 {len(results)} 个结果（显示前15个）")
                for r in results[:15]:
                    col1, col2 = st.columns([3, 1])
                    with col1:
                        st.text(f"{r.get('型号', '')} | {r.get('名称', '')}")
                    with col2:
                        st.text(f"¥{r.get('单价', 0):.0f}")

    # ─── 主区域 3个Tab ───
    tab1, tab2, tab3 = st.tabs(["📋 项目配置", "⚡ 元器件管理", "📊 成本分析报告"])

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

                col1, col2, col3, col4, col5, col6 = st.columns([0.5, 2, 1.5, 1, 1, 0.8])
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
                col1, col2, col3 = st.columns([2, 1, 1])

                with col1:
                    model_input = st.text_input("型号", key="model_input",
                                                placeholder="输入元器件型号")
                    if model_input:
                        suggestions = [item for item in db
                                       if model_input.upper() in str(item.get('型号', '')).upper()][:5]
                        if suggestions:
                            for s in suggestions:
                                if st.button(f"选择: {s.get('型号', '')}", key=f"sel_{s.get('型号', '')}"):
                                    st.session_state.model_input = s.get('型号', '')
                                    st.rerun()

                with col2:
                    qty_input = st.number_input("数量", min_value=0, value=1, key="qty_input")

                with col3:
                    name_input = st.text_input("名称", key="name_input", placeholder="自动填充")
                    if model_input:
                        match = lookup_price(model_input, db)
                        if match:
                            st.session_state.name_input = match['name']

                col_a, col_b = st.columns([1, 1])
                with col_a:
                    auto_current = 0
                    if model_input:
                        auto_current = extract_current_from_model(model_input)
                    if auto_current > 0 and st.session_state.get('current_input', 0) == 0:
                        st.session_state.current_input = auto_current
                    current_input = st.number_input("额定电流 (A)", min_value=0, value=0,
                                                    key="current_input")
                    if auto_current > 0 and current_input == auto_current:
                        st.caption(f"💡 自动识别: {auto_current}A")

                with col_b:
                    breaker_type = st.selectbox("类型", ["塑壳断路器", "框架断路器", "电流互感器",
                                                         "数显仪表", "其他"], key="type_input")

                if st.button("✅ 添加到清单", type="primary", use_container_width=True):
                    if model_input and qty_input > 0:
                        match = lookup_price(model_input, db)
                        current = current_input if current_input > 0 else extract_current_from_model(model_input)
                        component = {
                            'model': model_input,
                            'name': name_input or (match['name'] if match else '未知'),
                            'qty': qty_input,
                            'current': current,
                            'type': breaker_type,
                            'unit_price': match['unit_price'] if match else 0,
                            'retail_price': match['retail_price'] if match else 0,
                            'brand': match['brand'] if match else '未找到',
                            'matched': match is not None,
                        }
                        st.session_state.cabinet_list[idx]['components'].append(component)
                        st.success(f"✅ 已添加到 {cab['name']}: {model_input} × {qty_input}")
                        st.rerun()
                    else:
                        st.warning("请填写型号和数量")

            # 显示元器件清单
            components = cab['components']
            if components:
                st.divider()
                st.subheader(f"📋 {cab['name']} - 元器件清单")

                display_data = []
                for i, comp in enumerate(components):
                    amount = round(comp['unit_price'], 0) * comp['qty']
                    display_data.append({
                        '序号': i + 1,
                        '名称': comp['name'],
                        '型号': comp['model'],
                        '电流(A)': comp['current'],
                        '数量': comp['qty'],
                        '单价': f"¥{comp['unit_price']:.0f}",
                        '金额': f"¥{amount:,.0f}",
                        '品牌': comp['brand'],
                        '状态': '✅' if comp['matched'] else '⚠️未匹配',
                    })

                df_display = pd.DataFrame(display_data)
                st.dataframe(df_display, use_container_width=True, hide_index=True)

                # 删除按钮
                cols = st.columns(min(len(components), 10))
                for i, col in enumerate(cols):
                    if i < len(components):
                        comp = components[i]
                        if col.button(f"🗑️{comp['model'][:12]}", key=f"del_{idx}_{i}"):
                            st.session_state.cabinet_list[idx]['components'].pop(i)
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
            run_project_report(st.session_state.cabinet_list, copper_price, db)
        else:
            st.info("👈 请在'元器件管理'中点击'计算成本分析报告'")

def run_project_report(cabinet_list: list, copper_price: float, db: list):
    """生成项目成本分析报告"""
    project_name = st.session_state.project_name or "未命名项目"
    st.header(f"📊 成本分析报告 - {project_name}")

    # 计算每台柜子
    cabinet_results = []
    for cab in cabinet_list:
        if cab['components']:
            result = calc_single_cabinet(cab, copper_price, db)
            cabinet_results.append(result)

    if not cabinet_results:
        st.warning("没有包含元器件的柜子")
        return

    # ─── 每台柜子明细 ───
    st.subheader("🗄️ 各柜明细")

    project_total_cost = 0
    project_total_profit = 0

    for result in cabinet_results:
        with st.expander(f"{'🔴 ' if result == cabinet_results[0] else ''}{result['name']} ({result['type']}, {result['width']}m)", expanded=(result == cabinet_results[0])):
            c1, c2, c3 = st.columns(3)
            c1.metric("元器件费用", f"¥{result['comp_cost']:,.0f}")
            c2.metric("电缆费用", f"¥{result['cable_cost']:,.0f}")
            c3.metric("铜排成本", f"¥{result['copper_cost']:,.0f}")

            c4, c5, c6 = st.columns(3)
            c4.metric(result['accessory_label'], f"¥{result['accessory_cost']:,.0f}")
            c5.metric("箱体费用", f"¥{result['cabinet_cost']:,.0f}")
            c6.metric("小计", f"¥{result['subtotal']:,.0f}")

            # 铜排信息
            with st.expander("📐 铜排详情"):
                cd = result['copper_detail']
                st.write(f"""
                - 总电流: {result['total_current']:,.0f}A
                - 出线路数: {result['outgoing_circuits']}路
                - 降容系数: {result['derate']}
                - 降容后电流: {result['reduced_current']:,.0f}A
                - 铜排规格: TMY-{cd['copper_spec']*100:.0f}×{cd['copper_spec']*100:.0f}
                - 三相铜排: ¥{cd['phase_cost']:,.0f}
                - 零线: ¥{cd['neutral_cost']:,.0f}
                - 地线: ¥{cd['ground_cost']:,.0f}
                - 电缆费用: ¥{cd['cable_cost']:,.0f}
                """)

            # 利润明细
            with st.expander("💰 利润明细"):
                profit_items = [
                    ('元器件', result['comp_cost'], result['comp_profit']),
                    ('铜排', result['copper_cost'], result['copper_profit']),
                    ('辅助材料', result['accessory_cost'], result['accessory_profit']),
                    ('箱体', result['cabinet_cost'], result['cabinet_profit']),
                ]
                profit_df = pd.DataFrame([
                    {'项目': name, '成本': cost, '利润': profit,
                     '利润率': f"{profit/cost*100:.1f}%" if cost > 0 else "0%"}
                    for name, cost, profit in profit_items
                ])
                st.dataframe(profit_df, use_container_width=True, hide_index=True)

            project_total_cost += result['subtotal']
            project_total_profit += result['total_profit']

    # ─── 项目汇总 ───
    st.divider()
    st.subheader("📈 项目汇总")

    profit_rate = (project_total_profit / project_total_cost * 100) if project_total_cost > 0 else 0
    final_price = project_total_cost + project_total_profit
    tax_price = round(final_price * 1.13, 0)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("总成本", f"¥{project_total_cost:,.0f}")
    c2.metric("总利润", f"¥{project_total_profit:,.0f}")
    c3.metric("利润率", f"{profit_rate:.1f}%")
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
                '电缆': f"¥{r['cable_cost']:,.0f}",
                '铜排': f"¥{r['copper_cost']:,.0f}",
                '辅助材料': f"¥{r['accessory_cost']:,.0f}",
                '箱体': f"¥{r['cabinet_cost']:,.0f}",
                '小计': f"¥{r['subtotal']:,.0f}",
                '利润': f"¥{r['total_profit']:,.0f}",
            })
        summary_data.append({
            '柜号': '合计', '类型': '', '柜宽(m)': '',
            '元器件': '', '电缆': '', '铜排': '', '辅助材料': '', '箱体': '',
            '小计': f"¥{project_total_cost:,.0f}",
            '利润': f"¥{project_total_profit:,.0f}",
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
    lines = [f"项目名称,{st.session_state.project_name or '未命名项目'}", ""]

    for cab, result in zip(cabinet_list, cabinet_results):
        if not cab['components']:
            continue
        lines.append(f"=== {result['name']} ({result['type']}, {result['width']}m) ===")
        lines.append("序号,名称,型号,电流(A),数量,单价,金额,品牌")
        for i, c in enumerate(cab['components']):
            amount = round(c['unit_price'], 0) * c['qty']
            lines.append(f"{i+1},{c['name']},{c['model']},{c['current']},{c['qty']},{round(c['unit_price'],0)},{amount},{c['brand']}")
        lines.append(f"元器件费用,{result['comp_cost']}")
        lines.append(f"电缆费用,{result['cable_cost']}")
        lines.append(f"铜排成本,{result['copper_cost']}")
        lines.append(f"辅助材料,{result['accessory_cost']}")
        lines.append(f"箱体费用,{result['cabinet_cost']}")
        lines.append(f"小计,{result['subtotal']}")
        lines.append(f"利润,{result['total_profit']}")
        lines.append("")

    lines.append("=== 项目汇总 ===")
    lines.append(f"总成本,{total_cost}")
    lines.append(f"总利润,{total_profit}")
    lines.append(f"不含税报价,{final_price}")
    lines.append(f"含税报价(13%),{tax_price}")

    return "\n".join(lines)


if __name__ == "__main__":
    main()
