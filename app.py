#!/usr/bin/env python3
"""
配电设备成本计算系统 - Cost Calculator Web App
基于Excel公式逻辑，自动计算铜排成本和项目总价
"""

import streamlit as st
import pandas as pd
import json
import os
import re
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

# ─── 数据加载 ─────────────────────────────────────────────
@st.cache_data
def load_price_db():
    """加载产品价格数据库"""
    db_path = DATA_DIR / "price_db.json"
    if db_path.exists():
        with open(db_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    # 回退到直接读Excel
    df = pd.read_excel(BASE_DIR / "成本.xlsx", sheet_name='库', usecols='A:F', nrows=2637)
    df = df.dropna(subset=['型号', '名称']).reset_index(drop=True)
    return df.to_dict(orient='records')

@st.cache_data
def load_copper_specs():
    """加载铜排规格映射表"""
    # 电流阈值 → 铜排截面积(mm²)
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
        63: 80,
        100: 120,
        160: 120,
        250: 192,
        400: 192,
        500: 401,
        630: 401,
        800: 0,   # 框架断路器用不同公式
        1000: 0,
        1250: 0,
    }

# ─── 核心计算函数 ─────────────────────────────────────────

def lookup_price(model: str, db: list) -> dict:
    """根据型号查找价格（模拟VLOOKUP）"""
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
    # 框架断路器
    m = re.search(r'E1C\s*(\d+)', model, re.IGNORECASE)
    if m:
        return int(m.group(1))
    # 3WL系列（西门子框架）：缩写映射
    m = re.search(r'3WL\w*N(\d{2})', model, re.IGNORECASE)
    if m:
        wl_map = {'08': 800, '10': 1000, '12': 1250, '16': 1600, '20': 2000}
        return wl_map.get(m.group(1), 0)
    # MT系列（施耐德框架）
    m = re.search(r'MT\s*(\d+)', model, re.IGNORECASE)
    if m:
        return int(m.group(1))
    # AN系列（LS产电框架）
    m = re.search(r'AN-\d+D\d+-(\d+)A', model, re.IGNORECASE)
    if m:
        return int(m.group(1))
    # 塑壳断路器
    m = re.search(r'TM[DA]\s*(\d+)', model, re.IGNORECASE)
    if m:
        return int(m.group(1))
    # NSX系列（施耐德塑壳，必须在TM脱扣器之前匹配）
    m = re.search(r'NSX\s*(\d+)', model, re.IGNORECASE)
    if m:
        return int(m.group(1))
    # 3VL系列（西门子塑壳）
    m = re.search(r'3VL\w*(\d{3})', model, re.IGNORECASE)
    if m:
        return int(m.group(1))
    # CB系列（GE塑壳）
    m = re.search(r'CB\w*/(\d+)A', model, re.IGNORECASE)
    if m:
        return int(m.group(1))
    # CM1系列（常熟塑壳）
    m = re.search(r'CM1[-/](\d+)', model, re.IGNORECASE)
    if m:
        return int(m.group(1))
    # NM1系列（正泰塑壳）
    m = re.search(r'NM1[-/](\d+)', model, re.IGNORECASE)
    if m:
        return int(m.group(1))
    # 兜底：末尾数字+A
    m = re.search(r'(\d{3,4})A', model)
    if m:
        return int(m.group(1))
    return 0

def get_breaker_type(name: str) -> str:
    """判断断路器类型"""
    if '框架' in str(name):
        return 'frame'
    return 'mccb'

def calc_cable_cost(breaker_type: str, current: int, qty: int,
                    cable_params: dict, copper_price: float, cabinet_width: float) -> float:
    """计算单个断路器的电缆费用"""
    if qty <= 0 or current <= 0:
        return 0.0

    if breaker_type == 'frame':
        # 框架断路器: 2.5 * 数量 * 铜排规格 * 铜价 * 8.9 * 3
        # 铜排规格需要根据电流确定
        copper_spec = get_copper_spec_by_current(current)
        return 2.5 * qty * copper_spec * copper_price * 8.9 * 3
    else:
        # 塑壳断路器: 电缆宽度 * 数量 * ((柜宽/105-1)/2+1)
        cable_width = cable_params.get(current, 0)
        if cable_width == 0:
            return 0.0
        factor = (cabinet_width / 105 - 1) / 2 + 1
        return cable_width * qty * factor

def get_copper_spec_by_current(total_current: float) -> float:
    """根据总电流选择铜排截面积"""
    specs = load_copper_specs()
    for low, high, spec in specs:
        if low <= total_current < high:
            return spec
    return 0.72  # 默认最大规格

def get_copper_threshold_by_current(total_current: float) -> float:
    """根据总电流获取铜排截面对应的电流阈值（用于铜排成本计算）"""
    # 返回铜排宽度值，用于 L 列的逻辑
    # 映射: 电流 → 对应的L值
    thresholds = [500, 755, 990, 1160, 1500]
    specs = [0.18, 0.24, 0.36, 0.48, 0.60, 0.72]
    for i, (t, s) in enumerate(zip(thresholds, specs)):
        if total_current < t:
            return thresholds[i-1] if i > 0 else thresholds[0]
    return 1500

def calc_copper_busbar_cost(total_cable_cost: float, total_current: float,
                            copper_price: float) -> dict:
    """
    计算铜排成本（核心公式）
    
    原始公式 (H51):
    =ROUND((7)*L51*K39*8.9 + 2*L51/2*K39*8.9 + 2*L51/4*K39*8.9 + J51, 0)
    
    其中:
    - 7 = 柜内铜排总长度系数 (米)
    - L51 = 铜排截面积对应的电流阈值
    - K39 = 铜价 (元/单位)
    - 8.9 = 铜密度 (g/cm³)
    - L51/2 = N排(零线)截面积
    - L51/4 = PE排(地线)截面积
    - J51 = 总电缆费用
    """
    L = get_copper_threshold_by_current(total_current)
    K = copper_price  # 铜价
    density = 8.9      # 铜密度

    # 三相铜排成本: 7 * L * K * 8.9
    phase_cost = 7 * L * K * density
    # 零线(N): 2 * (L/2) * K * 8.9
    neutral_cost = 2 * (L / 2) * K * density
    # 地线(PE): 2 * (L/4) * K * 8.9
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
    """
    利润计算（阶梯式）
    
    原始公式:
    IF(金额<10000, 金额*0.3, IF(金额>49999.99, 金额*0.1, (0.29-0.000003125*金额)*金额))
    
    利润率:
    - 金额 < 10000: 30%
    - 金额 > 50000: 10%
    - 10000~50000: 29% - 0.0003125% * 金额 (线性递减)
    """
    if amount <= 0:
        return 0.0
    if amount < 10000:
        return amount * 0.3
    elif amount > 49999.99:
        return amount * 0.1
    else:
        return (0.29 - 0.000003125 * amount) * amount

def calc_cabinet_cost(copper_cost_detail: dict, copper_price: float) -> float:
    """
    计算箱体费用
    
    原始公式 (H54): =2200 + O51
    O51 = SUM(O37:O50) = 400 * 总数量
    """
    # O51 是固定的 400*数量，这里简化为基础箱体费用
    return 2200 + copper_cost_detail.get('cable_cost', 0) * 0  # 需要实际数量，暂用固定值

def calc_accessory_cost(outgoing_circuits: int) -> float:
    """
    计算辅助材料费用
    
    原始公式 (O52): =63*2 + 100 + 250*2 + 400*3
    这是固定值: 2266
    """
    # 按出线路数动态计算
    # 基础值 + 每路附加
    base = 63 * 2 + 100  # 接线端子等
    return base + outgoing_circuits * 400

# ─── UI ─────────────────────────────────────────────────

def main():
    st.title("⚡ 配电设备成本计算系统")
    st.caption("自动计算铜排成本 · 智能匹配元器件价格 · 阶梯利润计算")

    db = load_price_db()

    # ─── 侧边栏参数 ───
    with st.sidebar:
        st.header("⚙️ 项目参数")

        copper_price = st.number_input("铜价 (元)", value=100, min_value=0, step=1,
                                        help="当前铜价，影响铜排成本计算")
        cabinet_width = st.number_input("柜宽 (米)", value=0.8, min_value=0.1, max_value=2.0,
                                         step=0.1, help="配电柜宽度，影响电缆长度计算")
        st.divider()
        st.subheader("📋 快速添加")
        st.caption("从价格库搜索添加元器件")

        search_term = st.text_input("搜索型号或名称", placeholder="如: XT1N160 或 塑壳断路器")
        if search_term:
            results = [item for item in db
                       if search_term.upper() in str(item.get('型号', '')).upper()
                       or search_term in str(item.get('名称', ''))]
            if results:
                st.write(f"找到 {len(results)} 个结果（显示前20个）")
                for r in results[:20]:
                    col1, col2 = st.columns([3, 1])
                    with col1:
                        st.text(f"{r.get('型号', '')} | {r.get('名称', '')}")
                    with col2:
                        st.text(f"¥{r.get('单价', 0):.0f}")

    # ─── 主区域 ───
    tab1, tab2, tab3 = st.tabs(["📝 元器件清单", "📊 成本分析", "📖 使用说明"])

    with tab1:
        st.subheader("元器件清单")

        # 初始化session_state
        if 'components' not in st.session_state:
            st.session_state.components = []

        # 添加元器件表单
        with st.expander("➕ 添加元器件", expanded=True):
            col1, col2, col3 = st.columns([2, 1, 1])

            with col1:
                model_input = st.text_input("型号", key="model_input",
                                            placeholder="输入元器件型号，如 XT1N160 TMD 63 3P FF")
                # 自动补全建议
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
                name_input = st.text_input("名称", key="name_input",
                                           placeholder="自动填充或手动输入")
                # 查找建议名称
                if model_input:
                    match = lookup_price(model_input, db)
                    if match:
                        st.session_state.name_input = match['name']

            col_a, col_b = st.columns([1, 1])
            with col_a:
                # 自动从型号提取电流并填入输入框
                auto_current = 0
                if model_input:
                    auto_current = extract_current_from_model(model_input)
                if auto_current > 0 and st.session_state.get('current_input', 0) == 0:
                    st.session_state.current_input = auto_current
                current_input = st.number_input("额定电流 (A)", min_value=0, value=0,
                                                key="current_input",
                                                help="留空则自动从型号提取")
                if auto_current > 0 and current_input == auto_current:
                    st.caption(f"💡 已自动识别电流: {auto_current}A（可手动修改）")

            with col_b:
                breaker_type = st.selectbox("类型", ["塑壳断路器", "框架断路器", "电流互感器",
                                                     "数显仪表", "其他"],
                                            key="type_input")

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
                    st.session_state.components.append(component)
                    st.success(f"✅ 已添加: {model_input} × {qty_input}")
                    st.rerun()
                else:
                    st.warning("请填写型号和数量")

        # 显示元器件清单
        if st.session_state.components:
            st.divider()
            st.subheader("📋 已添加元器件")

            # 表格显示
            display_data = []
            total_amount = 0
            for i, comp in enumerate(st.session_state.components):
                amount = round(comp['unit_price'], 0) * comp['qty']
                total_amount += amount
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
            cols = st.columns(min(len(st.session_state.components), 10))
            for i, col in enumerate(cols):
                if i < len(st.session_state.components):
                    comp = st.session_state.components[i]
                    if col.button(f"🗑️ {comp['model'][:15]}", key=f"del_{i}"):
                        st.session_state.components.pop(i)
                        st.rerun()

            col_clear, col_calc = st.columns([1, 2])
            with col_clear:
                if st.button("🗑️ 清空清单"):
                    st.session_state.components = []
                    st.rerun()

            with col_calc:
                if st.button("📊 计算铜排成本", type="primary", use_container_width=True):
                    st.session_state.show_calc = True
                    st.rerun()

    with tab2:
        if not st.session_state.components:
            st.info("👈 请先在'元器件清单'中添加元器件并点击'计算铜排成本'")
        elif st.session_state.get('show_calc'):
            run_cost_analysis(st.session_state.components, copper_price, cabinet_width, copper_density, db)

    with tab3:
        show_instructions()

def run_cost_analysis(components: list, copper_price: float, cabinet_width: float,
                     copper_density: float, db: list):
    """执行成本分析"""
    st.header("📊 成本分析报告")

    cable_params = load_breaker_cable_params()

    # ─── 1. 元器件费用明细 ───
    st.subheader("1️⃣ 元器件费用")
    comp_data = []
    total_comp_cost = 0
    total_cable_cost = 0
    total_current = 0

    for comp in components:
        rounded_price = round(comp['unit_price'], 0)
        amount = rounded_price * comp['qty']
        total_comp_cost += amount

        # 计算电缆费用
        breaker_type = 'frame' if '框架' in comp.get('type', '') else 'mccb'
        cable = calc_cable_cost(breaker_type, comp['current'], comp['qty'],
                               cable_params, copper_price, cabinet_width)
        total_cable_cost += cable

        # 累计电流
        total_current += comp['current'] * comp['qty']

        comp_data.append({
            '名称': comp['name'],
            '型号': comp['model'],
            '电流': comp['current'],
            '数量': comp['qty'],
            '单价': rounded_price,
            '金额': amount,
            '电缆费': round(cable, 2),
        })

    df_comp = pd.DataFrame(comp_data)
    st.dataframe(df_comp, use_container_width=True, hide_index=True)

    st.metric("元器件总价", f"¥{total_comp_cost:,.0f}")
    st.metric("电缆费用合计", f"¥{total_cable_cost:,.0f}")

    # ─── 2. 铜排规格自动选择 ───
    st.divider()
    st.subheader("2️⃣ 铜排规格自动选择")

    # 计算总电流（含降容系数）
    reduced_current = total_current * 0.8 if total_current > 5000 else (total_current * 0.7 if total_current > 3000 else total_current)
    copper_spec = get_copper_spec_by_current(reduced_current)
    copper_threshold = get_copper_threshold_by_current(reduced_current)

    col1, col2, col3 = st.columns(3)
    col1.metric("总电流", f"{total_current:,.0f} A")
    col2.metric("降容后电流", f"{reduced_current:,.0f} A")
    col3.metric("铜排规格", f"TMY-{copper_spec * 100:.0f}×{copper_spec * 100:.0f}")

    # 铜排规格参考表
    st.caption("📋 铜排规格参考表（电流阈值 → 推荐铜排截面积）")
    spec_df = pd.DataFrame([
        {'电流范围 (A)': '0 ~ 500', '推荐铜排': 'TMY-18×18', '截面积 (cm²)': 0.18},
        {'电流范围 (A)': '500 ~ 755', '推荐铜排': 'TMY-24×24', '截面积 (cm²)': 0.24},
        {'电流范围 (A)': '755 ~ 990', '推荐铜排': 'TMY-36×36', '截面积 (cm²)': 0.36},
        {'电流范围 (A)': '990 ~ 1160', '推荐铜排': 'TMY-48×48', '截面积 (cm²)': 0.48},
        {'电流范围 (A)': '1160 ~ 1500', '推荐铜排': 'TMY-60×60', '截面积 (cm²)': 0.60},
        {'电流范围 (A)': '1500+', '推荐铜排': 'TMY-72×72', '截面积 (cm²)': 0.72},
    ])
    st.dataframe(spec_df, use_container_width=True, hide_index=True)

    # ─── 3. 铜排成本计算 ───
    st.divider()
    st.subheader("3️⃣ 铜排成本计算")

    copper_detail = calc_copper_busbar_cost(total_cable_cost, reduced_current, copper_price)

    col1, col2 = st.columns(2)
    with col1:
        st.metric("铜排总价", f"¥{copper_detail['total_cost']:,.0f}")
        st.metric("铜价", f"¥{copper_price}/单位")

    with col2:
        st.metric("三相铜排", f"¥{copper_detail['phase_cost']:,.0f}")
        st.metric("零线+地线", f"¥{copper_detail['neutral_cost'] + copper_detail['ground_cost']:,.0f}")

    # 详细分解
    with st.expander("📐 铜排成本分解"):
        st.write(f"""
        **计算公式:**
        ```
        铜排总价 = 三相铜排 + 零线 + 地线 + 电缆费用
        
        三相铜排 = 7 × {copper_threshold} × {copper_price} × {copper_density} = ¥{copper_detail['phase_cost']:,.0f}
        零线(N)  = 2 × ({copper_threshold}/2) × {copper_price} × {copper_density} = ¥{copper_detail['neutral_cost']:,.0f}
        地线(PE) = 2 × ({copper_threshold}/4) × {copper_price} × {copper_density} = ¥{copper_detail['ground_cost']:,.0f}
        电缆费用 = ¥{copper_detail['cable_cost']:,.0f}
        ─────────────────────────────────
        合计      = ¥{copper_detail['total_cost']:,.0f}
        ```
        
        - **7** = 柜内铜排总长度系数（米）
        - **{copper_threshold}** = 铜排电流阈值（对应截面积 {copper_detail['copper_spec']}）
        - **{copper_price}** = 铜价
        - **{copper_density}** = 铜密度 (g/cm³)
        """)

    # ─── 4. 利润计算 ───
    st.divider()
    st.subheader("4️⃣ 利润与报价")

    outgoing_circuits = sum(c['qty'] for c in components if '断路器' in c.get('type', ''))
    accessory_cost = calc_accessory_cost(outgoing_circuits)
    cabinet_cost = 2200 + copper_detail['cable_cost']  # 简化计算

    # 总成本
    total_cost = total_comp_cost + copper_detail['total_cost'] + accessory_cost + cabinet_cost

    # 各项利润
    comp_profit = calc_profit(total_comp_cost)
    copper_profit = calc_profit(copper_detail['total_cost'])
    accessory_profit = calc_profit(accessory_cost)
    cabinet_profit = calc_profit(cabinet_cost)
    total_profit = comp_profit + copper_profit + accessory_profit + cabinet_profit

    # 利润率
    profit_rate = (total_profit / total_cost * 100) if total_cost > 0 else 0

    col1, col2, col3 = st.columns(3)
    col1.metric("总成本", f"¥{total_cost:,.0f}")
    col2.metric("总利润", f"¥{total_profit:,.0f}")
    col3.metric("利润率", f"{profit_rate:.1f}%")

    # 费用明细
    with st.expander("💰 利润计算明细"):
        profit_data = pd.DataFrame([
            {'项目': '元器件', '成本': total_comp_cost, '利润': comp_profit,
             '利润率': f"{comp_profit/total_comp_cost*100:.1f}%" if total_comp_cost > 0 else "0%"},
            {'项目': '铜排', '成本': copper_detail['total_cost'], '利润': copper_profit,
             '利润率': f"{copper_profit/copper_detail['total_cost']*100:.1f}%" if copper_detail['total_cost'] > 0 else "0%"},
            {'项目': '辅助材料', '成本': accessory_cost, '利润': accessory_profit,
             '利润率': f"{accessory_profit/accessory_cost*100:.1f}%" if accessory_cost > 0 else "0%"},
            {'项目': '箱体', '成本': cabinet_cost, '利润': cabinet_profit,
             '利润率': f"{cabinet_profit/cabinet_cost*100:.1f}%" if cabinet_cost > 0 else "0%"},
            {'项目': '合计', '成本': total_cost, '利润': total_profit, '利润率': f"{profit_rate:.1f}%"},
        ])
        st.dataframe(profit_data, use_container_width=True, hide_index=True)

        st.caption("""
        **阶梯利润率:**
        - 金额 < ¥10,000: **30%**
        - ¥10,000 ~ ¥50,000: **29% - 0.03125% × 金额**（线性递减）
        - 金额 > ¥50,000: **10%**
        """)

    # ─── 5. 最终报价 ───
    st.divider()
    st.subheader("5️⃣ 最终报价")

    final_price = total_cost + total_profit
    tax_price = round(final_price * 1.13, 0)  # 含13%增值税

    col1, col2 = st.columns(2)
    col1.metric("不含税报价", f"¥{final_price:,.0f}")
    col2.metric("含税报价 (13%)", f"¥{tax_price:,.0f}")

    # 导出
    st.download_button(
        "📥 导出报价明细",
        data=generate_export_data(components, copper_detail, total_cost, total_profit, tax_price),
        file_name="成本报价明细.csv",
        mime="text/csv",
    )


def generate_export_data(components, copper_detail, total_cost, total_profit, tax_price):
    """生成CSV导出数据"""
    import io
    lines = ["序号,名称,型号,电流(A),数量,单价,金额,品牌"]
    for i, c in enumerate(components):
        amount = round(c['unit_price'], 0) * c['qty']
        lines.append(f"{i+1},{c['name']},{c['model']},{c['current']},{c['qty']},{round(c['unit_price'],0)},{amount},{c['brand']}")
    lines.append(f"\n铜排成本,{copper_detail['total_cost']}")
    lines.append(f"总成本,{total_cost}")
    lines.append(f"总利润,{total_profit}")
    lines.append(f"含税报价,{tax_price}")
    return "\n".join(lines)


def show_instructions():
    """使用说明"""
    st.header("📖 使用说明")

    st.markdown("""
    ## 🚀 快速开始

    1. **在左侧设置参数**: 铜价、柜宽
    2. **添加元器件**: 输入型号，系统自动从价格库匹配价格
    3. **点击"计算铜排成本"**: 自动计算铜排规格和费用
    4. **查看分析报告**: 铜排成本、利润、最终报价

    ## 📐 计算逻辑

    ### 铜排规格自动选择
    根据所有元器件的**总电流**自动选择合适的铜排截面积：
    | 总电流 | 推荐铜排 |
    |--------|----------|
    | < 500A | TMY-18×18 |
    | 500~755A | TMY-24×24 |
    | 755~990A | TMY-36×36 |
    | 990~1160A | TMY-48×48 |
    | 1160~1500A | TMY-60×60 |
    | > 1500A | TMY-72×72 |

    ### 铜排成本公式
    ```
    铜排总价 = 7×L×铜价×8.9 + 2×(L/2)×铜价×8.9 + 2×(L/4)×铜价×8.9 + 电缆费用
    ```
    - 7米 = 柜内铜排总长度
    - L = 铜排电流阈值
    - 8.9 = 铜密度(g/cm³)

    ### 利润计算
    - 金额 < ¥10,000 → 30%
    - ¥10,000 ~ ¥50,000 → 线性递减(29% → 10%)
    - 金额 > ¥50,000 → 10%

    ## ⚡ 价格库
    系统内置 **2,212** 条产品数据，包含：
    - ABB/西门子/施耐德断路器
    - 电流互感器
    - 数显仪表
    - 其他配电设备
    """)


if __name__ == "__main__":
    main()
