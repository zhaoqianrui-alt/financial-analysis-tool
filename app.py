import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import anthropic
from dotenv import load_dotenv
import os

load_dotenv()
API_KEY = os.getenv("ANTHROPIC_API_KEY")

# ============================================================
# 数据获取函数
# ============================================================

def is_a_share(code):
    return code.strip().isdigit()

def process_us_data(code):
    import yfinance as yf
    ticker = yf.Ticker(code.upper())
    income = ticker.financials
    balance = ticker.balance_sheet
    rows = []
    for date in income.columns[:5]:
        year = date.year
        try:
            revenue = income.loc["Total Revenue", date] / 1e8
            net_income = income.loc["Net Income", date] / 1e8
            total_assets = balance.loc["Total Assets", date] / 1e8
            equity = balance.loc["Stockholders Equity", date] / 1e8
            total_liabilities = total_assets - equity
            rows.append({
                "年份": year,
                "营业收入(亿美元)": round(revenue, 2),
                "净利润(亿美元)": round(net_income, 2),
                "总负债(亿美元)": round(total_liabilities, 2),
                "总资产(亿美元)": round(total_assets, 2),
            })
        except KeyError:
            continue
    return pd.DataFrame(rows).dropna().sort_values("年份").reset_index(drop=True)

def process_a_share_data(code):
    import akshare as ak
    df = ak.stock_financial_analysis_indicator(symbol=code)
    df = df[df["日期"].str.endswith("12-31")].copy()
    df["年份"] = df["日期"].str[:4].astype(int)
    return df.sort_values("年份").tail(5).reset_index(drop=True)

# ============================================================
# 预测建模函数
# 背景知识：利润表结构
# 营业收入
# - 营业成本
# = 毛利润
# - 销售费用、管理费用、研发费用（统称"期间费用"）
# = 营业利润（EBIT）
# - 所得税（税率 × 税前利润）
# = 净利润
# ============================================================

def build_income_statement(
    base_revenue,        # 基准年营业收入
    driver_names,        # 收入驱动因子名称列表，如["销量(万辆)", "单车价值(元)"]
    driver_base,         # 基准年各因子数值列表
    driver_forecasts,    # 未来3年各因子预测值，格式：[[Y1因子1, Y1因子2], [Y2...], [Y3...]]
    gross_margin,        # 毛利率假设（%）
    expense_ratio,       # 期间费用率假设（%，占收入比）
    tax_rate,            # 所得税率假设（%）
    base_year,           # 基准年份
    hist_df=None,        # 历史数据（可选，用于展示历史+预测对比）
    rev_col=None,        # 历史数据中收入列名
    net_col=None,        # 历史数据中净利润列名
):
    """
    构建预测利润表
    
    收入预测逻辑：
    - 如果只有1个因子：收入 = 基准收入 × (因子预测值 / 因子基准值)
    - 如果有2个因子：收入 = 因子1预测值 × 因子2预测值
      （适合"销量×单价"这类乘积型驱动）
    - 如果有3个因子：收入 = 因子1 × 因子2 × 因子3
    
    这是真实财务建模中最常用的收入预测方法
    """
    forecast_rows = []
    
    for i, year_drivers in enumerate(driver_forecasts):
        year = base_year + i + 1
        
        # 根据因子数量计算预测收入
        if len(driver_names) == 1:
            # 单因子：按比例缩放
            forecast_rev = base_revenue * (year_drivers[0] / driver_base[0])
        elif len(driver_names) == 2:
            # 双因子：两个因子相乘（如销量×单价）
            # 但单位可能不匹配，所以用"相对基准年的变化倍数"来算
            ratio = (year_drivers[0] / driver_base[0]) * (year_drivers[1] / driver_base[1])
            forecast_rev = base_revenue * ratio
        else:
            # 三因子
            ratio = (year_drivers[0] / driver_base[0]) * (year_drivers[1] / driver_base[1]) * (year_drivers[2] / driver_base[2])
            forecast_rev = base_revenue * ratio
        
        # 利润表推导
        # 背景知识：毛利润 = 收入 × 毛利率
        gross_profit = forecast_rev * gross_margin / 100
        # 期间费用 = 收入 × 费用率
        operating_expense = forecast_rev * expense_ratio / 100
        # 营业利润（EBIT）= 毛利润 - 期间费用
        ebit = gross_profit - operating_expense
        # 净利润 = 营业利润 × (1 - 税率)
        net_income = ebit * (1 - tax_rate / 100)
        # 净利率 = 净利润 / 收入
        net_margin = net_income / forecast_rev * 100
        
        forecast_rows.append({
            "年份": str(year),
            "营业收入": round(forecast_rev, 2),
            "毛利润": round(gross_profit, 2),
            "营业利润(EBIT)": round(ebit, 2),
            "净利润": round(net_income, 2),
            "毛利率%": round(gross_margin, 1),
            "净利率%": round(net_margin, 1),
            "类型": "预测",
        })
    
    forecast_df = pd.DataFrame(forecast_rows)
    
    # 如果有历史数据，拼接在一起展示
    if hist_df is not None and rev_col and net_col:
        hist_rows = []
        for _, row in hist_df.iterrows():
            rev = row[rev_col]
            net = row[net_col]
            if pd.notna(rev) and pd.notna(net) and rev != 0:
                hist_rows.append({
                    "年份": str(int(row["年份"])),
                    "营业收入": rev,
                    "毛利润": None,
                    "营业利润(EBIT)": None,
                    "净利润": net,
                    "毛利率%": None,
                    "净利率%": round(net / rev * 100, 1),
                    "类型": "历史",
                })
        if hist_rows:
            hist_display = pd.DataFrame(hist_rows)
            combined = pd.concat([hist_display, forecast_df], ignore_index=True)
            return forecast_df, combined
    
    return forecast_df, forecast_df


# ============================================================
# 页面主体
# ============================================================

st.title("📊 AI财务分析工具 v2")
st.write("输入股票代码，自动获取财务数据并生成分析")

col1, col2 = st.columns([2, 1])
with col1:
    stock_code = st.text_input(
        "股票代码",
        placeholder="A股填6位数字（如 600031），美股填字母（如 AAPL）"
    )
with col2:
    company_name = st.text_input("公司名称", placeholder="例如：拓普集团")

# 用 session_state 存储已获取的数据，避免每次交互都重新请求
# 背景知识：Streamlit每次用户点击任何按钮，整个页面代码都会重新运行一遍
# session_state 是一个"跨次运行保持数据"的机制，就像全局变量
if "fetched_df" not in st.session_state:
    st.session_state.fetched_df = None
if "fetched_company" not in st.session_state:
    st.session_state.fetched_company = None
if "is_us_stock" not in st.session_state:
    st.session_state.is_us_stock = False

if stock_code and company_name:
    if st.button("🔍 自动获取数据"):
        with st.spinner("正在从网络获取财务数据，请稍候..."):
            try:
                if is_a_share(stock_code):
                    st.info(f"✅ 识别为A股，正在获取 {company_name} 的数据...")
                    df = process_a_share_data(stock_code)
                    st.session_state.fetched_df = df
                    st.session_state.fetched_company = company_name
                    st.session_state.is_us_stock = False
                    st.subheader("📋 财务指标数据（年报）")
                    st.dataframe(df)
                    st.success("数据获取成功！")
                else:
                    st.info(f"✅ 识别为美股，正在从 Yahoo Finance 获取 {company_name} 的数据...")
                    df = process_us_data(stock_code)
                    st.session_state.fetched_df = df
                    st.session_state.fetched_company = company_name
                    st.session_state.is_us_stock = True

                    if df.empty:
                        st.warning("未获取到数据，请检查股票代码是否正确")
                    else:
                        rev_col = "营业收入(亿美元)"
                        net_col = "净利润(亿美元)"
                        liab_col = "总负债(亿美元)"
                        asset_col = "总资产(亿美元)"

                        df["净利率%"] = round(df[net_col] / df[rev_col] * 100, 2)
                        df["资产负债率%"] = round(df[liab_col] / df[asset_col] * 100, 2)
                        df["收入增速%"] = round(df[rev_col].pct_change() * 100, 2)

                        st.subheader("📋 财务数据")
                        st.dataframe(df)

                        st.subheader("📈 历史趋势")
                        chart_df = df[["年份", "净利率%", "资产负债率%"]].copy()
                        chart_df["年份"] = chart_df["年份"].astype(str)
                        fig = px.line(
                            chart_df.melt(id_vars="年份", var_name="指标", value_name="数值"),
                            x="年份", y="数值", color="指标",
                            markers=True, title=f"{company_name} 关键财务指标趋势"
                        )
                        fig.update_layout(hovermode="x unified", plot_bgcolor="white")
                        st.plotly_chart(fig, use_container_width=True)

                        st.subheader("🚨 风险预警")
                        latest = df.iloc[-1]
                        if latest["资产负债率%"] > 70:
                            st.error(f"⚠️ 资产负债率过高：{latest['资产负债率%']}%（警戒线70%）")
                        else:
                            st.success(f"✅ 资产负债率正常：{latest['资产负债率%']}%")
                        if latest["净利率%"] < 0:
                            st.error(f"🔴 净利率为负：{latest['净利率%']}%，公司出现亏损")
                        else:
                            st.success(f"✅ 净利率正常：{latest['净利率%']}%")

            except Exception as e:
                st.error(f"获取数据时出错：{e}")
                st.write("💡 请检查：1）股票代码是否正确  2）网络是否正常")

# ============================================================
# 预测建模模块
# 只有在已获取数据后才显示
# ============================================================

if st.session_state.fetched_df is not None and st.session_state.is_us_stock:
    st.divider()
    st.header("🔮 预测建模模块")
    st.write("基于你的核心判断，构建未来3年利润表预测模型")

    df = st.session_state.fetched_df
    cname = st.session_state.fetched_company
    rev_col = "营业收入(亿美元)"
    net_col = "净利润(亿美元)"

    # 取最新一年作为基准年
    latest_row = df.dropna(subset=[rev_col]).iloc[-1]
    base_year = int(latest_row["年份"])
    base_revenue = latest_row[rev_col]

    st.info(f"基准年：**{base_year}年**，营业收入：**{base_revenue} 亿美元**")

    # ---- 第一步：定义收入驱动因子 ----
    st.subheader("第一步：定义收入驱动因子")
    st.write("收入 = 因子1 × 因子2（×因子3）的乘积关系，例如：销量 × 单价，或用户数 × ARPU")

    num_drivers = st.radio("驱动因子数量", [1, 2, 3], index=1, horizontal=True)

    driver_names = []
    driver_bases = []
    for i in range(num_drivers):
        c1, c2 = st.columns([2, 1])
        with c1:
            name = st.text_input(
                f"因子{i+1}名称",
                value=["销量（万件）", "单价（美元）", "市场份额%"][i],
                key=f"driver_name_{i}"
            )
        with c2:
            base_val = st.number_input(
                f"因子{i+1}基准值（{base_year}年）",
                value=100.0,
                key=f"driver_base_{i}"
            )
        driver_names.append(name)
        driver_bases.append(base_val)

    # ---- 第二步：输入未来3年假设 ----
    st.subheader("第二步：输入未来3年预测假设")
    st.write("**收入驱动因子预测**")

    forecast_years = [base_year + 1, base_year + 2, base_year + 3]
    driver_forecasts = []

    # 用表格形式展示三年×多因子的输入
    # 背景知识：st.columns 可以创建多列，让输入框排成表格状
    year_cols = st.columns(3)
    year_driver_values = []

    for yi, year in enumerate(forecast_years):
        with year_cols[yi]:
            st.markdown(f"**{year}年**")
            year_vals = []
            for di in range(num_drivers):
                val = st.number_input(
                    driver_names[di],
                    value=float(driver_bases[di]),
                    key=f"forecast_{yi}_{di}"
                )
                year_vals.append(val)
            year_driver_values.append(year_vals)

    st.write("**利润表假设**")
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        gross_margin = st.number_input("毛利率%", value=40.0, min_value=0.0, max_value=100.0, step=0.5)
    with col_b:
        expense_ratio = st.number_input("期间费用率%", value=20.0, min_value=0.0, max_value=100.0, step=0.5,
                                         help="销售+管理+研发费用合计占收入的比例")
    with col_c:
        tax_rate = st.number_input("所得税率%", value=25.0, min_value=0.0, max_value=50.0, step=0.5)

    # ---- 第三步：生成预测结果 ----
    # 背景知识：用 session_state 存预测结果
    # 这样点"AI解读"时页面重跑，预测结果还在，不会消失
    if st.button("📊 生成预测利润表"):
        forecast_df, combined_df = build_income_statement(
            base_revenue=base_revenue,
            driver_names=driver_names,
            driver_base=driver_bases,
            driver_forecasts=year_driver_values,
            gross_margin=gross_margin,
            expense_ratio=expense_ratio,
            tax_rate=tax_rate,
            base_year=base_year,
            hist_df=df,
            rev_col=rev_col,
            net_col=net_col,
        )
        # 把预测结果和假设参数存入 session_state
        st.session_state.forecast_df = forecast_df
        st.session_state.combined_df = combined_df
        st.session_state.forecast_params = {
            "gross_margin": gross_margin,
            "expense_ratio": expense_ratio,
            "tax_rate": tax_rate,
            "cname": cname,
        }

    # 只要 session_state 里有预测结果，就显示（无论是刚生成还是之前生成的）
    if st.session_state.get("forecast_df") is not None:
        forecast_df = st.session_state.forecast_df
        combined_df = st.session_state.combined_df
        params = st.session_state.forecast_params

        st.subheader("📋 预测利润表（亿美元）")
        display_cols = ["年份", "营业收入", "毛利润", "营业利润(EBIT)", "净利润", "毛利率%", "净利率%"]
        st.dataframe(forecast_df[display_cols].set_index("年份"))

        # 历史+预测对比图
        st.subheader("📈 历史 vs 预测：收入与净利润")
        hist_data = combined_df[combined_df["类型"] == "历史"]
        forecast_data = combined_df[combined_df["类型"] == "预测"]

        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=hist_data["年份"], y=hist_data["营业收入"],
            name="历史收入", marker_color="#4472C4", opacity=0.8
        ))
        fig.add_trace(go.Bar(
            x=forecast_data["年份"], y=forecast_data["营业收入"],
            name="预测收入", marker_color="#4472C4", opacity=0.4,
            marker_pattern_shape="/"
        ))
        fig.add_trace(go.Scatter(
            x=hist_data["年份"], y=hist_data["净利润"],
            name="历史净利润", mode="lines+markers",
            line=dict(color="#ED7D31", width=2)
        ))
        fig.add_trace(go.Scatter(
            x=forecast_data["年份"], y=forecast_data["净利润"],
            name="预测净利润", mode="lines+markers",
            line=dict(color="#ED7D31", width=2, dash="dash")
        ))
        fig.update_layout(
            title=f"{params['cname']} 营业收入与净利润：历史 + 预测",
            barmode="group", hovermode="x unified",
            plot_bgcolor="white", yaxis_title="亿美元",
            legend=dict(orientation="h", yanchor="bottom", y=1.02)
        )
        st.plotly_chart(fig, use_container_width=True)

        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(
            x=hist_data["年份"], y=hist_data["净利率%"],
            name="历史净利率", mode="lines+markers",
            line=dict(color="#70AD47", width=2)
        ))
        fig2.add_trace(go.Scatter(
            x=forecast_data["年份"], y=forecast_data["净利率%"],
            name="预测净利率", mode="lines+markers",
            line=dict(color="#70AD47", width=2, dash="dash")
        ))
        fig2.update_layout(
            title="净利率趋势：历史 + 预测",
            plot_bgcolor="white", yaxis_title="%", hovermode="x unified"
        )
        st.plotly_chart(fig2, use_container_width=True)

        # AI解读按钮——现在和预测按钮同级，不嵌套，点击后页面重跑也能正常工作
        if st.button("🤖 AI解读预测结果"):
            with st.spinner("AI正在分析预测模型..."):
                forecast_text = ""
                for _, row in forecast_df.iterrows():
                    forecast_text += f"{row['年份']}年：收入{row['营业收入']}亿美元，净利润{row['净利润']}亿美元，净利率{row['净利率%']}%\n"

                prompt = f"""你是一名资深证券研究员。以下是{params['cname']}的财务预测数据：

{forecast_text}

预测假设：毛利率{params['gross_margin']}%，期间费用率{params['expense_ratio']}%，所得税率{params['tax_rate']}%

请从以下角度给出专业点评（500字左右）：
1. 这个预测情景是乐观、中性还是保守？
2. 毛利率和费用率假设是否合理？
3. 这个盈利能力水平在同行业中处于什么位置？
4. 主要的预测风险是什么？"""

                client = anthropic.Anthropic(api_key=API_KEY)
                message = client.messages.create(
                    model="claude-opus-4-6",
                    max_tokens=1024,
                    messages=[{"role": "user", "content": prompt}]
                )
                st.subheader("🤖 AI预测解读")
                st.write(message.content[0].text)

# ============================================================
# 三表勾稽模块
# 只有在预测利润表已生成后才显示
# ============================================================

if st.session_state.get("forecast_df") is not None:
    st.divider()
    st.header("📑 三表勾稽模型")
    st.write("在利润表预测的基础上，补全资产负债表和现金流量表，确保三表数据互相吻合")

    forecast_df = st.session_state.forecast_df
    params = st.session_state.forecast_params

    # ---- 三表假设输入 ----
    st.subheader("补充假设")
    st.write("以下假设用于推导资产负债表和现金流量表")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("**折旧与资本开支**")
        # 折旧率：固定资产每年折旧的比例
        # 背景知识：买了100亿设备，折旧率10%，每年在利润表计10亿折旧费用
        # 折旧是"非现金支出"——钱已经花了（买设备时），但每年分摊计入费用
        # 所以计算现金流时要把折旧加回来
        depreciation_rate = st.number_input(
            "折旧率%（占固定资产）",
            value=10.0, min_value=0.0, max_value=50.0, step=0.5,
            help="每年折旧金额 = 固定资产净值 × 折旧率"
        )
        # 资本开支率：每年新增资本开支占收入的比例
        # 背景知识：公司扩张需要买设备、建厂房，这笔钱叫资本开支(Capex)
        # 资本开支不直接进利润表，而是进资产负债表的固定资产，再通过折旧慢慢进利润表
        capex_rate = st.number_input(
            "资本开支率%（占收入）",
            value=5.0, min_value=0.0, max_value=50.0, step=0.5,
            help="每年资本开支 = 营业收入 × 资本开支率"
        )
    with col2:
        st.markdown("**营运资本（周转天数）**")
        # 背景知识：营运资本 = 应收账款 + 存货 - 应付账款
        # 用"天数"衡量更直观：应收账款天数=30天，意思是平均30天收回货款
        # 天数越长，占用现金越多，对公司越不利
        ar_days = st.number_input(
            "应收账款天数",
            value=60, min_value=0, max_value=365, step=5,
            help="平均多少天收回货款，天数越短越好"
        )
        inv_days = st.number_input(
            "存货天数",
            value=45, min_value=0, max_value=365, step=5,
            help="存货平均在仓库多少天"
        )
        ap_days = st.number_input(
            "应付账款天数",
            value=45, min_value=0, max_value=365, step=5,
            help="平均多少天付给供应商，天数越长对公司越有利"
        )
    with col3:
        st.markdown("**融资与分红**")
        # 分红率：净利润中分配给股东的比例
        dividend_payout = st.number_input(
            "分红率%（占净利润）",
            value=30.0, min_value=0.0, max_value=100.0, step=5.0,
            help="净利润中有多少比例以分红形式派发"
        )
        # 基准年资产负债表数据（用于推导预测期）
        st.markdown("**基准年资产负债表**")
        base_cash = st.number_input("基准年现金（亿美元）", value=50.0, step=1.0)
        base_fixed_assets = st.number_input("基准年固定资产净值（亿美元）", value=100.0, step=1.0)
        base_equity = st.number_input("基准年所有者权益（亿美元）", value=80.0, step=1.0)
        base_debt = st.number_input("基准年有息负债（亿美元）", value=50.0, step=1.0)

    if st.button("📊 生成三表模型"):

        # ---- 逐年推导三张报表 ----
        # 背景知识：三表必须"勾稽"，即互相之间数字要对得上
        # 任何一个地方逻辑错了，整个模型就失真

        income_rows = []      # 利润表汇总
        cashflow_rows = []    # 现金流量表
        balance_rows = []     # 资产负债表

        # 上一年的资产负债表数值（第一年用基准年数据）
        prev_cash = base_cash
        prev_fixed_assets = base_fixed_assets
        prev_equity = base_equity
        prev_debt = base_debt

        for _, row in forecast_df.iterrows():
            year = row["年份"]
            revenue = row["营业收入"]
            net_income = row["净利润"]
            ebit = row["营业利润(EBIT)"]

            # ==============================
            # 第一步：从利润表算出各项数值
            # ==============================
            capex = revenue * capex_rate / 100          # 资本开支
            depreciation = prev_fixed_assets * depreciation_rate / 100  # 折旧（基于上期固定资产）
            dividends = net_income * dividend_payout / 100  # 分红

            # 营运资本计算
            # 背景知识：
            # 应收账款 = 收入 × (天数/365)  →  卖出去但还没收到钱的部分
            # 存货 = 收入 × (天数/365)      →  还没卖出去的货
            # 应付账款 = 收入 × (天数/365)  →  买了原料但还没付钱的部分
            ar = revenue * ar_days / 365         # 应收账款
            inventory = revenue * inv_days / 365  # 存货
            ap = revenue * ap_days / 365          # 应付账款
            net_working_capital = ar + inventory - ap  # 净营运资本

            # ==============================
            # 第二步：构建现金流量表
            # 背景知识：现金流量表分三部分
            # 经营活动：日常经营产生/消耗的现金
            # 投资活动：买卖资产产生/消耗的现金（主要是资本开支）
            # 融资活动：借钱还钱、分红产生/消耗的现金
            # ==============================

            # 经营活动现金流
            # = 净利润（利润表底线）
            # + 折旧（加回，因为折旧不实际花钱）
            # - 营运资本增加（营运资本增加意味着更多现金被"锁"在应收/存货里）
            # 简化处理：用当期营运资本直接估算，不做期初期末差
            cfo = net_income + depreciation - (net_working_capital * 0.1)
            # 注：0.1是简化系数，实际应用中应用期末-期初的营运资本变化

            # 投资活动现金流（资本开支是现金流出，所以是负数）
            cfi = -capex

            # 融资活动现金流（分红是现金流出）
            # 假设有息负债维持不变（不新借也不还）
            cff = -dividends

            # 自由现金流 = 经营现金流 - 资本开支
            # 背景知识：自由现金流是公司"真正赚到手"的钱，是估值最重要的指标
            fcf = cfo + cfi

            # 现金净变化
            net_cash_change = cfo + cfi + cff

            # ==============================
            # 第三步：构建资产负债表
            # 背景知识：资产负债表是某一时点的"家底"快照
            # 必须满足：总资产 = 总负债 + 所有者权益
            # ==============================

            # 资产端
            end_cash = prev_cash + net_cash_change           # 期末现金
            end_fixed_assets = prev_fixed_assets + capex - depreciation  # 固定资产净值
            # 总资产 = 现金 + 固定资产 + 营运资本（应收+存货）
            total_assets = end_cash + end_fixed_assets + ar + inventory

            # 权益端
            # 留存收益增加 = 净利润 - 分红
            retained_earnings_increase = net_income - dividends
            end_equity = prev_equity + retained_earnings_increase

            # 负债端
            # 应付账款（流动负债的一部分）
            total_liabilities = ap + prev_debt  # 应付账款 + 有息负债

            # ==============================
            # 勾稽验证：总资产 = 总负债 + 所有者权益
            # 如果不等，说明模型有逻辑错误
            # ==============================
            balance_check = total_assets - (total_liabilities + end_equity)

            # 记录三张报表数据
            income_rows.append({
                "年份": year,
                "营业收入": round(revenue, 2),
                "折旧摊销": round(depreciation, 2),
                "营业利润(EBIT)": round(ebit, 2),
                "净利润": round(net_income, 2),
                "净利率%": round(row["净利率%"], 1),
            })

            cashflow_rows.append({
                "年份": year,
                "经营活动现金流(CFO)": round(cfo, 2),
                "资本开支(Capex)": round(-capex, 2),
                "投资活动现金流(CFI)": round(cfi, 2),
                "融资活动现金流(CFF)": round(cff, 2),
                "自由现金流(FCF)": round(fcf, 2),
                "现金净变化": round(net_cash_change, 2),
            })

            balance_rows.append({
                "年份": year,
                "现金": round(end_cash, 2),
                "应收账款": round(ar, 2),
                "存货": round(inventory, 2),
                "固定资产净值": round(end_fixed_assets, 2),
                "总资产": round(total_assets, 2),
                "应付账款": round(ap, 2),
                "有息负债": round(prev_debt, 2),
                "总负债": round(total_liabilities, 2),
                "所有者权益": round(end_equity, 2),
                "勾稽差异": round(balance_check, 2),  # 理想情况下应接近0
            })

            # 更新"上一年"数值，用于下一年计算
            prev_cash = end_cash
            prev_fixed_assets = end_fixed_assets
            prev_equity = end_equity

        # ---- 展示三张报表 ----
        income_display = pd.DataFrame(income_rows).set_index("年份")
        cashflow_display = pd.DataFrame(cashflow_rows).set_index("年份")
        balance_display = pd.DataFrame(balance_rows).set_index("年份")

        st.subheader("📋 利润表预测（亿美元）")
        st.dataframe(income_display)

        st.subheader("📋 现金流量表预测（亿美元）")
        st.dataframe(cashflow_display)

        st.subheader("📋 资产负债表预测（亿美元）")
        st.dataframe(balance_display)

        # 勾稽验证提示
        max_diff = balance_display["勾稽差异"].abs().max()
        if max_diff < 1:
            st.success(f"✅ 三表勾稽验证通过！最大差异：{round(max_diff, 3)} 亿美元（接近0，模型自洽）")
        else:
            st.warning(f"⚠️ 三表存在勾稽差异：{round(max_diff, 2)} 亿美元，请检查假设是否合理")

        # 自由现金流图表
        st.subheader("📈 自由现金流 vs 净利润")
        fig_cf = go.Figure()
        fig_cf.add_trace(go.Bar(
            x=cashflow_display.index,
            y=cashflow_display["自由现金流(FCF)"],
            name="自由现金流", marker_color="#70AD47"
        ))
        fig_cf.add_trace(go.Scatter(
            x=income_display.index,
            y=income_display["净利润"],
            name="净利润", mode="lines+markers",
            line=dict(color="#ED7D31", width=2)
        ))
        fig_cf.update_layout(
            title="自由现金流 vs 净利润对比（差距越小说明利润质量越高）",
            plot_bgcolor="white",
            hovermode="x unified",
            yaxis_title="亿美元",
            barmode="group"
        )
        st.plotly_chart(fig_cf, use_container_width=True)

        # 存储三表数据到session_state，供后续估值模块使用
        st.session_state.cashflow_df = pd.DataFrame(cashflow_rows)
        st.session_state.balance_df = pd.DataFrame(balance_rows)
        st.session_state.income_rows = income_rows
        st.session_state.cashflow_rows = cashflow_rows
        st.session_state.balance_rows = balance_rows
        st.session_state.three_table_params = {
            "depreciation_rate": depreciation_rate,
            "capex_rate": capex_rate,
            "ar_days": ar_days,
            "inv_days": inv_days,
            "ap_days": ap_days,
            "dividend_payout": dividend_payout,
        }

    # AI解读按钮——和"生成三表模型"同级，不嵌套
    if st.session_state.get("cashflow_rows") is not None:
        if st.button("🤖 AI解读三表质量"):
            income_rows = st.session_state.income_rows
            cashflow_rows = st.session_state.cashflow_rows
            balance_rows = st.session_state.balance_rows
            tp = st.session_state.three_table_params
            p = st.session_state.forecast_params

            with st.spinner("AI正在分析..."):
                prompt = f"""你是资深财务分析师。以下是{p['cname']}的预测三表关键数据：

利润表净利润：{[r['净利润'] for r in income_rows]}（亿美元）
自由现金流：{[r['自由现金流(FCF)'] for r in cashflow_rows]}（亿美元）
资产负债率：{[round(r['总负债']/r['总资产']*100,1) for r in balance_rows]}（%）

三表假设：折旧率{tp['depreciation_rate']}%，资本开支率{tp['capex_rate']}%，应收账款{tp['ar_days']}天，分红率{tp['dividend_payout']}%

请从专业角度分析（300字）：
1. 利润质量如何？自由现金流和净利润的差距说明什么？
2. 资本开支强度是否合理？对未来增长有何影响？
3. 营运资本假设是否符合行业惯例？
4. 这家公司的财务健康状况如何评价？"""

                client = anthropic.Anthropic(api_key=API_KEY)
                message = client.messages.create(
                    model="claude-opus-4-6",
                    max_tokens=1500,
                    messages=[{"role": "user", "content": prompt}]
                )
                st.subheader("🤖 AI三表质量解读")
                st.write(message.content[0].text)

# ============================================================
# 保留原有手动上传功能
# ============================================================
st.divider()
with st.expander("📁 手动上传Excel文件（备用方案）"):
    uploaded_file = st.file_uploader("上传Excel文件", type=["xlsx"])
    manual_company = st.text_input("公司名称", placeholder="例如：拓普集团", key="manual")

    if uploaded_file and manual_company:
        df_m = pd.read_excel(uploaded_file)
        df_m["净利率%"] = round(df_m["净利润"] / df_m["营业收入"] * 100, 2)
        df_m["资产负债率%"] = round(df_m["总负债"] / df_m["总资产"] * 100, 2)
        df_m["收入增速%"] = round(df_m["营业收入"].pct_change() * 100, 2)

        st.dataframe(df_m[["年份", "净利率%", "资产负债率%", "收入增速%"]])

        chart_df = df_m[["年份", "净利率%", "资产负债率%"]].copy()
        chart_df["年份"] = chart_df["年份"].astype(str)
        fig = px.line(
            chart_df.melt(id_vars="年份", var_name="指标", value_name="数值"),
            x="年份", y="数值", color="指标",
            markers=True, title=f"{manual_company} 财务趋势"
        )
        st.plotly_chart(fig, use_container_width=True)

        if st.button("生成AI分析报告", key="manual_ai"):
            with st.spinner("AI正在分析..."):
                data_text = ""
                for _, row in df_m.iterrows():
                    data_text += f"{int(row['年份'])}年：净利率{row['净利率%']}%，资产负债率{row['资产负债率%']}%，收入增速{row['收入增速%']}%\n"
                prompt = f"请作为证券研究员分析{manual_company}，200字：\n{data_text}\n分析盈利能力、财务风险和投资价值。"
                client = anthropic.Anthropic(api_key=API_KEY)
                message = client.messages.create(
                    model="claude-opus-4-6",
                    max_tokens=1024,
                    messages=[{"role": "user", "content": prompt}]
                )
                st.subheader("AI分析报告")
                st.write(message.content[0].text)