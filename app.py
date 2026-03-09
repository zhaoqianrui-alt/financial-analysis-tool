import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import anthropic
from dotenv import load_dotenv
import os

load_dotenv()
API_KEY = os.getenv("ANTHROPIC_API_KEY")
if not API_KEY:
    st.error("❗ ANTHROPIC_API_KEY not found. Please add it to your .env file.")
    st.stop()
client = anthropic.Anthropic(api_key=API_KEY)

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
# 页面主体 - 标签页布局（每个Tab独立可用）
# ============================================================

st.title("AI Financial Analysis Tool")
st.caption("Powered by Claude AI · Built for equity research")

# session_state 初始化
for key, default in {
    "fetched_df": None,
    "fetched_company": None,
    "is_us_stock": False,
    "forecast_df": None,
    "combined_df": None,
    "forecast_params": None,
    "cashflow_rows": None,
    "income_rows": None,
    "balance_rows": None,
    "three_table_params": None,
    "dcf_result": None,
    "dcf_params": None,
}.items():
    if key not in st.session_state:
        st.session_state[key] = default

tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Data",
    "📈 Forecast",
    "📑 3-Statement",
    "💰 Valuation",
])

# ============================================================
# TAB 1: DATA
# ============================================================
with tab1:
    st.header("Market Data")
    st.write("Enter a stock ticker to automatically fetch financial data.")

    col1, col2 = st.columns([2, 1])
    with col1:
        stock_code = st.text_input("Ticker Symbol",
            placeholder="A-share: 6-digit number (e.g. 600031) · US stock: letters (e.g. AAPL)")
    with col2:
        company_name = st.text_input("Company Name", placeholder="e.g. Apple")

    if stock_code and company_name:
        if st.button("🔍 Fetch Data"):
            with st.spinner("Fetching financial data..."):
                try:
                    if is_a_share(stock_code):
                        df = process_a_share_data(stock_code)
                        st.session_state.fetched_df = df
                        st.session_state.fetched_company = company_name
                        st.session_state.is_us_stock = False
                        st.success("Data fetched successfully!")
                    else:
                        df = process_us_data(stock_code)
                        st.session_state.fetched_df = df
                        st.session_state.fetched_company = company_name
                        st.session_state.is_us_stock = True
                        if df.empty:
                            st.warning("No data found. Please check the ticker symbol.")
                        else:
                            rev_col = "营业收入(亿美元)"
                            net_col = "净利润(亿美元)"
                            liab_col = "总负债(亿美元)"
                            asset_col = "总资产(亿美元)"
                            df["净利率%"] = round(df[net_col] / df[rev_col] * 100, 2)
                            df["资产负债率%"] = round(df[liab_col] / df[asset_col] * 100, 2)
                            df["收入增速%"] = round(df[rev_col].pct_change() * 100, 2)
                            st.session_state.fetched_df = df
                            st.success("Data fetched successfully!")
                except Exception as e:
                    st.error(f"Error: {e}")

    if st.session_state.fetched_df is not None and st.session_state.is_us_stock:
        df = st.session_state.fetched_df
        cname = st.session_state.fetched_company
        st.subheader(f"{cname} — Historical Financials (USD bn)")
        st.dataframe(df)
        chart_df = df[["年份", "净利率%", "资产负债率%"]].copy()
        chart_df["年份"] = chart_df["年份"].astype(str)
        fig = px.line(chart_df.melt(id_vars="年份", var_name="Metric", value_name="Value"),
            x="年份", y="Value", color="Metric", markers=True,
            title=f"{cname} — Net Margin % & Debt Ratio %")
        fig.update_layout(hovermode="x unified", plot_bgcolor="white")
        st.plotly_chart(fig, use_container_width=True)
        latest = df.iloc[-1]
        st.subheader("Risk Flags")
        if latest["资产负债率%"] > 70:
            st.error(f"⚠️ High leverage: {latest['资产负债率%']}%")
        else:
            st.success(f"✅ Leverage normal: {latest['资产负债率%']}%")
        if latest["净利率%"] < 0:
            st.error(f"🔴 Negative net margin: {latest['净利率%']}%")
        else:
            st.success(f"✅ Net margin positive: {latest['净利率%']}%")

    with st.expander("📁 Manual Upload (Excel fallback)"):
        uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
        manual_company = st.text_input("Company Name", placeholder="e.g. Tuopu Group", key="manual")
        if uploaded_file and manual_company:
            df_m = pd.read_excel(uploaded_file)
            df_m["净利率%"] = round(df_m["净利润"] / df_m["营业收入"] * 100, 2)
            df_m["资产负债率%"] = round(df_m["总负债"] / df_m["总资产"] * 100, 2)
            df_m["收入增速%"] = round(df_m["营业收入"].pct_change() * 100, 2)
            st.dataframe(df_m[["年份", "净利率%", "资产负债率%", "收入增速%"]])
            if st.button("Generate AI Report", key="manual_ai"):
                with st.spinner("Analyzing..."):
                    data_text = "".join(f"{int(r['年份'])}: Net margin {r['净利率%']}%, Debt ratio {r['资产负债率%']}%\n" for _, r in df_m.iterrows())
                    prompt = f"As a securities analyst, analyze {manual_company} in 200 words:\n{data_text}"
                    msg = client.messages.create(model="claude-opus-4-6", max_tokens=1024,
                        messages=[{"role": "user", "content": prompt}])
                    st.write(msg.content[0].text)

# ============================================================
# TAB 2: FORECAST
# ============================================================
with tab2:
    st.header("Income Statement Forecast")

    # ---- 数据来源：自动填入 or 手动输入 ----
    # 背景知识：如果 Data tab 已经获取了数据，就自动读取基准年数据
    # 如果没有，用户可以在下面的手动输入框直接填数字，一样能跑

    if st.session_state.fetched_df is not None and st.session_state.is_us_stock:
        df = st.session_state.fetched_df
        cname_f = st.session_state.fetched_company
        latest_row = df.dropna(subset=["营业收入(亿美元)"]).iloc[-1]
        auto_base_year = int(latest_row["年份"])
        auto_base_rev = float(latest_row["营业收入(亿美元)"])
        st.success(f"✅ Using data from Data tab: **{cname_f}**, base year **{auto_base_year}**, revenue **{auto_base_rev} bn**")
        use_auto = True
    else:
        use_auto = False

    with st.expander("📥 Manual Input — enter base year data directly", expanded=not use_auto):
        m_cname = st.text_input("Company name", value="My Company", key="f_cname")
        m_base_year = st.number_input("Base year", value=2024, min_value=2000, max_value=2030, step=1, key="f_year")
        m_base_rev = st.number_input("Base year revenue (USD bn)", value=100.0, min_value=0.1, step=1.0, key="f_rev")

    # 决定用哪个数据源
    if use_auto:
        f_cname = cname_f
        f_base_year = auto_base_year
        f_base_rev = auto_base_rev
        f_hist_df = df
    else:
        f_cname = m_cname
        f_base_year = int(m_base_year)
        f_base_rev = m_base_rev
        f_hist_df = None

    st.info(f"Base year: **{f_base_year}** · Revenue: **{f_base_rev} bn USD**")

    st.subheader("Step 1 — Revenue Drivers")
    num_drivers = st.radio("Number of drivers", [1, 2, 3], index=1, horizontal=True)
    driver_names, driver_bases = [], []
    for i in range(num_drivers):
        c1, c2 = st.columns([2, 1])
        with c1:
            name = st.text_input(f"Driver {i+1} name",
                value=["Units Sold (mn)", "ASP (USD)", "Market Share%"][i], key=f"driver_name_{i}")
        with c2:
            base_val = st.number_input(f"Base value ({f_base_year})", value=100.0, key=f"driver_base_{i}")
        driver_names.append(name)
        driver_bases.append(base_val)

    st.subheader("Step 2 — 3-Year Projections")
    forecast_years = [f_base_year + 1, f_base_year + 2, f_base_year + 3]
    year_cols = st.columns(3)
    year_driver_values = []
    for yi, year in enumerate(forecast_years):
        with year_cols[yi]:
            st.markdown(f"**{year}**")
            year_vals = []
            for di in range(num_drivers):
                val = st.number_input(driver_names[di], value=float(driver_bases[di]), key=f"forecast_{yi}_{di}")
                year_vals.append(val)
            year_driver_values.append(year_vals)

    col_a, col_b, col_c = st.columns(3)
    with col_a:
        gross_margin = st.number_input("Gross margin %", value=40.0, min_value=0.0, max_value=100.0, step=0.5)
    with col_b:
        expense_ratio = st.number_input("Opex ratio %", value=20.0, min_value=0.0, max_value=100.0, step=0.5)
    with col_c:
        tax_rate = st.number_input("Tax rate %", value=25.0, min_value=0.0, max_value=50.0, step=0.5)

    if st.button("📈 Generate Forecast"):
        forecast_df, combined_df = build_income_statement(
            base_revenue=f_base_rev, driver_names=driver_names,
            driver_base=driver_bases, driver_forecasts=year_driver_values,
            gross_margin=gross_margin, expense_ratio=expense_ratio,
            tax_rate=tax_rate, base_year=f_base_year,
            hist_df=f_hist_df,
            rev_col="营业收入(亿美元)" if f_hist_df is not None else None,
            net_col="净利润(亿美元)" if f_hist_df is not None else None,
        )
        st.session_state.forecast_df = forecast_df
        st.session_state.combined_df = combined_df
        st.session_state.forecast_params = {
            "gross_margin": gross_margin, "expense_ratio": expense_ratio,
            "tax_rate": tax_rate, "cname": f_cname,
        }

    if st.session_state.forecast_df is not None:
        forecast_df = st.session_state.forecast_df
        combined_df = st.session_state.combined_df
        params = st.session_state.forecast_params

        st.subheader("Projected Income Statement (USD bn)")
        st.dataframe(forecast_df[["年份","营业收入","毛利润","营业利润(EBIT)","净利润","毛利率%","净利率%"]].set_index("年份"))

        hist_data = combined_df[combined_df["类型"] == "历史"]
        forecast_data = combined_df[combined_df["类型"] == "预测"]
        fig = go.Figure()
        fig.add_trace(go.Bar(x=hist_data["年份"], y=hist_data["营业收入"], name="Historical Revenue", marker_color="#4472C4", opacity=0.8))
        fig.add_trace(go.Bar(x=forecast_data["年份"], y=forecast_data["营业收入"], name="Forecast Revenue", marker_color="#4472C4", opacity=0.4, marker_pattern_shape="/"))
        fig.add_trace(go.Scatter(x=hist_data["年份"], y=hist_data["净利润"], name="Historical Net Income", mode="lines+markers", line=dict(color="#ED7D31", width=2)))
        fig.add_trace(go.Scatter(x=forecast_data["年份"], y=forecast_data["净利润"], name="Forecast Net Income", mode="lines+markers", line=dict(color="#ED7D31", width=2, dash="dash")))
        fig.update_layout(title=f"{params['cname']} — Revenue & Net Income",
            barmode="group", hovermode="x unified", plot_bgcolor="white", yaxis_title="USD bn",
            legend=dict(orientation="h", yanchor="bottom", y=1.02))
        st.plotly_chart(fig, use_container_width=True)

        if st.button("🤖 AI Commentary"):
            with st.spinner("Analyzing..."):
                ft = "".join(f"{r['年份']}: Rev {r['营业收入']} bn, NI {r['净利润']} bn, Margin {r['净利率%']}%\n" for _, r in forecast_df.iterrows())
                prompt = f"""Senior equity analyst. Forecast for {params['cname']}:
{ft}
Assumptions: GM {params['gross_margin']}%, Opex {params['expense_ratio']}%, Tax {params['tax_rate']}%
Commentary (~300 words): 1) Bull/base/bear? 2) Margin assumptions reasonable? 3) Key risks?"""
                msg = client.messages.create(model="claude-opus-4-6", max_tokens=1024,
                    messages=[{"role": "user", "content": prompt}])
                st.subheader("🤖 AI Commentary")
                st.write(msg.content[0].text)

# ============================================================
# TAB 3: 3-STATEMENT
# ============================================================
with tab3:
    st.header("3-Statement Model")

    # ---- 数据来源：自动填入 or 手动输入 ----
    if st.session_state.forecast_df is not None:
        auto_forecast = st.session_state.forecast_df
        auto_params = st.session_state.forecast_params
        st.success(f"✅ Using forecast from Forecast tab: **{auto_params['cname']}**")
        use_auto_3s = True
    else:
        use_auto_3s = False

    with st.expander("📥 Manual Input — enter projected P&L directly", expanded=not use_auto_3s):
        st.write("Enter your 3-year net income and revenue forecasts:")
        m3_cname = st.text_input("Company name", value="My Company", key="3s_cname")
        cols_y = st.columns(3)
        manual_years = []
        manual_revenues = []
        manual_net_incomes = []
        manual_ebits = []
        manual_margins = []
        for i, col in enumerate(cols_y):
            with col:
                yr = st.number_input(f"Year {i+1}", value=2025+i, min_value=2020, max_value=2035, step=1, key=f"3s_year_{i}")
                rev = st.number_input("Revenue (bn)", value=100.0*(1+0.05*i), step=1.0, key=f"3s_rev_{i}")
                ni = st.number_input("Net Income (bn)", value=15.0*(1+0.05*i), step=0.5, key=f"3s_ni_{i}")
                manual_years.append(str(int(yr)))
                manual_revenues.append(rev)
                manual_net_incomes.append(ni)
                ebit = ni / 0.75  # 反推EBIT（假设税率25%）
                manual_ebits.append(round(ebit, 2))
                manual_margins.append(round(ni/rev*100, 1) if rev > 0 else 0)

    # 构建 forecast_df 供三表使用
    if use_auto_3s:
        s3_forecast_df = auto_forecast
        s3_cname = auto_params["cname"]
    else:
        s3_forecast_df = pd.DataFrame({
            "年份": manual_years,
            "营业收入": manual_revenues,
            "净利润": manual_net_incomes,
            "营业利润(EBIT)": manual_ebits,
            "净利率%": manual_margins,
            "毛利率%": [40.0]*3,
            "毛利润": [r*0.4 for r in manual_revenues],
            "类型": ["预测"]*3,
        })
        s3_cname = m3_cname

    st.subheader("Balance Sheet & Cash Flow Assumptions")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("**Depreciation & Capex**")
        depreciation_rate = st.number_input("D&A rate % (of PP&E)", value=10.0, min_value=0.0, max_value=50.0, step=0.5)
        capex_rate = st.number_input("Capex % (of revenue)", value=5.0, min_value=0.0, max_value=50.0, step=0.5)
    with col2:
        st.markdown("**Working Capital (Days)**")
        ar_days = st.number_input("DSO — Receivables days", value=60, min_value=0, max_value=365, step=5)
        inv_days = st.number_input("DIO — Inventory days", value=45, min_value=0, max_value=365, step=5)
        ap_days = st.number_input("DPO — Payables days", value=45, min_value=0, max_value=365, step=5)
    with col3:
        st.markdown("**Base Year Balance Sheet**")
        dividend_payout = st.number_input("Dividend payout %", value=30.0, min_value=0.0, max_value=100.0, step=5.0)
        base_cash = st.number_input("Cash (bn)", value=50.0, step=1.0)
        base_fixed_assets = st.number_input("PP&E net (bn)", value=100.0, step=1.0)
        base_equity = st.number_input("Equity (bn)", value=80.0, step=1.0)
        base_debt = st.number_input("Debt (bn)", value=50.0, step=1.0)

    if st.button("📑 Build 3-Statement Model"):
        income_rows, cashflow_rows, balance_rows = [], [], []
        prev_cash, prev_fa, prev_eq, prev_debt = base_cash, base_fixed_assets, base_equity, base_debt

        for _, row in s3_forecast_df.iterrows():
            year = row["年份"]
            revenue = row["营业收入"]
            net_income = row["净利润"]
            ebit = row["营业利润(EBIT)"]

            capex = revenue * capex_rate / 100
            depreciation = prev_fa * depreciation_rate / 100
            dividends = net_income * dividend_payout / 100
            ar = revenue * ar_days / 365
            inventory = revenue * inv_days / 365
            ap = revenue * ap_days / 365
            nwc = ar + inventory - ap
            cfo = net_income + depreciation - (nwc * 0.1)
            cfi = -capex
            cff = -dividends
            fcf = cfo + cfi
            net_cash_change = cfo + cfi + cff

            end_cash = prev_cash + net_cash_change
            end_fa = prev_fa + capex - depreciation
            total_assets = end_cash + end_fa + ar + inventory
            end_equity = prev_eq + (net_income - dividends)
            total_liabilities = ap + prev_debt
            check = total_assets - (total_liabilities + end_equity)

            income_rows.append({"Year": year, "Revenue": round(revenue,2), "D&A": round(depreciation,2), "EBIT": round(ebit,2), "Net Income": round(net_income,2), "Net Margin%": round(row["净利率%"],1)})
            cashflow_rows.append({"Year": year, "CFO": round(cfo,2), "Capex": round(-capex,2), "CFI": round(cfi,2), "CFF": round(cff,2), "FCF": round(fcf,2), "Net Change": round(net_cash_change,2)})
            balance_rows.append({"Year": year, "Cash": round(end_cash,2), "Receivables": round(ar,2), "Inventory": round(inventory,2), "PP&E": round(end_fa,2), "Total Assets": round(total_assets,2), "Payables": round(ap,2), "Debt": round(prev_debt,2), "Total Liabilities": round(total_liabilities,2), "Equity": round(end_equity,2), "Check": round(check,2)})

            prev_cash, prev_fa, prev_eq = end_cash, end_fa, end_equity

        st.session_state.income_rows = income_rows
        st.session_state.cashflow_rows = cashflow_rows
        st.session_state.balance_rows = balance_rows
        st.session_state.three_table_params = {"depreciation_rate": depreciation_rate, "capex_rate": capex_rate, "ar_days": ar_days, "inv_days": inv_days, "ap_days": ap_days, "dividend_payout": dividend_payout, "cname": s3_cname}

    if st.session_state.cashflow_rows is not None:
        income_rows = st.session_state.income_rows
        cashflow_rows = st.session_state.cashflow_rows
        balance_rows = st.session_state.balance_rows

        st.subheader("Income Statement")
        st.dataframe(pd.DataFrame(income_rows).set_index("Year"))
        st.subheader("Cash Flow Statement")
        st.dataframe(pd.DataFrame(cashflow_rows).set_index("Year"))
        st.subheader("Balance Sheet")
        bd = pd.DataFrame(balance_rows).set_index("Year")
        st.dataframe(bd)

        max_diff = bd["Check"].abs().max()
        if max_diff < 1:
            st.success(f"✅ Balance sheet checks out. Max discrepancy: {round(max_diff,3)} bn")
        else:
            st.warning(f"⚠️ Discrepancy: {round(max_diff,2)} bn")

        fig_cf = go.Figure()
        fig_cf.add_trace(go.Bar(x=[r["Year"] for r in cashflow_rows], y=[r["FCF"] for r in cashflow_rows], name="FCF", marker_color="#70AD47"))
        fig_cf.add_trace(go.Scatter(x=[r["Year"] for r in income_rows], y=[r["Net Income"] for r in income_rows], name="Net Income", mode="lines+markers", line=dict(color="#ED7D31", width=2)))
        fig_cf.update_layout(title="FCF vs Net Income", plot_bgcolor="white", hovermode="x unified", yaxis_title="bn")
        st.plotly_chart(fig_cf, use_container_width=True)

        if st.button("🤖 AI 3-Statement Commentary"):
            with st.spinner("Analyzing..."):
                tp = st.session_state.three_table_params
                prompt = f"""Senior financial analyst. 3-statement model for {tp['cname']}:
Net Income: {[r['Net Income'] for r in income_rows]} bn
FCF: {[r['FCF'] for r in cashflow_rows]} bn
Debt ratio: {[round(r['Total Liabilities']/r['Total Assets']*100,1) for r in balance_rows]}%
Assumptions: D&A {tp['depreciation_rate']}%, Capex {tp['capex_rate']}%, DSO {tp['ar_days']}d, DPO {tp['ap_days']}d
Analysis (~300 words): 1) Earnings quality? 2) Capex intensity? 3) Working capital norms? 4) Financial health?"""
                msg = client.messages.create(model="claude-opus-4-6", max_tokens=1500,
                    messages=[{"role": "user", "content": prompt}])
                st.subheader("🤖 AI Commentary")
                st.write(msg.content[0].text)

# ============================================================
# TAB 4: VALUATION
# ============================================================
with tab4:
    st.header("DCF Valuation")

    # ---- 数据来源：自动填入 or 手动输入 ----
    if st.session_state.cashflow_rows is not None:
        auto_fcf = [r["FCF"] for r in st.session_state.cashflow_rows]
        auto_years = [r["Year"] for r in st.session_state.cashflow_rows]
        auto_vcname = st.session_state.three_table_params.get("cname", "Company")
        st.success(f"✅ Using FCF from 3-Statement tab: **{auto_vcname}** · FCF: {auto_fcf} bn")
        use_auto_v = True
    else:
        use_auto_v = False

    with st.expander("📥 Manual Input — enter FCF directly", expanded=not use_auto_v):
        v_cname = st.text_input("Company name", value="My Company", key="v_cname")
        vcols = st.columns(3)
        manual_fcf = []
        manual_vyears = []
        for i, col in enumerate(vcols):
            with col:
                yr = st.number_input(f"Year {i+1}", value=2025+i, min_value=2020, max_value=2035, step=1, key=f"v_year_{i}")
                fcf_val = st.number_input(f"FCF (bn)", value=50.0 + i*5, step=1.0, key=f"v_fcf_{i}")
                manual_fcf.append(fcf_val)
                manual_vyears.append(str(int(yr)))

    # 决定用哪个数据源
    if use_auto_v:
        v_fcf_list = auto_fcf
        v_years = auto_years
        v_company = auto_vcname
    else:
        v_fcf_list = manual_fcf
        v_years = manual_vyears
        v_company = v_cname

    st.info(f"FCF inputs: **{v_years[0]}: {v_fcf_list[0]} bn** · **{v_years[1]}: {v_fcf_list[1]} bn** · **{v_years[2]}: {v_fcf_list[2]} bn**")

    st.subheader("DCF Assumptions")
    col1, col2, col3 = st.columns(3)
    with col1:
        wacc = st.number_input("WACC %", value=9.0, min_value=1.0, max_value=30.0, step=0.5,
            help="Weighted avg cost of capital. Tech: typically 8–12%")
    with col2:
        terminal_growth = st.number_input("Terminal growth rate %", value=3.0, min_value=0.0, max_value=10.0, step=0.5,
            help="Perpetual growth after Year 3. Typically ~GDP (2–4%)")
    with col3:
        shares_outstanding = st.number_input("Shares outstanding (bn)", value=152.0, min_value=0.1, step=1.0)
        net_debt = st.number_input("Net debt (bn)", value=310.0, step=10.0,
            help="Debt minus cash. Negative = net cash")

    st.subheader("Sensitivity Analysis Range")
    col_a, col_b = st.columns(2)
    with col_a:
        wacc_low = st.number_input("WACC low %", value=7.0, min_value=1.0, max_value=29.0, step=0.5)
        wacc_high = st.number_input("WACC high %", value=12.0, min_value=2.0, max_value=30.0, step=0.5)
    with col_b:
        tg_low = st.number_input("Terminal growth low %", value=1.0, min_value=0.0, max_value=9.0, step=0.5)
        tg_high = st.number_input("Terminal growth high %", value=5.0, min_value=0.5, max_value=10.0, step=0.5)

    if st.button("💰 Run DCF Valuation"):
        import numpy as np

        def dcf_valuation(fcf_list, wacc_pct, tg_pct, shares, nd):
            w, g = wacc_pct/100, tg_pct/100
            pv_fcf = sum(f/(1+w)**(i+1) for i,f in enumerate(fcf_list))
            tv = fcf_list[-1]*(1+g)/(w-g)
            pv_tv = tv/(1+w)**len(fcf_list)
            ev = pv_fcf + pv_tv
            eq = ev - nd
            pps = eq/shares if shares > 0 else 0
            return {"pv_fcf": round(pv_fcf,1), "pv_terminal": round(pv_tv,1),
                "ev": round(ev,1), "equity_value": round(eq,1),
                "price_per_share": round(pps,2),
                "terminal_pct": round(pv_tv/ev*100,1) if ev > 0 else 0}

        base = dcf_valuation(v_fcf_list, wacc, terminal_growth, shares_outstanding, net_debt)

        c1,c2,c3,c4 = st.columns(4)
        c1.metric("PV of FCFs", f"{base['pv_fcf']} bn")
        c2.metric("PV of Terminal Value", f"{base['pv_terminal']} bn", f"{base['terminal_pct']}% of EV")
        c3.metric("Enterprise Value", f"{base['ev']} bn")
        c4.metric("Intrinsic Value / Share", f"${base['price_per_share']}")

        if base["terminal_pct"] > 80:
            st.warning(f"⚠️ Terminal value is {base['terminal_pct']}% of EV — highly sensitive to long-term assumptions")
        else:
            st.success(f"✅ Terminal value is {base['terminal_pct']}% of EV")

        wacc_steps = np.arange(wacc_low, wacc_high+0.1, 1.0)
        tg_steps = np.arange(tg_low, tg_high+0.1, 1.0)
        matrix_data, all_prices = {}, []
        for w in wacc_steps:
            row_data = {}
            for g in tg_steps:
                if w <= g:
                    row_data[f"g={g:.0f}%"] = "N/A"
                else:
                    r = dcf_valuation(v_fcf_list, w, g, shares_outstanding, net_debt)
                    row_data[f"g={g:.0f}%"] = f"${r['price_per_share']}"
                    all_prices.append(r["price_per_share"])
            matrix_data[f"WACC={w:.0f}%"] = row_data

        st.subheader("Sensitivity Matrix — Intrinsic Value per Share ($)")
        st.dataframe(pd.DataFrame(matrix_data).T, use_container_width=True)

        if all_prices:
            fig_val = go.Figure()
            fig_val.add_trace(go.Box(y=all_prices, name="Valuation range",
                marker_color="#4472C4", boxpoints="all", jitter=0.3))
            fig_val.add_hline(y=base["price_per_share"], line_dash="dash", line_color="red",
                annotation_text=f"Base ${base['price_per_share']}")
            fig_val.update_layout(title="Intrinsic Value Distribution",
                plot_bgcolor="white", yaxis_title="USD per share", showlegend=False)
            st.plotly_chart(fig_val, use_container_width=True)
            low, high, mid = round(min(all_prices),2), round(max(all_prices),2), round(sum(all_prices)/len(all_prices),2)
            st.info(f"📌 Range: **${low} — ${high}** · Median: **${mid}**")

        st.session_state.dcf_result = base
        st.session_state.dcf_params = {"wacc": wacc, "terminal_growth": terminal_growth,
            "shares": shares_outstanding, "net_debt": net_debt,
            "price_low": min(all_prices) if all_prices else 0,
            "price_high": max(all_prices) if all_prices else 0,
            "cname": v_company}

    if st.session_state.dcf_result is not None:
        if st.button("🤖 AI Valuation Report"):
            with st.spinner("Generating report..."):
                base = st.session_state.dcf_result
                dp = st.session_state.dcf_params
                prompt = f"""Senior equity research analyst. Valuation report for {dp['cname']}.
DCF: base case ${base['price_per_share']}/share, EV {base['ev']} bn, terminal {base['terminal_pct']}% of EV
Range: ${dp['price_low']:.1f}–${dp['price_high']:.1f}
WACC {dp['wacc']}%, terminal growth {dp['terminal_growth']}%
FCF forecast: {v_fcf_list} bn
Report (~400 words): 1) Under/overvalued? 2) Assumption sensitivity 3) Risks 4) Buy/Hold/Sell + target price"""
                msg = client.messages.create(model="claude-opus-4-6", max_tokens=2000,
                    messages=[{"role": "user", "content": prompt}])
                st.subheader("🤖 AI Valuation Report")
                st.write(msg.content[0].text)
