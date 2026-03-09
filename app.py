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
# 页面主体 - 标签页布局
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
    "wacc_range": (7.0, 12.0),
    "tg_range": (1.0, 5.0),
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
        stock_code = st.text_input(
            "Ticker Symbol",
            placeholder="A-share: 6-digit number (e.g. 600031) · US stock: letters (e.g. AAPL)"
        )
    with col2:
        company_name = st.text_input("Company Name", placeholder="e.g. Apple")

    if stock_code and company_name:
        if st.button("🔍 Fetch Data", key="fetch_btn"):
            with st.spinner("Fetching financial data..."):
                try:
                    if is_a_share(stock_code):
                        st.info(f"✅ A-share detected. Fetching {company_name}...")
                        df = process_a_share_data(stock_code)
                        st.session_state.fetched_df = df
                        st.session_state.fetched_company = company_name
                        st.session_state.is_us_stock = False
                        st.subheader("Financial Indicators (Annual)")
                        st.dataframe(df)
                        st.success("Data fetched successfully!")
                    else:
                        st.info(f"✅ US stock detected. Fetching {company_name} from Yahoo Finance...")
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

                except Exception as e:
                    st.error(f"Error fetching data: {e}")
                    st.write("💡 Check: 1) Ticker symbol correct  2) Network connection")

    # 展示已获取的数据（每次进tab都显示）
    if st.session_state.fetched_df is not None and st.session_state.is_us_stock:
        df = st.session_state.fetched_df
        cname = st.session_state.fetched_company
        rev_col = "营业收入(亿美元)"
        net_col = "净利润(亿美元)"
        liab_col = "总负债(亿美元)"
        asset_col = "总资产(亿美元)"

        st.subheader(f"{cname} — Historical Financials (USD bn)")
        st.dataframe(df)

        st.subheader("Key Ratios Trend")
        chart_df = df[["年份", "净利率%", "资产负债率%"]].copy()
        chart_df["年份"] = chart_df["年份"].astype(str)
        fig = px.line(
            chart_df.melt(id_vars="年份", var_name="Metric", value_name="Value"),
            x="年份", y="Value", color="Metric",
            markers=True, title=f"{cname} — Net Margin % & Debt Ratio %"
        )
        fig.update_layout(hovermode="x unified", plot_bgcolor="white")
        st.plotly_chart(fig, use_container_width=True)

        st.subheader("Risk Flags")
        latest = df.iloc[-1]
        if latest["资产负债率%"] > 70:
            st.error(f"⚠️ High leverage: {latest['资产负债率%']}% (threshold: 70%)")
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
                    data_text = ""
                    for _, row in df_m.iterrows():
                        data_text += f"{int(row['年份'])}: Net margin {row['净利率%']}%, Debt ratio {row['资产负债率%']}%, Revenue growth {row['收入增速%']}%\n"
                    prompt = f"As a securities analyst, analyze {manual_company} in 200 words:\n{data_text}\nCover profitability, financial risk, and investment value."
                    client = anthropic.Anthropic(api_key=API_KEY)
                    message = client.messages.create(
                        model="claude-opus-4-6", max_tokens=1024,
                        messages=[{"role": "user", "content": prompt}]
                    )
                    st.subheader("AI Analysis Report")
                    st.write(message.content[0].text)

# ============================================================
# TAB 2: FORECAST
# ============================================================
with tab2:
    st.header("Income Statement Forecast")
    st.write("Define revenue drivers and build a 3-year income statement projection.")

    if st.session_state.fetched_df is None or not st.session_state.is_us_stock:
        st.info("👈 Please fetch US stock data in the **Data** tab first.")
    else:
        df = st.session_state.fetched_df
        cname = st.session_state.fetched_company
        rev_col = "营业收入(亿美元)"
        net_col = "净利润(亿美元)"

        latest_row = df.dropna(subset=[rev_col]).iloc[-1]
        base_year = int(latest_row["年份"])
        base_revenue = latest_row[rev_col]

        st.info(f"Base year: **{base_year}** · Revenue: **{base_revenue} bn USD**")

        st.subheader("Step 1 — Revenue Drivers")
        st.write("Revenue = Driver 1 × Driver 2 (× Driver 3), e.g. Units Sold × ASP")
        num_drivers = st.radio("Number of drivers", [1, 2, 3], index=1, horizontal=True)

        driver_names, driver_bases = [], []
        for i in range(num_drivers):
            c1, c2 = st.columns([2, 1])
            with c1:
                name = st.text_input(f"Driver {i+1} name",
                    value=["Units Sold (mn)", "ASP (USD)", "Market Share%"][i],
                    key=f"driver_name_{i}")
            with c2:
                base_val = st.number_input(f"Driver {i+1} base ({base_year})",
                    value=100.0, key=f"driver_base_{i}")
            driver_names.append(name)
            driver_bases.append(base_val)

        st.subheader("Step 2 — 3-Year Projections")
        st.write("**Revenue driver forecasts**")
        forecast_years = [base_year + 1, base_year + 2, base_year + 3]
        year_cols = st.columns(3)
        year_driver_values = []
        for yi, year in enumerate(forecast_years):
            with year_cols[yi]:
                st.markdown(f"**{year}**")
                year_vals = []
                for di in range(num_drivers):
                    val = st.number_input(driver_names[di],
                        value=float(driver_bases[di]), key=f"forecast_{yi}_{di}")
                    year_vals.append(val)
                year_driver_values.append(year_vals)

        st.write("**P&L assumptions**")
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            gross_margin = st.number_input("Gross margin %", value=40.0, min_value=0.0, max_value=100.0, step=0.5)
        with col_b:
            expense_ratio = st.number_input("Opex ratio %", value=20.0, min_value=0.0, max_value=100.0, step=0.5,
                help="SG&A + R&D as % of revenue")
        with col_c:
            tax_rate = st.number_input("Tax rate %", value=25.0, min_value=0.0, max_value=50.0, step=0.5)

        if st.button("📈 Generate Income Statement Forecast"):
            forecast_df, combined_df = build_income_statement(
                base_revenue=base_revenue, driver_names=driver_names,
                driver_base=driver_bases, driver_forecasts=year_driver_values,
                gross_margin=gross_margin, expense_ratio=expense_ratio,
                tax_rate=tax_rate, base_year=base_year,
                hist_df=df, rev_col=rev_col, net_col=net_col,
            )
            st.session_state.forecast_df = forecast_df
            st.session_state.combined_df = combined_df
            st.session_state.forecast_params = {
                "gross_margin": gross_margin, "expense_ratio": expense_ratio,
                "tax_rate": tax_rate, "cname": cname,
            }

        if st.session_state.forecast_df is not None:
            forecast_df = st.session_state.forecast_df
            combined_df = st.session_state.combined_df
            params = st.session_state.forecast_params

            st.subheader("Projected Income Statement (USD bn)")
            display_cols = ["年份", "营业收入", "毛利润", "营业利润(EBIT)", "净利润", "毛利率%", "净利率%"]
            st.dataframe(forecast_df[display_cols].set_index("年份"))

            st.subheader("Historical vs Forecast")
            hist_data = combined_df[combined_df["类型"] == "历史"]
            forecast_data = combined_df[combined_df["类型"] == "预测"]

            fig = go.Figure()
            fig.add_trace(go.Bar(x=hist_data["年份"], y=hist_data["营业收入"], name="Historical Revenue", marker_color="#4472C4", opacity=0.8))
            fig.add_trace(go.Bar(x=forecast_data["年份"], y=forecast_data["营业收入"], name="Forecast Revenue", marker_color="#4472C4", opacity=0.4, marker_pattern_shape="/"))
            fig.add_trace(go.Scatter(x=hist_data["年份"], y=hist_data["净利润"], name="Historical Net Income", mode="lines+markers", line=dict(color="#ED7D31", width=2)))
            fig.add_trace(go.Scatter(x=forecast_data["年份"], y=forecast_data["净利润"], name="Forecast Net Income", mode="lines+markers", line=dict(color="#ED7D31", width=2, dash="dash")))
            fig.update_layout(title=f"{params['cname']} — Revenue & Net Income", barmode="group",
                hovermode="x unified", plot_bgcolor="white", yaxis_title="USD bn",
                legend=dict(orientation="h", yanchor="bottom", y=1.02))
            st.plotly_chart(fig, use_container_width=True)

            fig2 = go.Figure()
            fig2.add_trace(go.Scatter(x=hist_data["年份"], y=hist_data["净利率%"], name="Historical", mode="lines+markers", line=dict(color="#70AD47", width=2)))
            fig2.add_trace(go.Scatter(x=forecast_data["年份"], y=forecast_data["净利率%"], name="Forecast", mode="lines+markers", line=dict(color="#70AD47", width=2, dash="dash")))
            fig2.update_layout(title="Net Margin % — Historical vs Forecast", plot_bgcolor="white", yaxis_title="%", hovermode="x unified")
            st.plotly_chart(fig2, use_container_width=True)

            if st.button("🤖 AI Commentary on Forecast"):
                with st.spinner("Analyzing..."):
                    forecast_text = "".join(f"{row['年份']}: Revenue {row['营业收入']} bn, Net Income {row['净利润']} bn, Net margin {row['净利率%']}%\n" for _, row in forecast_df.iterrows())
                    prompt = f"""You are a senior equity research analyst. Here is the financial forecast for {params['cname']}:

{forecast_text}

Assumptions: Gross margin {params['gross_margin']}%, Opex ratio {params['expense_ratio']}%, Tax rate {params['tax_rate']}%

Provide a professional commentary (~300 words):
1. Is this scenario bullish, base, or bearish?
2. Are the margin assumptions reasonable?
3. How does this profitability compare to industry peers?
4. What are the key forecast risks?"""
                    client = anthropic.Anthropic(api_key=API_KEY)
                    message = client.messages.create(model="claude-opus-4-6", max_tokens=1024,
                        messages=[{"role": "user", "content": prompt}])
                    st.subheader("🤖 AI Forecast Commentary")
                    st.write(message.content[0].text)

# ============================================================
# TAB 3: 3-STATEMENT
# ============================================================
with tab3:
    st.header("3-Statement Model")
    st.write("Extend the income statement forecast to a full balance sheet and cash flow statement.")

    if st.session_state.forecast_df is None:
        st.info("👈 Please complete the income statement forecast in the **Forecast** tab first.")
    else:
        forecast_df = st.session_state.forecast_df
        params = st.session_state.forecast_params

        st.subheader("Additional Assumptions")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown("**Depreciation & Capex**")
            depreciation_rate = st.number_input("Depreciation rate % (of PP&E)", value=10.0, min_value=0.0, max_value=50.0, step=0.5, help="Annual depreciation = PP&E × rate")
            capex_rate = st.number_input("Capex rate % (of revenue)", value=5.0, min_value=0.0, max_value=50.0, step=0.5)
        with col2:
            st.markdown("**Working Capital (Days)**")
            ar_days = st.number_input("Receivables days (DSO)", value=60, min_value=0, max_value=365, step=5)
            inv_days = st.number_input("Inventory days (DIO)", value=45, min_value=0, max_value=365, step=5)
            ap_days = st.number_input("Payables days (DPO)", value=45, min_value=0, max_value=365, step=5)
        with col3:
            st.markdown("**Financing & Dividends**")
            dividend_payout = st.number_input("Dividend payout % (of net income)", value=30.0, min_value=0.0, max_value=100.0, step=5.0)
            st.markdown("**Base Year Balance Sheet**")
            base_cash = st.number_input("Cash (USD bn)", value=50.0, step=1.0)
            base_fixed_assets = st.number_input("PP&E net (USD bn)", value=100.0, step=1.0)
            base_equity = st.number_input("Shareholders' equity (USD bn)", value=80.0, step=1.0)
            base_debt = st.number_input("Interest-bearing debt (USD bn)", value=50.0, step=1.0)

        if st.button("📑 Build 3-Statement Model"):
            income_rows, cashflow_rows, balance_rows = [], [], []
            prev_cash, prev_fixed_assets_val, prev_equity, prev_debt = base_cash, base_fixed_assets, base_equity, base_debt

            for _, row in forecast_df.iterrows():
                year = row["年份"]
                revenue = row["营业收入"]
                net_income = row["净利润"]
                ebit = row["营业利润(EBIT)"]

                capex = revenue * capex_rate / 100
                depreciation = prev_fixed_assets_val * depreciation_rate / 100
                dividends = net_income * dividend_payout / 100
                ar = revenue * ar_days / 365
                inventory = revenue * inv_days / 365
                ap = revenue * ap_days / 365
                net_working_capital = ar + inventory - ap
                cfo = net_income + depreciation - (net_working_capital * 0.1)
                cfi = -capex
                cff = -dividends
                fcf = cfo + cfi
                net_cash_change = cfo + cfi + cff

                end_cash = prev_cash + net_cash_change
                end_fixed_assets = prev_fixed_assets_val + capex - depreciation
                total_assets = end_cash + end_fixed_assets + ar + inventory
                end_equity = prev_equity + (net_income - dividends)
                total_liabilities = ap + prev_debt
                balance_check = total_assets - (total_liabilities + end_equity)

                income_rows.append({"Year": year, "Revenue": round(revenue, 2), "D&A": round(depreciation, 2), "EBIT": round(ebit, 2), "Net Income": round(net_income, 2), "Net Margin%": round(row["净利率%"], 1)})
                cashflow_rows.append({"Year": year, "CFO": round(cfo, 2), "Capex": round(-capex, 2), "CFI": round(cfi, 2), "CFF": round(cff, 2), "FCF": round(fcf, 2), "Net Change": round(net_cash_change, 2)})
                balance_rows.append({"Year": year, "Cash": round(end_cash, 2), "Receivables": round(ar, 2), "Inventory": round(inventory, 2), "PP&E net": round(end_fixed_assets, 2), "Total Assets": round(total_assets, 2), "Payables": round(ap, 2), "Debt": round(prev_debt, 2), "Total Liabilities": round(total_liabilities, 2), "Equity": round(end_equity, 2), "Check": round(balance_check, 2)})

                prev_cash = end_cash
                prev_fixed_assets_val = end_fixed_assets
                prev_equity = end_equity

            st.session_state.income_rows = income_rows
            st.session_state.cashflow_rows = cashflow_rows
            st.session_state.balance_rows = balance_rows
            st.session_state.three_table_params = {"depreciation_rate": depreciation_rate, "capex_rate": capex_rate, "ar_days": ar_days, "inv_days": inv_days, "ap_days": ap_days, "dividend_payout": dividend_payout}

        if st.session_state.cashflow_rows is not None:
            income_rows = st.session_state.income_rows
            cashflow_rows = st.session_state.cashflow_rows
            balance_rows = st.session_state.balance_rows

            st.subheader("Income Statement (USD bn)")
            st.dataframe(pd.DataFrame(income_rows).set_index("Year"))
            st.subheader("Cash Flow Statement (USD bn)")
            st.dataframe(pd.DataFrame(cashflow_rows).set_index("Year"))
            st.subheader("Balance Sheet (USD bn)")
            balance_display = pd.DataFrame(balance_rows).set_index("Year")
            st.dataframe(balance_display)

            max_diff = balance_display["Check"].abs().max()
            if max_diff < 1:
                st.success(f"✅ Balance sheet checks out. Max discrepancy: {round(max_diff, 3)} bn")
            else:
                st.warning(f"⚠️ Balance sheet discrepancy: {round(max_diff, 2)} bn — review assumptions")

            fig_cf = go.Figure()
            fig_cf.add_trace(go.Bar(x=[r["Year"] for r in cashflow_rows], y=[r["FCF"] for r in cashflow_rows], name="Free Cash Flow", marker_color="#70AD47"))
            fig_cf.add_trace(go.Scatter(x=[r["Year"] for r in income_rows], y=[r["Net Income"] for r in income_rows], name="Net Income", mode="lines+markers", line=dict(color="#ED7D31", width=2)))
            fig_cf.update_layout(title="FCF vs Net Income (smaller gap = higher earnings quality)", plot_bgcolor="white", hovermode="x unified", yaxis_title="USD bn")
            st.plotly_chart(fig_cf, use_container_width=True)

            if st.button("🤖 AI Commentary on 3-Statement Quality"):
                with st.spinner("Analyzing..."):
                    tp = st.session_state.three_table_params
                    p = st.session_state.forecast_params
                    prompt = f"""You are a senior financial analyst. Here is the 3-statement model for {p['cname']}:

Net Income: {[r['Net Income'] for r in income_rows]} (USD bn)
Free Cash Flow: {[r['FCF'] for r in cashflow_rows]} (USD bn)
Debt ratio: {[round(r['Total Liabilities']/r['Total Assets']*100,1) for r in balance_rows]} (%)

Assumptions: D&A rate {tp['depreciation_rate']}%, Capex rate {tp['capex_rate']}%, DSO {tp['ar_days']} days, Dividend payout {tp['dividend_payout']}%

Provide a professional analysis (~300 words):
1. Earnings quality: what does the FCF vs Net Income gap indicate?
2. Is the capex intensity reasonable? What does it imply for growth?
3. Are the working capital assumptions in line with industry norms?
4. Overall financial health assessment."""
                    client = anthropic.Anthropic(api_key=API_KEY)
                    message = client.messages.create(model="claude-opus-4-6", max_tokens=1500,
                        messages=[{"role": "user", "content": prompt}])
                    st.subheader("🤖 AI 3-Statement Commentary")
                    st.write(message.content[0].text)

# ============================================================
# TAB 4: VALUATION
# ============================================================
with tab4:
    st.header("DCF Valuation")
    st.write("Discount projected free cash flows to estimate intrinsic value per share.")

    if st.session_state.cashflow_rows is None:
        st.info("👈 Please complete the 3-Statement model in the **3-Statement** tab first.")
    else:
        cashflow_rows = st.session_state.cashflow_rows
        params = st.session_state.forecast_params
        fcf_list = [r["FCF"] for r in cashflow_rows]
        years = [r["Year"] for r in cashflow_rows]

        st.info(f"Projected FCF: **{years[0]}: {fcf_list[0]} bn** · **{years[1]}: {fcf_list[1]} bn** · **{years[2]}: {fcf_list[2]} bn**")

        st.subheader("DCF Assumptions")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown("**Discount Rate**")
            wacc = st.number_input("WACC %", value=9.0, min_value=1.0, max_value=30.0, step=0.5, help="Weighted average cost of capital. Typical range for tech: 8–12%")
        with col2:
            st.markdown("**Terminal Value**")
            terminal_growth = st.number_input("Terminal growth rate %", value=3.0, min_value=0.0, max_value=10.0, step=0.5, help="Perpetual growth rate after Year 3. Typically ~GDP growth (2–4%)")
        with col3:
            st.markdown("**Equity Bridge**")
            shares_outstanding = st.number_input("Shares outstanding (bn)", value=152.0, min_value=0.1, step=1.0)
            net_debt = st.number_input("Net debt (USD bn)", value=310.0, step=10.0, help="Debt minus cash. Negative = net cash position")

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

            def dcf_valuation(fcf_list, wacc_pct, tg_pct, shares, net_debt_val):
                w = wacc_pct / 100
                g = tg_pct / 100
                pv_fcf = sum(fcf / (1 + w) ** (i + 1) for i, fcf in enumerate(fcf_list))
                terminal_value = fcf_list[-1] * (1 + g) / (w - g)
                pv_terminal = terminal_value / (1 + w) ** len(fcf_list)
                ev = pv_fcf + pv_terminal
                equity_value = ev - net_debt_val
                price_per_share = equity_value / shares if shares > 0 else 0
                return {"pv_fcf": round(pv_fcf, 1), "terminal_value": round(terminal_value, 1),
                    "pv_terminal": round(pv_terminal, 1), "ev": round(ev, 1),
                    "equity_value": round(equity_value, 1), "price_per_share": round(price_per_share, 2),
                    "terminal_pct": round(pv_terminal / ev * 100, 1) if ev > 0 else 0}

            base = dcf_valuation(fcf_list, wacc, terminal_growth, shares_outstanding, net_debt)

            st.subheader("Base Case Results")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("PV of FCFs", f"{base['pv_fcf']} bn")
            c2.metric("PV of Terminal Value", f"{base['pv_terminal']} bn", f"{base['terminal_pct']}% of EV")
            c3.metric("Enterprise Value", f"{base['ev']} bn")
            c4.metric("Intrinsic Value / Share", f"${base['price_per_share']}")

            if base["terminal_pct"] > 80:
                st.warning(f"⚠️ Terminal value is {base['terminal_pct']}% of EV — valuation is highly sensitive to long-term assumptions")
            else:
                st.success(f"✅ Terminal value is {base['terminal_pct']}% of EV — near-term cash flows provide meaningful support")

            # 敏感性矩阵（用number_input替代slider，避免渲染冲突）
            st.subheader("Sensitivity Matrix — Intrinsic Value per Share ($)")
            wacc_steps = np.arange(wacc_low, wacc_high + 0.1, 1.0)
            tg_steps = np.arange(tg_low, tg_high + 0.1, 1.0)

            matrix_data = {}
            all_prices = []
            for w in wacc_steps:
                row_data = {}
                for g in tg_steps:
                    if w <= g:
                        row_data[f"g={g:.0f}%"] = "N/A"
                    else:
                        r = dcf_valuation(fcf_list, w, g, shares_outstanding, net_debt)
                        row_data[f"g={g:.0f}%"] = f"${r['price_per_share']}"
                        all_prices.append(r["price_per_share"])
                matrix_data[f"WACC={w:.0f}%"] = row_data

            matrix_df = pd.DataFrame(matrix_data).T
            st.dataframe(matrix_df, use_container_width=True)

            if all_prices:
                fig_val = go.Figure()
                fig_val.add_trace(go.Box(y=all_prices, name="Valuation range", marker_color="#4472C4", boxpoints="all", jitter=0.3))
                fig_val.add_hline(y=base["price_per_share"], line_dash="dash", line_color="red",
                    annotation_text=f"Base case ${base['price_per_share']}")
                fig_val.update_layout(title="Intrinsic Value Distribution (across WACC & terminal growth scenarios)",
                    plot_bgcolor="white", yaxis_title="USD per share", showlegend=False)
                st.plotly_chart(fig_val, use_container_width=True)

                low = round(min(all_prices), 2)
                high = round(max(all_prices), 2)
                mid = round(sum(all_prices) / len(all_prices), 2)
                st.info(f"📌 Valuation range: **${low} — ${high}** · Median: **${mid}**")

            st.session_state.dcf_result = base
            st.session_state.dcf_params = {
                "wacc": wacc, "terminal_growth": terminal_growth,
                "shares": shares_outstanding, "net_debt": net_debt,
                "price_low": min(all_prices) if all_prices else 0,
                "price_high": max(all_prices) if all_prices else 0,
            }

        if st.session_state.dcf_result is not None:
            if st.button("🤖 AI Valuation Report"):
                with st.spinner("Generating valuation report..."):
                    base = st.session_state.dcf_result
                    dp = st.session_state.dcf_params
                    p = st.session_state.forecast_params
                    cashflow_rows = st.session_state.cashflow_rows
                    fcf_list = [r["FCF"] for r in cashflow_rows]

                    prompt = f"""You are a senior equity research analyst writing a valuation report for {p['cname']}.

DCF Results:
- Base case intrinsic value: ${base['price_per_share']} per share
- Enterprise value: {base['ev']} bn USD
- Terminal value as % of EV: {base['terminal_pct']}%
- Valuation range: ${dp['price_low']:.1f} — ${dp['price_high']:.1f}

Key assumptions: WACC {dp['wacc']}%, terminal growth {dp['terminal_growth']}%
FCF forecast: {fcf_list} (USD bn, next 3 years)

Write a professional valuation report (~400 words) covering:
1. Valuation conclusion: is the stock undervalued or overvalued vs current market price?
2. Sensitivity of the valuation to key assumptions
3. Upside and downside risks
4. Investment recommendation (Buy / Hold / Sell) with target price range"""

                    client = anthropic.Anthropic(api_key=API_KEY)
                    message = client.messages.create(model="claude-opus-4-6", max_tokens=2000,
                        messages=[{"role": "user", "content": prompt}])
                    st.subheader("🤖 AI Valuation Report")
                    st.write(message.content[0].text)
