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
    return pd.DataFrame(rows).sort_values("年份").reset_index(drop=True)

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

请从以下角度给出专业点评（200字左右）：
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