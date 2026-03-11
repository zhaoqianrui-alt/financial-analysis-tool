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
# 背景知识：把数据获取逻辑单独写成函数，主程序只需要调用函数
# 这样代码更清晰，以后修改数据源也方便
# ============================================================

def is_a_share(code):
    """
    判断是A股还是美股
    A股代码规律：6位纯数字，比如 600031、000001、300750
    美股代码规律：纯字母，比如 AAPL、TSLA、NVDA
    """
    return code.strip().isdigit()


def process_us_data(code):
    """
    获取并处理美股数据
    yfinance 返回的数据：列是日期，行是财务科目
    我们需要转换成：列是科目，行是年份
    单位换算：美股数据是美元，除以1亿转成亿美元
    """
    import yfinance as yf

    ticker = yf.Ticker(code.upper())
    income = ticker.financials       # 利润表
    balance = ticker.balance_sheet   # 资产负债表

    rows = []
    for date in income.columns[:5]:  # 最近5年
        year = date.year
        try:
            revenue = income.loc["Total Revenue", date] / 1e8
            net_income = income.loc["Net Income", date] / 1e8
            total_assets = balance.loc["Total Assets", date] / 1e8

            # 总负债 = 总资产 - 股东权益
            # 背景知识：资产 = 负债 + 所有者权益，所以负债 = 资产 - 权益
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

    df = pd.DataFrame(rows).sort_values("年份").reset_index(drop=True)
    return df


def process_a_share_data(code):
    """
    获取并处理A股数据
    akshare 提供多种接口，这里用财务分析指标接口
    只取年报数据（日期以 12-31 结尾的）
    """
    import akshare as ak

    df = ak.stock_financial_analysis_indicator(symbol=code)

    # 只保留年报（12月31日的数据）
    df = df[df["日期"].str.endswith("12-31")].copy()
    df["年份"] = df["日期"].str[:4].astype(int)
    df = df.sort_values("年份").tail(5).reset_index(drop=True)

    return df


# ============================================================
# 页面主体
# ============================================================

st.title("📊 AI财务分析工具 v2")
st.write("输入股票代码，自动获取财务数据并生成分析")

# 输入区域：用两列排列，让界面更整洁
# 背景知识：st.columns([2,1]) 把页面分成2:1的两列
col1, col2 = st.columns([2, 1])
with col1:
    stock_code = st.text_input(
        "股票代码",
        placeholder="A股填6位数字（如 600031），美股填字母（如 AAPL）"
    )
with col2:
    company_name = st.text_input("公司名称", placeholder="例如：拓普集团")

if stock_code and company_name:
    if st.button("🔍 自动获取数据"):
        with st.spinner("正在从网络获取财务数据，请稍候..."):
            try:
                if is_a_share(stock_code):
                    # ---- A股流程 ----
                    st.info(f"✅ 识别为A股，正在获取 {company_name} 的数据...")
                    df = process_a_share_data(stock_code)
                    st.subheader("📋 财务指标数据（年报）")
                    st.dataframe(df)
                    st.success("数据获取成功！A股数据包含多项财务指标，后续版本将自动提取关键指标绘图。")

                else:
                    # ---- 美股流程 ----
                    st.info(f"✅ 识别为美股，正在从 Yahoo Finance 获取 {company_name} 的数据...")
                    df = process_us_data(stock_code)

                    if df.empty:
                        st.warning("未获取到数据，请检查股票代码是否正确")
                    else:
                        st.subheader("📋 财务数据")
                        st.dataframe(df)

                        # 计算指标
                        rev_col = "营业收入(亿美元)"
                        net_col = "净利润(亿美元)"
                        liab_col = "总负债(亿美元)"
                        asset_col = "总资产(亿美元)"

                        df["净利率%"] = round(df[net_col] / df[rev_col] * 100, 2)
                        df["资产负债率%"] = round(df[liab_col] / df[asset_col] * 100, 2)
                        df["收入增速%"] = round(df[rev_col].pct_change() * 100, 2)

                        st.subheader("📈 趋势图")
                        chart_df = df[["年份", "净利率%", "资产负债率%"]].copy()
                        chart_df["年份"] = chart_df["年份"].astype(str)
                        fig = px.line(
                            chart_df.melt(id_vars="年份", var_name="指标", value_name="数值"),
                            x="年份", y="数值", color="指标",
                            markers=True,
                            title=f"{company_name} 关键财务指标趋势"
                        )
                        fig.update_layout(hovermode="x unified", plot_bgcolor="white")
                        st.plotly_chart(fig, use_container_width=True)

                        # 风险预警
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

                        # AI分析
                        if st.button("生成AI分析报告"):
                            with st.spinner("AI正在分析..."):
                                data_text = ""
                                for _, row in df.iterrows():
                                    data_text += f"{int(row['年份'])}年：净利率{row['净利率%']}%，资产负债率{row['资产负债率%']}%，收入增速{row['收入增速%']}%\n"

                                prompt = f"""请你作为专业证券研究员，根据以下财务数据对{company_name}进行分析，输出200字左右的专业分析报告：

{data_text}

请分析：盈利能力趋势、财务风险、以及投资价值。"""

                                client = anthropic.Anthropic(api_key=API_KEY)
                                message = client.messages.create(
                                    model="claude-opus-4-6",
                                    max_tokens=1024,
                                    messages=[{"role": "user", "content": prompt}]
                                )
                                st.subheader("AI分析报告")
                                st.write(message.content[0].text)

            except Exception as e:
                st.error(f"获取数据时出错：{e}")
                st.write("💡 请检查：1）股票代码是否正确  2）网络是否正常")

# ============================================================
# 保留原有手动上传功能作为备用
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
