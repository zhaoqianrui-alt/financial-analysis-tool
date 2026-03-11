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
    info = ticker.info or {}
    company_name = info.get("longName") or info.get("shortName") or code.upper()
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
    df = pd.DataFrame(rows).dropna().sort_values("年份").reset_index(drop=True)
    return df, company_name

def process_a_share_data(code):
    import akshare as ak

    def safe_float(val):
        try:
            return float(str(val).replace(",", "").replace("%", "").strip())
        except:
            return None

    def extract_rows(raw, date_col):
        """从原始DataFrame提取标准化指标"""
        rows = []
        for _, row in raw.iterrows():
            def get(candidates):
                for c in candidates:
                    if c in raw.columns and pd.notna(row.get(c)):
                        v = safe_float(row[c])
                        if v is not None:
                            return v
                return None

            net_margin  = get(["净利润率(%)", "销售净利率(%)", "净利率", "净利润率", "摊薄净利润率(%)"])
            debt_ratio  = get(["资产负债率(%)", "资产负债比率(%)", "资产负债率"])
            roe         = get(["净资产收益率(%)", "加权净资产收益率(%)", "ROE", "摊薄净资产收益率(%)"])
            rev_growth  = get(["主营业务收入增长率(%)", "营业收入增长率(%)", "收入增速", "营业总收入同比增长率(%)"])

            rows.append({
                "年份": row["年份"],
                "净利率%":    round(net_margin, 2)  if net_margin  is not None else None,
                "资产负债率%": round(debt_ratio, 2)  if debt_ratio  is not None else None,
                "ROE%":       round(roe, 2)          if roe         is not None else None,
                "收入增速%":  round(rev_growth, 2)   if rev_growth  is not None else None,
            })
        return rows

    # ---- 尝试主接口 ----
    try:
        raw = ak.stock_financial_analysis_indicator(symbol=code)

        if raw is None or raw.empty:
            raise ValueError("主接口返回空数据")

        # 自动识别日期列
        date_col = None
        for candidate in ["日期", "报告期", "period", "date"]:
            if candidate in raw.columns:
                date_col = candidate
                break
        if date_col is None:
            date_col = raw.columns[0]

        # 只保留年报
        filtered = raw[raw[date_col].astype(str).str.endswith("12-31")].copy()

        if filtered.empty:
            raise ValueError("过滤后无年报数据")

        filtered["年份"] = filtered[date_col].astype(str).str[:4].astype(int)
        filtered = filtered.sort_values("年份").tail(5).reset_index(drop=True)

        rows = extract_rows(filtered, date_col)
        result = pd.DataFrame(rows)

        if result.empty or result.dropna(how="all", subset=["净利率%","资产负债率%","ROE%"]).empty:
            raise ValueError("提取指标后数据为空")

        return result

    except Exception as e1:
        # ---- 备用接口：stock_financial_abstract ----
        try:
            raw2 = ak.stock_financial_abstract(symbol=code)

            if raw2 is None or raw2.empty:
                raise ValueError(f"备用接口也返回空数据，原始错误：{e1}")

            date_col2 = None
            for candidate in ["报告期", "日期", "period"]:
                if candidate in raw2.columns:
                    date_col2 = candidate
                    break
            if date_col2 is None:
                date_col2 = raw2.columns[0]

            filtered2 = raw2[raw2[date_col2].astype(str).str.endswith("12-31")].copy()

            if filtered2.empty:
                # 备用接口可能没有年报筛选，取最近5条
                filtered2 = raw2.tail(5).copy()

            filtered2["年份"] = filtered2[date_col2].astype(str).str[:4].astype(int)
            filtered2 = filtered2.sort_values("年份").tail(5).reset_index(drop=True)

            # 备用接口列名不同，直接返回原始数据
            filtered2 = filtered2.rename(columns={date_col2: "报告期"})
            return filtered2

        except Exception as e2:
            raise ValueError(f"A股数据获取失败。\n主接口错误：{e1}\n备用接口错误：{e2}\n\n请检查：\n1）股票代码是否正确（6位数字）\n2）网络是否正常\n3）该股票是否在akshare支持范围内")

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
# AUTO-PILOT：AI自动推断所有参数并一键跑完全流程
# ============================================================

def run_autopilot(df, cname, is_us):
    """
    输入：历史财务数据 DataFrame + 公司名
    输出：所有模块的计算结果（写入 session_state）
    
    核心思路：
    1. 用 Claude 读取历史数据，推断合理的假设参数
    2. 用这些参数调用现有的建模函数
    3. 把结果存入 session_state，各 tab 直接读取展示
    """
    import numpy as np
    import json

    # ── 防护：验证输入数据 ──────────────────────────────────────
    if df is None:
        raise ValueError("No data loaded. Please fetch company data first.")
    if not hasattr(df, 'columns'):
        raise ValueError("Data format error — please re-fetch the data.")
    if df.empty:
        raise ValueError("Data is empty. Please try fetching the data again.")
    if "年份" not in df.columns:
        raise ValueError(f"Unexpected data format (columns: {list(df.columns)}). Please re-fetch.")

    # ── Step 1：整理历史数据摘要，发给 AI ──────────────────────
    if is_us:
        rev_col  = "营业收入(亿美元)"
        net_col  = "净利润(亿美元)"
        asset_col = "总资产(亿美元)"
        liab_col  = "总负债(亿美元)"
        currency  = "USD bn"
    else:
        rev_col  = "营业收入(亿元)" if "营业收入(亿元)" in df.columns else None
        net_col  = "净利润(亿元)"   if "净利润(亿元)"  in df.columns else None
        asset_col = None
        liab_col  = None
        currency  = "CNY bn"

    # 构建发给 AI 的数据摘要
    has_abs = rev_col and rev_col in df.columns and net_col and net_col in df.columns

    if has_abs:
        latest = df.dropna(subset=[rev_col]).iloc[-1]
        base_rev  = float(latest[rev_col])
        base_year = int(latest["年份"])
        hist_summary = "\n".join(
            f"  {int(r['年份'])}: Revenue {r[rev_col]} {currency}, Net Income {r[net_col]} {currency}, "
            f"Net Margin {round(r[net_col]/r[rev_col]*100,1) if r[rev_col] else 'N/A'}%"
            for _, r in df.dropna(subset=[rev_col]).iterrows()
        )
    else:
        # A股只有比率数据
        base_rev  = 100.0   # 占位，用户可在各tab修改
        base_year = int(df["年份"].max()) if "年份" in df.columns else 2024
        hist_summary = "\n".join(
            f"  {int(r['年份'])}: Net Margin {r.get('净利率%','N/A')}%, "
            f"Debt Ratio {r.get('资产负债率%','N/A')}%, ROE {r.get('ROE%','N/A')}%"
            for _, r in df.iterrows()
        )

    # 获取资产负债表数据（用于推断shares/net_debt等）
    if asset_col and asset_col in df.columns:
        latest_assets  = float(df.dropna(subset=[asset_col]).iloc[-1][asset_col])
        latest_liab    = float(df.dropna(subset=[liab_col]).iloc[-1][liab_col])  if liab_col and liab_col in df.columns else latest_assets * 0.5
    else:
        latest_assets  = base_rev * 2.5 if has_abs else 250.0
        latest_liab    = latest_assets * 0.5

    equity_est   = latest_assets - latest_liab
    net_debt_est = latest_liab - (latest_assets * 0.15)   # 粗估净债务
    shares_est   = round(equity_est / 20, 1)              # 粗估股数（假设BVPS~20）
    shares_est   = max(shares_est, 0.5)

    # ── Step 2：调用 Claude，让它推断所有参数 ─────────────────
    prompt = f"""You are a senior equity analyst. Based on the following historical financial data for {cname}, 
infer reasonable forward-looking assumptions for a 3-year financial model.

HISTORICAL DATA ({currency}):
{hist_summary}

TASK: Return ONLY a valid JSON object (no markdown, no explanation) with these exact keys:

{{
  "revenue_growth_y1": <% e.g. 8.0>,
  "revenue_growth_y2": <% e.g. 7.0>,
  "revenue_growth_y3": <% e.g. 6.0>,
  "gross_margin": <% e.g. 42.0>,
  "expense_ratio": <% e.g. 18.0>,
  "tax_rate": <% e.g. 25.0>,
  "depreciation_rate": <% of PP&E, e.g. 10.0>,
  "capex_rate": <% of revenue, e.g. 5.0>,
  "ar_days": <integer, e.g. 55>,
  "inv_days": <integer, e.g. 40>,
  "ap_days": <integer, e.g. 45>,
  "dividend_payout": <% e.g. 30.0>,
  "wacc": <% e.g. 9.0>,
  "terminal_growth": <% e.g. 3.0>,
  "pe_bear": <e.g. 18.0>,
  "pe_base": <e.g. 25.0>,
  "pe_bull": <e.g. 32.0>,
  "pb_bear": <e.g. 2.0>,
  "pb_base": <e.g. 3.5>,
  "pb_bull": <e.g. 5.0>,
  "reasoning": "<2-3 sentences explaining your key assumptions>"
}}

Base your estimates on the historical trends above. Be conservative but realistic.
If data is limited (ratio-only), use industry averages appropriate for the sector."""

    msg = client.messages.create(
        model="claude-opus-4-6",
        max_tokens=1000,
        messages=[{"role": "user", "content": prompt}]
    )

    raw = msg.content[0].text.strip()
    # 去掉可能的 markdown 代码块
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    params = json.loads(raw.strip())

    reasoning = params.get("reasoning", "AI-generated assumptions based on historical data.")

    # ── Step 3：用 AI 参数跑 Forecast ─────────────────────────
    g1 = params["revenue_growth_y1"] / 100
    g2 = params["revenue_growth_y2"] / 100
    g3 = params["revenue_growth_y3"] / 100

    rev_y1 = base_rev * (1 + g1)
    rev_y2 = rev_y1  * (1 + g2)
    rev_y3 = rev_y2  * (1 + g3)

    driver_names     = ["Revenue Growth Rate (%)"]
    driver_base      = [100.0]
    driver_forecasts = [
        [100 * (1 + g1)],
        [100 * (1 + g1) * (1 + g2)],
        [100 * (1 + g1) * (1 + g2) * (1 + g3)],
    ]

    forecast_df, combined_df = build_income_statement(
        base_revenue   = base_rev,
        driver_names   = driver_names,
        driver_base    = driver_base,
        driver_forecasts = driver_forecasts,
        gross_margin   = params["gross_margin"],
        expense_ratio  = params["expense_ratio"],
        tax_rate       = params["tax_rate"],
        base_year      = base_year,
        hist_df        = df if has_abs else None,
        rev_col        = rev_col,
        net_col        = net_col,
    )

    st.session_state.forecast_df     = forecast_df
    st.session_state.combined_df     = combined_df
    st.session_state.forecast_params = {
        "gross_margin":  params["gross_margin"],
        "expense_ratio": params["expense_ratio"],
        "tax_rate":      params["tax_rate"],
        "cname":         cname,
    }

    # ── Step 4：用预测结果跑三表 ──────────────────────────────
    dep_rate  = params["depreciation_rate"]
    capex_rate = params["capex_rate"]
    ar_days   = int(params["ar_days"])
    inv_days  = int(params["inv_days"])
    ap_days   = int(params["ap_days"])
    div_pct   = params["dividend_payout"]

    # 估算基年资产负债表
    base_cash_val  = latest_assets * 0.15
    base_fa_val    = latest_assets * 0.40
    base_eq_val    = equity_est
    base_debt_val  = latest_liab * 0.6

    income_rows, cashflow_rows, balance_rows = [], [], []
    prev_cash, prev_fa, prev_eq, prev_debt = base_cash_val, base_fa_val, base_eq_val, base_debt_val

    for _, row in forecast_df.iterrows():
        revenue    = row["营业收入"]
        net_income = row["净利润"]
        ebit       = row["营业利润(EBIT)"]
        year       = row["年份"]

        capex        = revenue * capex_rate / 100
        depreciation = prev_fa * dep_rate / 100
        dividends    = net_income * div_pct / 100
        ar  = revenue * ar_days  / 365
        inv = revenue * inv_days / 365
        ap  = revenue * ap_days  / 365
        nwc = ar + inv - ap
        cfo = net_income + depreciation - (nwc * 0.1)
        cfi = -capex
        cff = -dividends
        fcf = cfo + cfi
        net_cash_change = cfo + cfi + cff

        end_cash   = prev_cash + net_cash_change
        end_fa     = prev_fa + capex - depreciation
        total_assets = end_cash + end_fa + ar + inv
        end_equity = prev_eq + (net_income - dividends)
        total_liab = ap + prev_debt
        check = total_assets - (total_liab + end_equity)

        income_rows.append({"Year": year, "Revenue": round(revenue,2), "D&A": round(depreciation,2), "EBIT": round(ebit,2), "Net Income": round(net_income,2), "Net Margin%": round(row["净利率%"],1)})
        cashflow_rows.append({"Year": year, "CFO": round(cfo,2), "Capex": round(-capex,2), "CFI": round(cfi,2), "CFF": round(cff,2), "FCF": round(fcf,2), "Net Change": round(net_cash_change,2)})
        balance_rows.append({"Year": year, "Cash": round(end_cash,2), "Receivables": round(ar,2), "Inventory": round(inv,2), "PP&E": round(end_fa,2), "Total Assets": round(total_assets,2), "Payables": round(ap,2), "Debt": round(prev_debt,2), "Total Liabilities": round(total_liab,2), "Equity": round(end_equity,2), "Check": round(check,2)})

        prev_cash, prev_fa, prev_eq = end_cash, end_fa, end_equity

    st.session_state.income_rows      = income_rows
    st.session_state.cashflow_rows    = cashflow_rows
    st.session_state.balance_rows     = balance_rows
    st.session_state.three_table_params = {
        "depreciation_rate": dep_rate, "capex_rate": capex_rate,
        "ar_days": ar_days, "inv_days": inv_days, "ap_days": ap_days,
        "dividend_payout": div_pct, "cname": cname,
    }

    # ── Step 5：跑 DCF ─────────────────────────────────────────
    fcf_list = [r["FCF"] for r in cashflow_rows]
    wacc     = params["wacc"]
    tg       = params["terminal_growth"]

    def dcf_val(fcf_list, wacc_pct, tg_pct, shares, nd):
        w, g = wacc_pct/100, tg_pct/100
        if w <= g:
            return None
        pv_fcf = sum(f/(1+w)**(i+1) for i,f in enumerate(fcf_list))
        tv     = fcf_list[-1]*(1+g)/(w-g)
        pv_tv  = tv/(1+w)**len(fcf_list)
        ev     = pv_fcf + pv_tv
        eq_val = ev - nd
        pps    = eq_val/shares if shares > 0 else 0
        return {"pv_fcf": round(pv_fcf,1), "pv_terminal": round(pv_tv,1),
                "ev": round(ev,1), "equity_value": round(eq_val,1),
                "price_per_share": round(pps,2),
                "terminal_pct": round(pv_tv/ev*100,1) if ev > 0 else 0}

    base_dcf = dcf_val(fcf_list, wacc, tg, shares_est, net_debt_est)
    if base_dcf is None:
        base_dcf = {"pv_fcf":0,"pv_terminal":0,"ev":0,"equity_value":0,"price_per_share":0,"terminal_pct":0}

    wacc_steps = np.arange(max(wacc-2,1), wacc+3, 1.0)
    tg_steps   = np.arange(max(tg-1,0.5), tg+2, 0.5)
    all_prices = []
    for w in wacc_steps:
        for g in tg_steps:
            r = dcf_val(fcf_list, w, g, shares_est, net_debt_est)
            if r:
                all_prices.append(r["price_per_share"])

    st.session_state.dcf_result = base_dcf
    st.session_state.dcf_params = {
        "wacc": wacc, "terminal_growth": tg,
        "shares": shares_est, "net_debt": net_debt_est,
        "price_low":  round(min(all_prices),2) if all_prices else 0,
        "price_high": round(max(all_prices),2) if all_prices else 0,
        "cname": cname,
    }

    # ── Step 6：跑 PE/PB ──────────────────────────────────────
    last_ni  = income_rows[-1]["Net Income"]
    last_eq  = balance_rows[-1]["Equity"]
    eps      = last_ni / shares_est  if shares_est > 0 else 0
    bvps     = last_eq / shares_est  if shares_est > 0 else 0

    st.session_state.pepb_result = {
        "eps": round(eps,2), "bvps": round(bvps,2),
        "pe_bear": params["pe_bear"], "pe_base": params["pe_base"], "pe_bull": params["pe_bull"],
        "pb_bear": params["pb_bear"], "pb_base": params["pb_base"], "pb_bull": params["pb_bull"],
        "pe_price_bear": round(eps * params["pe_bear"],2),
        "pe_price_base": round(eps * params["pe_base"],2),
        "pe_price_bull": round(eps * params["pe_bull"],2),
        "pb_price_bear": round(bvps * params["pb_bear"],2),
        "pb_price_base": round(bvps * params["pb_base"],2),
        "pb_price_bull": round(bvps * params["pb_bull"],2),
    }

    return params, reasoning, base_dcf, all_prices


# ============================================================
# 页面主体 - 标签页布局（每个Tab独立可用）
# ============================================================

st.title("📊 AI Financial Analysis Tool")
st.markdown("""
**Understand any company's financial health in minutes — no finance background needed.**

This tool walks you through four steps: fetch real data → build a forecast → model the full financials → estimate fair value.
Each step is independent: you can jump straight to Valuation if you already have your own numbers.
""")

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
    "pepb_result": None,
    "ai_report_text": None,
    "autopilot_done": False,
    "autopilot_params": None,
    "autopilot_reasoning": None,
}.items():
    if key not in st.session_state:
        st.session_state[key] = default

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Data",
    "📈 Forecast",
    "📑 3-Statement",
    "💰 Valuation",
    "📄 Export Report",
])

# ============================================================
# TAB 1: DATA
# ============================================================
with tab1:
    # ── 封面入口选择 ──────────────────────────────────────────
    if "entry_mode" not in st.session_state:
        st.session_state.entry_mode = None

    if st.session_state.entry_mode is None:
        st.markdown("""
<div style="text-align:center;padding:40px 0 20px;">
  <h2 style="color:#2E5C8A;">Where would you like to start?</h2>
  <p style="color:#555;font-size:1.05em;">Choose the path that fits your situation.</p>
</div>""", unsafe_allow_html=True)

        col_l, col_r = st.columns(2, gap="large")
        with col_l:
            st.markdown("""
<div style="background:#EBF2FA;border:2px solid #4472C4;border-radius:14px;padding:30px 24px;text-align:center;">
  <div style="font-size:2.5em;">🔍</div>
  <h3 style="color:#2E5C8A;margin:10px 0 6px;">I have a stock ticker</h3>
  <p style="color:#555;font-size:0.95em;">Enter a ticker symbol (e.g. AAPL, TSLA) and we'll automatically pull 5 years of real financial data from the market.</p>
  <p style="color:#777;font-size:0.85em;">Best for: investors researching public companies</p>
</div>""", unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("Start with a ticker →", use_container_width=True, key="entry_ticker"):
                st.session_state.entry_mode = "ticker"
                st.rerun()

        with col_r:
            st.markdown("""
<div style="background:#F0FAF0;border:2px solid #70AD47;border-radius:14px;padding:30px 24px;text-align:center;">
  <div style="font-size:2.5em;">📝</div>
  <h3 style="color:#2E5C8A;margin:10px 0 6px;">I'll enter my own numbers</h3>
  <p style="color:#555;font-size:0.95em;">You already have financial data — from an annual report, spreadsheet, or your own research. Skip straight to building forecasts.</p>
  <p style="color:#777;font-size:0.85em;">Best for: analysts with existing data</p>
</div>""", unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("Enter my own numbers →", use_container_width=True, key="entry_manual"):
                st.session_state.entry_mode = "manual"
                st.rerun()

    # ── 入口 A：股票代码 ──────────────────────────────────────
    elif st.session_state.entry_mode == "ticker":
        if st.button("← Back", key="back_ticker"):
            st.session_state.entry_mode = None
            st.rerun()

        st.header("📊 Fetch Company Data")
        st.write("Enter a ticker symbol — we'll look up the company name and pull all the numbers automatically.")

        stock_code = st.text_input(
            "Ticker Symbol",
            placeholder="US stock: AAPL, TSLA, NVDA  ·  A-share (China): 600031, 688017",
            help="A ticker is the short code used to identify a stock on an exchange. For example, Apple's ticker is AAPL."
        )

        if stock_code:
            if st.button("🔍 Fetch Data"):
                with st.spinner("Looking up company and fetching data..."):
                    try:
                        if is_a_share(stock_code):
                            df = process_a_share_data(stock_code)
                            auto_cname = stock_code  # A股暂无自动名称
                            st.session_state.fetched_df = df
                            st.session_state.fetched_company = auto_cname
                            st.session_state.is_us_stock = False
                            st.success(f"✅ Data fetched for {auto_cname}!")
                        else:
                            df, auto_cname = process_us_data(stock_code)
                            st.session_state.fetched_df = df
                            st.session_state.fetched_company = auto_cname
                            st.session_state.is_us_stock = True
                            if df.empty:
                                st.warning("No data found. Please double-check the ticker symbol.")
                            else:
                                rev_col = "营业收入(亿美元)"
                                net_col = "净利润(亿美元)"
                                liab_col = "总负债(亿美元)"
                                asset_col = "总资产(亿美元)"
                                df["净利率%"] = round(df[net_col] / df[rev_col] * 100, 2)
                                df["资产负债率%"] = round(df[liab_col] / df[asset_col] * 100, 2)
                                df["收入增速%"] = round(df[rev_col].pct_change() * 100, 2)
                                st.session_state.fetched_df = df
                                st.success(f"✅ Found **{auto_cname}** — data loaded!")
                    except Exception as e:
                        st.error(f"Error: {e}")

        if st.session_state.fetched_df is not None:
            df = st.session_state.fetched_df
            cname = st.session_state.fetched_company
            is_us = st.session_state.is_us_stock

            st.subheader(f"{cname} — {'Historical Financials (USD bn)' if is_us else '历史财务指标（A股）'}")
            st.dataframe(df)

            st.subheader("📖 What do these numbers mean?")
            latest = df.iloc[-1]

            def metric_card(label, value, good_range, explanation, is_good):
                color = "#d4edda" if is_good else "#f8d7da"
                icon = "✅" if is_good else "⚠️"
                st.markdown(f"""
<div style="background:{color};padding:14px 18px;border-radius:10px;margin-bottom:10px;">
<b>{icon} {label}: {value}</b><br>
<span style="font-size:0.9em;color:#333;">{explanation}</span><br>
<span style="font-size:0.82em;color:#555;">Healthy range: {good_range}</span>
</div>""", unsafe_allow_html=True)

            col_a, col_b = st.columns(2)
            with col_a:
                if pd.notna(latest.get("净利率%")):
                    nm = latest["净利率%"]
                    metric_card("Net Margin", f"{nm}%", "10–30% for most industries",
                        f"Out of every $100 in revenue, {cname} keeps ${round(nm,1)} as profit after all costs and taxes. {'Strong profitability.' if nm > 15 else 'Below average — watch cost trends.' if nm > 0 else 'The company is currently losing money.'}",
                        nm > 10)
                if pd.notna(latest.get("收入增速%")):
                    rg = latest["收入增速%"]
                    metric_card("Revenue Growth", f"{rg}%", ">10% = fast-growing, 0–10% = stable",
                        f"Revenue {'grew' if rg >= 0 else 'shrank'} by {abs(round(rg,1))}% last year. {'Strong momentum.' if rg > 15 else 'Steady growth.' if rg > 0 else 'Declining top line — needs investigation.'}",
                        rg > 5)
            with col_b:
                if pd.notna(latest.get("资产负债率%")):
                    dr = latest["资产负债率%"]
                    metric_card("Debt Ratio", f"{dr}%", "Below 60% is generally safe",
                        f"{round(dr,1)}% of {cname}'s assets are funded by debt. {'Conservative balance sheet.' if dr < 40 else 'Moderate leverage.' if dr < 60 else 'High leverage — financial risk is elevated.'}",
                        dr < 60)
                if pd.notna(latest.get("ROE%")):
                    roe = latest["ROE%"]
                    metric_card("Return on Equity (ROE)", f"{roe}%", "Above 15% is considered strong",
                        f"For every $100 shareholders invested, the company earned ${round(roe,1)}. {'Excellent capital efficiency.' if roe > 20 else 'Decent returns.' if roe > 10 else 'Low returns on equity.'}",
                        roe > 15)

            chart_cols = [c for c in ["净利率%", "资产负债率%", "ROE%", "收入增速%"] if c in df.columns]
            if chart_cols:
                st.subheader("📈 5-Year Trend")
                st.write("Trends often matter more than a single year's number. Look for consistent improvement or warning signs.")
                chart_df = df[["年份"] + chart_cols].copy()
                chart_df["年份"] = chart_df["年份"].astype(str)
                fig = px.line(chart_df.melt(id_vars="年份", var_name="Metric", value_name="Value"),
                    x="年份", y="Value", color="Metric", markers=True,
                    title=f"{cname} — Key Ratios (5-Year)")
                fig.update_layout(hovermode="x unified", plot_bgcolor="white",
                    legend=dict(orientation="h", yanchor="bottom", y=1.02))
                st.plotly_chart(fig, use_container_width=True)

            # ── 🚀 AUTO-PILOT 按钮（只在数据加载后显示）────────────
            st.divider()
            st.markdown("""
<div style="background:linear-gradient(135deg,#1a3a5c,#2E5C8A);padding:24px 28px;border-radius:14px;margin:10px 0;">
  <h3 style="color:white;margin:0 0 8px;">🚀 Auto-Pilot Mode</h3>
  <p style="color:#cce0ff;margin:0;font-size:0.98em;">
    Let AI read the historical data above and automatically set all assumptions —
    then run the full Forecast → 3-Statement → DCF → PE/PB pipeline in one click.
    You can review and adjust anything afterwards.
  </p>
</div>""", unsafe_allow_html=True)

            col_ap1, col_ap2 = st.columns([1, 2])
            with col_ap1:
                run_ap = st.button("🚀 Run Auto-Pilot", use_container_width=True, type="primary")
            with col_ap2:
                st.caption("⏱ Takes ~15 seconds · Uses AI to infer assumptions · All results editable afterwards")

            if run_ap:
                # 防护检查：确保数据已加载
                if st.session_state.fetched_df is None or st.session_state.fetched_df.empty:
                    st.error("⚠️ Please fetch company data first before running Auto-Pilot.")
                else:
                    with st.spinner("🤖 AI is reading the data and building your full financial model..."):
                        try:
                            ap_params, ap_reasoning, ap_dcf, ap_prices = run_autopilot(
                                df=st.session_state.fetched_df,
                                cname=st.session_state.fetched_company,
                                is_us=st.session_state.is_us_stock,
                            )
                            st.session_state.autopilot_params    = ap_params
                            st.session_state.autopilot_reasoning = ap_reasoning
                            st.session_state.autopilot_done      = True
                            st.success("✅ Auto-Pilot complete! All modules have been populated.")
                            st.rerun()
                        except Exception as e:
                            import traceback
                            st.error(f"Auto-Pilot failed: {e}")
                            with st.expander("🔍 Error details"):
                                st.code(traceback.format_exc())

        if st.session_state.get("autopilot_done"):
            ap_params = st.session_state.autopilot_params
            ap_reason = st.session_state.autopilot_reasoning

            with st.expander("📋 View AI-generated assumptions", expanded=True):
                st.info(f"💡 **AI Reasoning:** {ap_reason}")

                c1, c2, c3 = st.columns(3)
                with c1:
                    st.markdown("**📈 Revenue Forecast**")
                    st.metric("Year 1 growth", f"{ap_params['revenue_growth_y1']}%")
                    st.metric("Year 2 growth", f"{ap_params['revenue_growth_y2']}%")
                    st.metric("Year 3 growth", f"{ap_params['revenue_growth_y3']}%")
                with c2:
                    st.markdown("**💼 P&L Assumptions**")
                    st.metric("Gross Margin",  f"{ap_params['gross_margin']}%")
                    st.metric("Opex Ratio",    f"{ap_params['expense_ratio']}%")
                    st.metric("Tax Rate",      f"{ap_params['tax_rate']}%")
                with c3:
                    st.markdown("**💰 Valuation**")
                    st.metric("WACC",             f"{ap_params['wacc']}%")
                    st.metric("Terminal Growth",  f"{ap_params['terminal_growth']}%")
                    st.metric("Base PE",          f"{ap_params['pe_base']}x")

                st.markdown("---")
                dcf_r = st.session_state.dcf_result
                pepb  = st.session_state.pepb_result
                if dcf_r and pepb:
                    v1, v2, v3, v4 = st.columns(4)
                    v1.metric("DCF Intrinsic Value", f"${dcf_r['price_per_share']}/share")
                    v2.metric("PE Base Target",      f"${pepb['pe_price_base']}/share")
                    v3.metric("PB Base Target",      f"${pepb['pb_price_base']}/share")
                    dcf_lo = st.session_state.dcf_params.get("price_low", 0)
                    dcf_hi = st.session_state.dcf_params.get("price_high", 0)
                    consensus = round((dcf_r["price_per_share"] + pepb["pe_price_base"] + pepb["pb_price_base"]) / 3, 2)
                    v4.metric("🎯 Consensus", f"${consensus}/share")

                st.success("📄 All tabs are now populated. Head to **Export Report** to download the full analysis.")

    # ── 入口 B：手动上传 ──────────────────────────────────────
    elif st.session_state.entry_mode == "manual":
        if st.button("← Back", key="back_manual"):
            st.session_state.entry_mode = None
            st.rerun()

        st.header("📝 Upload Your Own Data")
        st.write("Upload an Excel file with your financial data. The file should have these columns:")
        st.code("年份 | 净利润 | 营业收入 | 总负债 | 总资产", language=None)
        st.info("💡 Once uploaded, head to the **Forecast** tab to start building your model. You can skip the Data tab entirely.")

        uploaded_file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])
        manual_company = st.text_input("Company Name", placeholder="e.g. My Portfolio Company")

        if uploaded_file and manual_company:
            df_m = pd.read_excel(uploaded_file)
            df_m["净利率%"] = round(df_m["净利润"] / df_m["营业收入"] * 100, 2)
            df_m["资产负债率%"] = round(df_m["总负债"] / df_m["总资产"] * 100, 2)
            df_m["收入增速%"] = round(df_m["营业收入"].pct_change() * 100, 2)
            st.session_state.fetched_df = df_m
            st.session_state.fetched_company = manual_company
            st.session_state.is_us_stock = False
            st.success(f"✅ Data loaded for {manual_company}")
            st.dataframe(df_m[["年份", "净利率%", "资产负债率%", "收入增速%"]])
            if st.button("Generate AI Summary", key="manual_ai"):
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
    st.header("📈 Income Statement Forecast")
    st.write("Build a 3-year projection of revenue and profit based on your assumptions about the business.")

    # ---- 数据来源 ----
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

    with st.expander("📥 Enter base year data manually", expanded=not use_auto):
        m_cname = st.text_input("Company name", value="My Company", key="f_cname")
        m_base_year = st.number_input("Base year", value=2024, min_value=2000, max_value=2030, step=1, key="f_year")
        m_base_rev = st.number_input("Base year revenue (USD bn)", value=100.0, min_value=0.1, step=1.0, key="f_rev",
            help="Total revenue the company earned in the most recent full year, in billions of USD.")

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

    # ── 收入驱动因子 ──────────────────────────────────────────
    st.subheader("Step 1 — Revenue Drivers")

    # 解释弹窗
    with st.expander("❓ What are Revenue Drivers — and what should I put here?"):
        st.markdown("""
**Revenue drivers are the key business variables that determine how much money the company makes.**

The idea is simple: instead of just guessing "revenue will grow 10%", you break revenue down into its components and forecast each one separately. This makes your assumptions more transparent and easier to challenge.

**Examples by industry:**

| Industry | Driver 1 | Driver 2 |
|----------|----------|----------|
| Car manufacturer | Units sold (mn cars) | Average selling price (USD) |
| Streaming service | Subscribers (mn) | Monthly fee (USD) |
| Retailer | Number of stores | Revenue per store (mn USD) |
| SaaS company | Paying customers | Annual revenue per customer (USD) |

**How it works in this tool:**
- You set what each driver was in the base year
- Then you forecast where each driver will be in years 1, 2, and 3
- Revenue = base revenue × (driver changes combined)

**Not sure what to use?** Just pick 1 driver called "Revenue Growth Rate (%)" and set the base to 100. Then forecast 105, 110, 115 for 5% annual growth.
""")

    num_drivers = st.radio("How many drivers do you want to use?", [1, 2, 3], index=1, horizontal=True,
        help="More drivers = more detailed model. Start with 2 if you're new to this.")

    driver_names, driver_bases = [], []
    default_names = ["Units Sold (mn)", "ASP (USD)", "Market Share%"]
    default_help = [
        "How many units/customers/subscribers does the company have? Enter the number in the base year.",
        "What is the average price per unit/customer/subscription? Enter in USD.",
        "What percentage of the total market does this company serve?"
    ]
    for i in range(num_drivers):
        st.markdown(f"**Driver {i+1}**")
        c1, c2 = st.columns([2, 1])
        with c1:
            name = st.text_input(f"Driver {i+1} name — what does it measure?",
                value=default_names[i], key=f"driver_name_{i}",
                help=default_help[i])
        with c2:
            base_val = st.number_input(f"Base value in {f_base_year}",
                value=100.0, key=f"driver_base_{i}",
                help=f"What was the value of this driver in {f_base_year}? Use the same unit as your driver name.")
        driver_names.append(name)
        driver_bases.append(base_val)

    # ── 3年预测 ───────────────────────────────────────────────
    st.subheader("Step 2 — 3-Year Forecasts")
    st.write("Now enter where you think each driver will be in each of the next 3 years.")
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

    # ── P&L 假设 ──────────────────────────────────────────────
    st.subheader("Step 3 — Profitability Assumptions")
    st.write("These ratios determine how much of the revenue turns into profit.")

    col_a, col_b, col_c = st.columns(3)
    with col_a:
        gross_margin = st.number_input("Gross Margin %", value=40.0, min_value=0.0, max_value=100.0, step=0.5,
            help="Revenue minus the direct cost of making the product/service, as a % of revenue.\n\nExample: If Apple sells a phone for $1,000 and it costs $600 to make, gross margin = 40%.\n\nTypical ranges: Software 60–80% | Manufacturing 20–40% | Retail 25–45%")
    with col_b:
        expense_ratio = st.number_input("Opex Ratio %", value=20.0, min_value=0.0, max_value=100.0, step=0.5,
            help="Operating expenses (sales, marketing, admin, R&D) as a % of revenue.\n\nThese are costs NOT directly tied to making the product — think office rent, salaries, advertising.\n\nLower is better. Best-in-class companies run at 10–20%.")
    with col_c:
        tax_rate = st.number_input("Tax Rate %", value=25.0, min_value=0.0, max_value=50.0, step=0.5,
            help="The effective corporate income tax rate.\n\nUS corporate tax: ~21%. Including state taxes: ~25–28%.\nChina: 25% standard rate. Tech companies may qualify for 15%.\n\nIf unsure, use 25%.")

    with st.expander("📊 How does this model calculate profit?"):
        st.markdown(f"""
The income statement is built step by step:

```
Revenue                    (your driver forecast)
− Cost of Goods Sold       (Revenue × (1 − Gross Margin%))
= Gross Profit             (Revenue × Gross Margin%)
− Operating Expenses       (Revenue × Opex Ratio%)
= Operating Profit (EBIT)
− Income Tax               (EBIT × Tax Rate%)
= Net Income               ← the bottom line
```

**Net Margin** = Net Income / Revenue — this is the % that ends up as profit.
""")

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

    # ============================================================
    # PE/PB 相对估值
    # 背景知识：
    # PE（市盈率）= 股价 / 每股收益(EPS)，反映市场愿意为每1元利润付多少钱
    # PB（市净率）= 股价 / 每股净资产(BVPS)，反映市场愿意为每1元净资产付多少钱
    # 相对估值的逻辑：给定一个"合理倍数"，反推目标价
    # ============================================================
    st.divider()
    st.subheader("📐 Relative Valuation — PE & PB")
    st.write("Estimate target price based on earnings and book value multiples.")

    # ---- 数据来源 ----
    # 优先从三表/预测模块读取净利润和股东权益
    auto_ni, auto_eq, auto_shares_rel = None, None, None
    if st.session_state.get("income_rows"):
        # 取最后一年预测净利润
        auto_ni = st.session_state.income_rows[-1]["Net Income"]
    if st.session_state.get("balance_rows"):
        auto_eq = st.session_state.balance_rows[-1]["Equity"]
    if st.session_state.get("dcf_params"):
        auto_shares_rel = st.session_state.dcf_params.get("shares")

    with st.expander("📥 Manual Input — enter financials directly",
                     expanded=(auto_ni is None)):
        rel_cname  = st.text_input("Company name", value=st.session_state.dcf_params.get("cname", "Company") if st.session_state.dcf_params else "Company", key="rel_cname")
        rel_ni     = st.number_input("Net Income — forward year (bn)", value=float(auto_ni) if auto_ni else 50.0, step=1.0, key="rel_ni")
        rel_eq     = st.number_input("Shareholders' Equity (bn)",      value=float(auto_eq) if auto_eq else 200.0, step=1.0, key="rel_eq")
        rel_shares = st.number_input("Shares outstanding (bn)",        value=float(auto_shares_rel) if auto_shares_rel else 152.0, step=1.0, key="rel_shares")

    # 如果上游有数据就用，否则用手动输入
    f_ni     = auto_ni     if auto_ni     else st.session_state.get("rel_ni",     50.0)
    f_eq     = auto_eq     if auto_eq     else st.session_state.get("rel_eq",     200.0)
    f_shares = auto_shares_rel if auto_shares_rel else st.session_state.get("rel_shares", 152.0)
    f_ni     = st.session_state.get("rel_ni",     f_ni)
    f_eq     = st.session_state.get("rel_eq",     f_eq)
    f_shares = st.session_state.get("rel_shares", f_shares)

    # ---- PE 假设 ----
    st.markdown("**PE Valuation**")
    col_pe1, col_pe2, col_pe3 = st.columns(3)
    with col_pe1:
        pe_bear = st.number_input("Bear case PE", value=20.0, min_value=1.0, step=1.0,
            help="Conservative multiple — e.g. trough historical PE")
    with col_pe2:
        pe_base = st.number_input("Base case PE", value=28.0, min_value=1.0, step=1.0,
            help="Mid-cycle or industry average PE")
    with col_pe3:
        pe_bull = st.number_input("Bull case PE", value=35.0, min_value=1.0, step=1.0,
            help="Premium multiple — e.g. peak historical or high-growth peers")

    # ---- PB 假设 ----
    st.markdown("**PB Valuation**")
    col_pb1, col_pb2, col_pb3 = st.columns(3)
    with col_pb1:
        pb_bear = st.number_input("Bear case PB", value=5.0, min_value=0.1, step=0.5,
            help="Low end of historical PB range")
    with col_pb2:
        pb_base = st.number_input("Base case PB", value=8.0, min_value=0.1, step=0.5,
            help="Average historical PB")
    with col_pb3:
        pb_bull = st.number_input("Bull case PB", value=12.0, min_value=0.1, step=0.5,
            help="Peak PB or premium peers")

    if st.button("📐 Run PE/PB Valuation"):
        # EPS = 净利润 / 总股数
        eps = f_ni / f_shares if f_shares > 0 else 0
        # BVPS = 所有者权益 / 总股数
        bvps = f_eq / f_shares if f_shares > 0 else 0

        # PE目标价
        pe_price_bear = round(eps * pe_bear, 2)
        pe_price_base = round(eps * pe_base, 2)
        pe_price_bull = round(eps * pe_bull, 2)

        # PB目标价
        pb_price_bear = round(bvps * pb_bear, 2)
        pb_price_base = round(bvps * pb_base, 2)
        pb_price_bull = round(bvps * pb_bull, 2)

        st.markdown(f"**EPS (forward):** ${round(eps,2)}  ·  **BVPS:** ${round(bvps,2)}")

        # ---- 展示结果表 ----
        st.subheader("PE & PB Target Price Summary")
        summary_df = pd.DataFrame({
            "Method": ["PE Valuation", "PB Valuation"],
            "Bear Case ($)": [pe_price_bear, pb_price_bear],
            "Base Case ($)": [pe_price_base, pb_price_base],
            "Bull Case ($)": [pe_price_bull, pb_price_bull],
        }).set_index("Method")
        st.dataframe(summary_df)

        # ---- 可视化：三法对比瀑布图 ----
        dcf_low  = st.session_state.dcf_params.get("price_low",  0) if st.session_state.dcf_params else 0
        dcf_high = st.session_state.dcf_params.get("price_high", 0) if st.session_state.dcf_params else 0
        dcf_base_price = st.session_state.dcf_result.get("price_per_share", 0) if st.session_state.dcf_result else 0

        fig_rel = go.Figure()

        # DCF 区间
        if dcf_low and dcf_high:
            fig_rel.add_trace(go.Bar(
                name="DCF", x=["DCF"],
                y=[dcf_high - dcf_low],
                base=[dcf_low],
                marker_color="#4472C4",
                width=0.4,
            ))
            fig_rel.add_trace(go.Scatter(
                x=["DCF"], y=[dcf_base_price],
                mode="markers", name="DCF Base",
                marker=dict(color="#4472C4", size=12, symbol="diamond"),
                showlegend=True,
            ))

        # PE 区间
        fig_rel.add_trace(go.Bar(
            name="PE", x=["PE"],
            y=[pe_price_bull - pe_price_bear],
            base=[pe_price_bear],
            marker_color="#ED7D31",
            width=0.4,
        ))
        fig_rel.add_trace(go.Scatter(
            x=["PE"], y=[pe_price_base],
            mode="markers", name="PE Base",
            marker=dict(color="#ED7D31", size=12, symbol="diamond"),
            showlegend=True,
        ))

        # PB 区间
        fig_rel.add_trace(go.Bar(
            name="PB", x=["PB"],
            y=[pb_price_bull - pb_price_bear],
            base=[pb_price_bear],
            marker_color="#70AD47",
            width=0.4,
        ))
        fig_rel.add_trace(go.Scatter(
            x=["PB"], y=[pb_price_base],
            mode="markers", name="PB Base",
            marker=dict(color="#70AD47", size=12, symbol="diamond"),
            showlegend=True,
        ))

        fig_rel.update_layout(
            title="Valuation Range Comparison — DCF vs PE vs PB",
            plot_bgcolor="white",
            yaxis_title="Target Price (USD)",
            barmode="overlay",
            showlegend=True,
            hovermode="x unified",
        )
        st.plotly_chart(fig_rel, use_container_width=True)

        # ---- 三法交叉：取重叠区间 ----
        st.subheader("📌 Cross-Validation Summary")
        all_lows  = [dcf_low, pe_price_bear, pb_price_bear] if dcf_low else [pe_price_bear, pb_price_bear]
        all_highs = [dcf_high, pe_price_bull, pb_price_bull] if dcf_high else [pe_price_bull, pb_price_bull]
        all_bases = [dcf_base_price, pe_price_base, pb_price_base] if dcf_base_price else [pe_price_base, pb_price_base]

        consensus_low  = round(max(all_lows), 2)   # 三法下限的最大值（最保守的共识）
        consensus_high = round(min(all_highs), 2)  # 三法上限的最小值（最谨慎的上限）
        consensus_mid  = round(sum(all_bases) / len(all_bases), 2)

        if consensus_low < consensus_high:
            st.success(f"✅ **Consensus range: ${consensus_low} — ${consensus_high}** · Average base case: **${consensus_mid}**")
            st.write("Three methods agree on this overlap zone — higher confidence in this range.")
        else:
            st.warning(f"⚠️ No overlapping range across methods. Average base case: **${consensus_mid}**")
            st.write("Methods diverge significantly — review assumptions for consistency.")

        # 存储PE/PB结果
        st.session_state.pepb_result = {
            "pe_bear": pe_price_bear, "pe_base": pe_price_base, "pe_bull": pe_price_bull,
            "pb_bear": pb_price_bear, "pb_base": pb_price_base, "pb_bull": pb_price_bull,
            "eps": round(eps, 2), "bvps": round(bvps, 2),
            "consensus_low": consensus_low, "consensus_high": consensus_high,
            "consensus_mid": consensus_mid,
            "cname": rel_cname,
        }

    # ---- AI综合三法报告 ----
    if st.session_state.get("pepb_result") is not None:
        if st.button("🤖 AI Cross-Valuation Report"):
            with st.spinner("Generating comprehensive valuation report..."):
                pr = st.session_state.pepb_result
                dcf_r = st.session_state.dcf_result
                dp = st.session_state.dcf_params

                dcf_summary = f"DCF base: ${dcf_r['price_per_share']}, range ${dp['price_low']:.1f}–${dp['price_high']:.1f}" if dcf_r else "DCF: not run"

                prompt = f"""You are a senior equity research analyst. Write a comprehensive valuation report for {pr['cname']}.

Valuation results:
- {dcf_summary}
- PE valuation: bear ${pr['pe_bear']} / base ${pr['pe_base']} / bull ${pr['pe_bull']} (EPS: ${pr['eps']})
- PB valuation: bear ${pr['pb_bear']} / base ${pr['pb_base']} / bull ${pr['pb_bull']} (BVPS: ${pr['bvps']})
- Consensus range: ${pr.get('consensus_low','N/A')} — ${pr.get('consensus_high','N/A')}, average base ${pr.get('consensus_mid','N/A')}

Write a professional valuation conclusion (~400 words):
1. Which method is most appropriate for this company and why?
2. What does the convergence (or divergence) between methods tell us?
3. Final target price recommendation with Bull/Base/Bear scenarios
4. Investment rating: Strong Buy / Buy / Hold / Sell / Strong Sell"""

                msg = client.messages.create(model="claude-opus-4-6", max_tokens=2000,
                    messages=[{"role": "user", "content": prompt}])
                st.session_state.ai_report_text = msg.content[0].text
                st.subheader("🤖 AI Cross-Valuation Report")
                st.write(msg.content[0].text)

# ============================================================
# TAB 5: EXPORT REPORT
# ============================================================
with tab5:
    st.header("📄 Export Research Report")
    st.write("Generate a professional Word document (.docx) summarising all your analysis.")

    # 检查有哪些模块已完成
    has_data     = st.session_state.fetched_df is not None
    has_forecast = st.session_state.forecast_df is not None
    has_3stmt    = st.session_state.cashflow_rows is not None
    has_dcf      = st.session_state.dcf_result is not None
    has_pepb     = st.session_state.pepb_result is not None
    has_autopilot = st.session_state.get("autopilot_done", False)

    # Auto-Pilot 完成提示
    if has_autopilot:
        ap = st.session_state.autopilot_params
        dcf_r = st.session_state.dcf_result
        pepb  = st.session_state.pepb_result
        if dcf_r and pepb:
            consensus = round((dcf_r["price_per_share"] + pepb.get("pe_price_base",0) + pepb.get("pb_price_base",0)) / 3, 2)
            st.success(f"🚀 **Auto-Pilot results ready** — Consensus target price: **${consensus}/share** · Click Generate Report below to export.")
        else:
            st.success("🚀 Auto-Pilot complete — all modules populated.")

    st.subheader("Completed Sections")
    col_s1, col_s2, col_s3, col_s4, col_s5 = st.columns(5)
    col_s1.metric("📊 Data",        "✅ Done" if has_data     else "⬜ Pending")
    col_s2.metric("📈 Forecast",    "✅ Done" if has_forecast else "⬜ Pending")
    col_s3.metric("📑 3-Statement", "✅ Done" if has_3stmt    else "⬜ Pending")
    col_s4.metric("💰 DCF",         "✅ Done" if has_dcf      else "⬜ Pending")
    col_s5.metric("📐 PE/PB",       "✅ Done" if has_pepb     else "⬜ Pending")

    if not (has_data or has_forecast or has_dcf or has_pepb):
        st.info("👈 Complete at least one analysis tab to generate a report.")
    else:
        report_title = st.text_input("Report title", value=f"{st.session_state.fetched_company or 'Company'} — Equity Research Report")
        analyst_name = st.text_input("Analyst name", value="")
        include_ai_summary = st.checkbox("Include AI executive summary", value=True)

        if st.button("📄 Generate Word Report"):
            with st.spinner("Building report..."):
                try:
                    import subprocess, json, tempfile, base64
                    from datetime import date

                    # ---- 收集所有数据 ----
                    cname_r = st.session_state.fetched_company or "Company"
                    today   = date.today().strftime("%B %d, %Y")

                    # AI执行摘要
                    exec_summary = ""
                    if include_ai_summary:
                        parts = []
                        if has_data:
                            latest = st.session_state.fetched_df.iloc[-1]
                            nm  = latest.get("净利率%", "N/A")
                            dr  = latest.get("资产负债率%", "N/A")
                            roe = latest.get("ROE%", "N/A")
                            parts.append(f"Historical: Net margin {nm}%, Debt ratio {dr}%, ROE {roe}%")
                        if has_forecast:
                            fp = st.session_state.forecast_params
                            fd = st.session_state.forecast_df
                            parts.append(f"Forecast: GM {fp['gross_margin']}%, Opex {fp['expense_ratio']}%, Tax {fp['tax_rate']}%. Projected net income: {list(fd['净利润'])}")
                        if has_dcf:
                            d = st.session_state.dcf_result
                            dp2 = st.session_state.dcf_params
                            parts.append(f"DCF: base ${d['price_per_share']}/share, EV ${d['ev']}bn, range ${dp2['price_low']:.1f}–${dp2['price_high']:.1f}")
                        if has_pepb:
                            pr = st.session_state.pepb_result
                            _c_low  = pr.get('consensus_low',  round((pr.get('pe_price_base', pr.get('pe_base',0)) + pr.get('pb_price_base', pr.get('pb_base',0)))/2*0.9, 2))
                            _c_high = pr.get('consensus_high', round((pr.get('pe_price_base', pr.get('pe_base',0)) + pr.get('pb_price_base', pr.get('pb_base',0)))/2*1.1, 2))
                            _c_mid  = pr.get('consensus_mid',  round((pr.get('pe_price_base', pr.get('pe_base',0)) + pr.get('pb_price_base', pr.get('pb_base',0)))/2, 2))
                            parts.append(f"PE/PB consensus: ${_c_low}–${_c_high}, mid ${_c_mid}")

                        ai_prompt = f"""Write a concise executive summary (200 words) for an equity research report on {cname_r}.
Data: {'; '.join(parts)}
Cover: investment thesis, key financial strengths/risks, valuation conclusion, and rating (Buy/Hold/Sell)."""
                        ai_msg = client.messages.create(model="claude-opus-4-6", max_tokens=600,
                            messages=[{"role": "user", "content": ai_prompt}])
                        exec_summary = ai_msg.content[0].text

                    # ---- 用纯标准库生成 docx（无需任何外部依赖）----
                    import io, zipfile, textwrap
                    from datetime import date as _date

                    def _xml(tag, attrs="", text="", children=""):
                        a = f" {attrs}" if attrs else ""
                        if text:
                            return f"<{tag}{a}>{text}</{tag}>"
                        return f"<{tag}{a}>{children}</{tag}>"

                    def _para(text, bold=False, size=24, color="000000", align="left", heading=None):
                        b  = "<w:b/><w:bCs/>" if bold else ""
                        sz = f"<w:sz w:val=\"{size}\"/><w:szCs w:val=\"{size}\"/>"
                        cl = f'<w:color w:val=\"{color}\"/>'
                        jc = f'<w:jc w:val=\"{align}\"/>'
                        ppr = f"<w:pPr>{jc}</w:pPr>"
                        rpr = f"<w:rPr>{b}{sz}{cl}</w:rPr>"
                        run = f"<w:r>{rpr}<w:t xml:space=\"preserve\">{text}</w:t></w:r>"
                        return f"<w:p>{ppr}{run}</w:p>"

                    def _spacer():
                        return "<w:p><w:r><w:t></w:t></w:r></w:p>"

                    def _table(headers, rows):
                        cell_w = max(1, 9360 // max(len(headers), 1))
                        def _cell(text, header=False, shade=None):
                            bg = ""
                            if header:
                                bg = '<w:shd w:val="clear" w:color="auto" w:fill="2E5C8A"/>'
                            elif shade:
                                bg = f'<w:shd w:val="clear" w:color="auto" w:fill="{shade}"/>'
                            tc_color = "FFFFFF" if header else "000000"
                            tc_bold  = "<w:b/><w:bCs/>" if header else ""
                            return (f'<w:tc>'
                                    f'<w:tcPr><w:tcW w:w="{cell_w}" w:type="dxa"/>{bg}'
                                    f'<w:tcMar><w:top w:w="80" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/>'
                                    f'<w:left w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar></w:tcPr>'
                                    f'<w:p><w:r><w:rPr>{tc_bold}<w:color w:val="{tc_color}"/><w:sz w:val="20"/></w:rPr>'
                                    f'<w:t xml:space="preserve">{str(text) if text is not None else "—"}</w:t></w:r></w:p>'
                                    f'</w:tc>')
                        hdr_row = "<w:tr>" + "".join(_cell(h, header=True) for h in headers) + "</w:tr>"
                        data_rows = ""
                        for ri, row in enumerate(rows):
                            shade = "EEF3FA" if ri % 2 == 1 else None
                            data_rows += "<w:tr>" + "".join(_cell(v, shade=shade) for v in row) + "</w:tr>"
                        return f'<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="9360" w:type="dxa"/></w:tblPr><w:tblGrid>{"".join(f"<w:gridCol w:w=\"{cell_w}\"/>" for _ in headers)}</w:tblGrid>{hdr_row}{data_rows}</w:tbl>'

                    # 组装文档内容
                    body_parts = []

                    # 封面
                    body_parts.append(_spacer())
                    body_parts.append(_para(report_title, bold=True, size=44, color="2E5C8A", align="center"))
                    body_parts.append(_para(f"Prepared by: {analyst_name or '—'}   ·   Date: {today}", size=22, color="666666", align="center"))
                    body_parts.append(_para("Powered by AI Financial Analysis Tool", size=20, color="888888", align="center"))
                    body_parts.append(_spacer())

                    # 执行摘要
                    if exec_summary:
                        body_parts.append(_para("Executive Summary", bold=True, size=32, color="2E5C8A"))
                        body_parts.append(_spacer())
                        for line in exec_summary.split("\n"):
                            if line.strip():
                                body_parts.append(_para(line.strip()))
                        body_parts.append(_spacer())

                    # Section 1: 历史数据
                    if has_data:
                        body_parts.append(_para("1. Historical Financial Overview", bold=True, size=28, color="2E5C8A"))
                        body_parts.append(_spacer())
                        df_r = st.session_state.fetched_df
                        cols = list(df_r.columns)
                        trows = []
                        for _, row in df_r.iterrows():
                            trows.append([str(round(v, 2)) if isinstance(v, float) else str(v) for v in row])
                        body_parts.append(_table(cols, trows))
                        body_parts.append(_spacer())

                    # Section 2: 预测
                    if has_forecast:
                        body_parts.append(_para("2. Income Statement Forecast (3-Year)", bold=True, size=28, color="2E5C8A"))
                        fp = st.session_state.forecast_params
                        body_parts.append(_para(f"Assumptions: GM {fp['gross_margin']}%  Opex {fp['expense_ratio']}%  Tax {fp['tax_rate']}%", size=20))
                        body_parts.append(_spacer())
                        fd = st.session_state.forecast_df[["年份","营业收入","净利润","净利率%"]]
                        body_parts.append(_table(list(fd.columns), fd.values.tolist()))
                        body_parts.append(_spacer())

                    # Section 3: 三表
                    if has_3stmt:
                        body_parts.append(_para("3. Three-Statement Model", bold=True, size=28, color="2E5C8A"))
                        body_parts.append(_spacer())
                        for label, key in [("Income Statement","income_rows"),("Cash Flow Statement","cashflow_rows"),("Balance Sheet","balance_rows")]:
                            rlist = st.session_state[key]
                            if rlist:
                                body_parts.append(_para(label, bold=True, size=24, color="1A3F6B"))
                                hdrs = list(rlist[0].keys())
                                body_parts.append(_table(hdrs, [[r[k] for k in hdrs] for r in rlist]))
                                body_parts.append(_spacer())

                    # Section 4: DCF
                    if has_dcf:
                        body_parts.append(_para("4. DCF Valuation", bold=True, size=28, color="2E5C8A"))
                        body_parts.append(_spacer())
                        d   = st.session_state.dcf_result
                        dp2 = st.session_state.dcf_params
                        body_parts.append(_table(["Metric","Value"],[
                            ["Base Case Price/Share", f"${d['price_per_share']}"],
                            ["Enterprise Value",      f"{d['ev']} bn"],
                            ["PV of FCFs",            f"{d['pv_fcf']} bn"],
                            ["PV of Terminal Value",  f"{d['pv_terminal']} bn ({d['terminal_pct']}% of EV)"],
                            ["WACC",                  f"{dp2['wacc']}%"],
                            ["Terminal Growth Rate",  f"{dp2['terminal_growth']}%"],
                            ["Valuation Range",       f"${dp2.get('price_low',0):.1f} — ${dp2.get('price_high',0):.1f}"],
                        ]))
                        body_parts.append(_spacer())

                    # Section 5: PE/PB
                    if has_pepb:
                        body_parts.append(_para("5. Relative Valuation (PE & PB)", bold=True, size=28, color="2E5C8A"))
                        body_parts.append(_spacer())
                        pr = st.session_state.pepb_result
                        pe_bear = pr.get("pe_price_bear", pr.get("pe_bear", 0))
                        pe_base = pr.get("pe_price_base", pr.get("pe_base", 0))
                        pe_bull = pr.get("pe_price_bull", pr.get("pe_bull", 0))
                        pb_bear = pr.get("pb_price_bear", pr.get("pb_bear", 0))
                        pb_base = pr.get("pb_price_base", pr.get("pb_base", 0))
                        pb_bull = pr.get("pb_price_bull", pr.get("pb_bull", 0))
                        body_parts.append(_table(["Method","Bear","Base","Bull"],[
                            ["PE Valuation", f"${pe_bear}", f"${pe_base}", f"${pe_bull}"],
                            ["PB Valuation", f"${pb_bear}", f"${pb_base}", f"${pb_bull}"],
                        ]))
                        c_low = pr.get("consensus_low",  round((pe_base+pb_base)/2*0.9, 2))
                        c_high= pr.get("consensus_high", round((pe_base+pb_base)/2*1.1, 2))
                        c_mid = pr.get("consensus_mid",  round((pe_base+pb_base)/2,     2))
                        body_parts.append(_spacer())
                        body_parts.append(_para(f"Consensus range: ${c_low} — ${c_high}  ·  Avg base: ${c_mid}", bold=True, color="2E5C8A"))
                        body_parts.append(_spacer())

                    # 免责声明
                    body_parts.append(_para("Disclaimer: This report is generated by an AI tool for informational purposes only and does not constitute investment advice.", size=18, color="888888"))

                    # 组装完整 docx XML
                    body_xml = "\n".join(body_parts)
                    doc_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
  xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:w10="urn:schemas-microsoft-com:office:word"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
  xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
  xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"
  xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
  xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
  xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  mc:Ignorable="w14 w15 w16se w16cid wp14">
  <w:body>
{body_xml}
  </w:body>
</w:document>'''

                    rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

                    styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr>
    <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
    <w:sz w:val="22"/><w:szCs w:val="22"/>
  </w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="table" w:styleId="TableGrid">
    <w:name w:val="Table Grid"/>
    <w:tblPr><w:tblBorders>
      <w:top w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>
      <w:left w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>
      <w:right w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>
      <w:insideV w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>
    </w:tblBorders></w:tblPr>
  </w:style>
</w:styles>'''

                    content_types_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''

                    root_rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

                    # 打包成 zip（docx 本质就是 zip）
                    buf = io.BytesIO()
                    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                        zf.writestr("[Content_Types].xml", content_types_xml)
                        zf.writestr("_rels/.rels", root_rels_xml)
                        zf.writestr("word/document.xml", doc_xml)
                        zf.writestr("word/styles.xml", styles_xml)
                        zf.writestr("word/_rels/document.xml.rels", rels_xml)
                    docx_bytes = buf.getvalue()

                    st.success("✅ Report generated!")
                    st.download_button(
                        label="⬇️ Download Report (.docx)",
                        data=docx_bytes,
                        file_name=f"{cname_r.replace(' ','_')}_Research_Report.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )


                except Exception as e:
                    st.error(f"Error: {e}")
