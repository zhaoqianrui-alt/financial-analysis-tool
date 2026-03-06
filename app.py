import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import anthropic

# ⚠️ 注意：API Key 建议放到环境变量里，这里暂时保留你的写法
from dotenv import load_dotenv
import os

load_dotenv()
API_KEY = os.getenv("ANTHROPIC_API_KEY")

# ============================================================
# 页面标题
# ============================================================
st.title("📊 AI财务分析工具")
st.write("上传公司财务数据，自动生成专业分析报告")

# ============================================================
# 模块一：主公司分析（你原来的功能 + 折线图优化）
# ============================================================
st.header("🏢 单公司分析")

uploaded_file = st.file_uploader("上传Excel文件（主公司）", type=["xlsx"])
company_name = st.text_input("公司名称", placeholder="例如：拓普集团")

if uploaded_file and company_name:
    df = pd.read_excel(uploaded_file)

    # 计算财务指标
    df["净利率%"] = round(df["净利润"] / df["营业收入"] * 100, 2)
    df["资产负债率%"] = round(df["总负债"] / df["总资产"] * 100, 2)
    df["收入增速%"] = round(df["营业收入"].pct_change() * 100, 2)

    # 显示数据表格
    st.subheader("财务数据")
    st.dataframe(df[["年份", "净利率%", "资产负债率%", "收入增速%"]])

    # --------------------------------------------------------
    # 折线图（用下拉菜单让用户选择想看哪些指标）
    # 背景知识：plotly 的 melt() 是把"宽表"转成"长表"
    # 宽表：一行一年，多列是不同指标
    # 长表：一行是"某年某指标的值"，这样 plotly 才能按颜色区分
    # --------------------------------------------------------
    st.subheader("📈 趋势图")

    available_metrics = ["净利率%", "资产负债率%", "收入增速%"]
    selected_metrics = st.multiselect(
        "选择要显示的指标（可多选）",
        options=available_metrics,
        default=["净利率%", "资产负债率%"]
    )

    if selected_metrics:
        chart_df = df[["年份"] + selected_metrics].copy()
        chart_df["年份"] = chart_df["年份"].astype(str)

        fig = px.line(
            chart_df.melt(id_vars="年份", var_name="指标", value_name="数值"),
            x="年份",
            y="数值",
            color="指标",
            markers=True,
            title=f"{company_name} 关键财务指标趋势",
            labels={"数值": "百分比 (%)", "年份": "年份"}
        )
        # 让图表更好看：加网格线、调整字体
        fig.update_layout(
            hovermode="x unified",   # 鼠标悬停时同时显示所有指标的值
            plot_bgcolor="white",
            yaxis=dict(gridcolor="#eeeeee"),
            legend=dict(orientation="h", yanchor="bottom", y=1.02)
        )
        st.plotly_chart(fig, use_container_width=True)

    # --------------------------------------------------------
    # 模块三：风险预警
    # 背景知识：根据金融行业经验设定的"警戒线"
    # 比如资产负债率 > 70% 意味着公司借钱太多，风险较高
    # --------------------------------------------------------
    st.subheader("🚨 风险预警")

    latest = df.iloc[-1]  # 取最新一年的数据
    warnings = []
    goods = []

    # 规则1：资产负债率
    debt_ratio = latest["资产负债率%"]
    if debt_ratio > 70:
        warnings.append(f"⚠️ **资产负债率过高**：{debt_ratio}%（警戒线：70%）——公司负债比例偏高，财务风险较大")
    elif debt_ratio > 60:
        warnings.append(f"🟡 **资产负债率偏高**：{debt_ratio}%（注意线：60%）——需要关注偿债能力")
    else:
        goods.append(f"✅ 资产负债率正常：{debt_ratio}%（低于60%，财务结构健康）")

    # 规则2：净利率
    net_margin = latest["净利率%"]
    if net_margin < 0:
        warnings.append(f"🔴 **净利率为负**：{net_margin}%——公司出现亏损，需警惕")
    elif net_margin < 5:
        warnings.append(f"🟡 **净利率偏低**：{net_margin}%（低于5%）——盈利能力较弱")
    else:
        goods.append(f"✅ 净利率良好：{net_margin}%（高于5%，盈利能力正常）")

    # 规则3：收入增速（如果有数据的话）
    if not pd.isna(latest["收入增速%"]):
        growth = latest["收入增速%"]
        if growth < -10:
            warnings.append(f"🔴 **收入大幅下滑**：{growth}%——业务规模明显萎缩")
        elif growth < 0:
            warnings.append(f"🟡 **收入出现下滑**：{growth}%——需关注业务发展趋势")
        else:
            goods.append(f"✅ 收入保持增长：{growth}%")

    # 显示预警结果
    if warnings:
        st.error("发现以下风险信号：")
        for w in warnings:
            st.markdown(w)
    if goods:
        st.success("以下指标表现良好：")
        for g in goods:
            st.markdown(g)

    # --------------------------------------------------------
    # AI分析报告（你原来的功能，保留不变）
    # --------------------------------------------------------
    if st.button("生成AI分析报告"):
        with st.spinner("AI正在分析，请稍候..."):
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


# ============================================================
# 模块二：同行业对比
# 背景知识：上传多个公司的Excel文件，系统把它们的指标
# 画在同一张图上，一眼就能看出谁更强
# ============================================================
st.divider()  # 画一条分隔线，让页面更清晰
st.header("⚖️ 同行业对比分析")
st.write("上传多个竞争对手的Excel文件，进行横向对比（文件格式需与主公司相同）")

# 允许用户一次上传多个文件
competitor_files = st.file_uploader(
    "上传竞争对手Excel文件（可多选）",
    type=["xlsx"],
    accept_multiple_files=True,
    key="competitors"  # key 是给 Streamlit 区分不同组件用的，不影响功能
)
competitor_names_input = st.text_input(
    "竞争对手名称（用逗号分隔）",
    placeholder="例如：比亚迪,国轩高科,亿纬锂能"
)

if competitor_files and competitor_names_input:
    competitor_names = [n.strip() for n in competitor_names_input.split(",")]

    if len(competitor_files) != len(competitor_names):
        st.warning(f"⚠️ 文件数量({len(competitor_files)})和名称数量({len(competitor_names)})不一致，请检查")
    else:
        # 把所有公司的"最新年份数据"收集到一个表格里做对比
        compare_rows = []

        # 如果主公司已经上传了，也加进来对比
        if uploaded_file and company_name:
            main_df = pd.read_excel(uploaded_file)
            main_df["净利率%"] = round(main_df["净利润"] / main_df["营业收入"] * 100, 2)
            main_df["资产负债率%"] = round(main_df["总负债"] / main_df["总资产"] * 100, 2)
            main_df["收入增速%"] = round(main_df["营业收入"].pct_change() * 100, 2)
            latest_main = main_df.iloc[-1]
            compare_rows.append({
                "公司": company_name,
                "净利率%": latest_main["净利率%"],
                "资产负债率%": latest_main["资产负债率%"],
                "收入增速%": latest_main["收入增速%"]
            })

        # 读取每个竞争对手的文件
        for file, name in zip(competitor_files, competitor_names):
            try:
                c_df = pd.read_excel(file)
                c_df["净利率%"] = round(c_df["净利润"] / c_df["营业收入"] * 100, 2)
                c_df["资产负债率%"] = round(c_df["总负债"] / c_df["总资产"] * 100, 2)
                c_df["收入增速%"] = round(c_df["营业收入"].pct_change() * 100, 2)
                latest_c = c_df.iloc[-1]
                compare_rows.append({
                    "公司": name,
                    "净利率%": latest_c["净利率%"],
                    "资产负债率%": latest_c["资产负债率%"],
                    "收入增速%": latest_c["收入增速%"]
                })
            except Exception as e:
                st.error(f"读取 {name} 的文件时出错：{e}")

        if compare_rows:
            compare_df = pd.DataFrame(compare_rows)

            st.subheader("📊 对比数据表")
            st.dataframe(compare_df.set_index("公司"))

            # 柱状图对比：每个指标一张图
            # 背景知识：柱状图适合"横向比较同一时间点的不同公司"
            # 折线图适合"纵向比较同一公司的时间变化"
            st.subheader("📊 指标对比柱状图")

            metric_to_show = st.selectbox(
                "选择要对比的指标",
                options=["净利率%", "资产负债率%", "收入增速%"]
            )

            fig_bar = px.bar(
                compare_df,
                x="公司",
                y=metric_to_show,
                color="公司",
                title=f"各公司 {metric_to_show} 对比（最新年份）",
                text=metric_to_show,  # 在柱子上显示数值
            )
            fig_bar.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
            fig_bar.update_layout(showlegend=False, plot_bgcolor="white")
            st.plotly_chart(fig_bar, use_container_width=True)

            # 雷达图：一张图同时展示所有指标的综合表现
            # 背景知识：雷达图像蜘蛛网，面积越大表示综合表现越好
            st.subheader("🕸️ 综合雷达图")

            fig_radar = go.Figure()
            categories = ["净利率%", "资产负债率%", "收入增速%"]

            for _, row in compare_df.iterrows():
                values = [row["净利率%"], row["资产负债率%"], row["收入增速%"]]
                values_closed = values + [values[0]]  # 闭合雷达图需要首尾相接
                fig_radar.add_trace(go.Scatterpolar(
                    r=values_closed,
                    theta=categories + [categories[0]],
                    fill="toself",
                    name=row["公司"],
                    opacity=0.5
                ))

            fig_radar.update_layout(
                polar=dict(radialaxis=dict(visible=True)),
                title="多公司综合财务指标雷达图"
            )
            st.plotly_chart(fig_radar, use_container_width=True)