"""
AI销售分析助手
功能：上传销售Excel -> 销售概况 -> 产品分析 -> 趋势预测 -> 客户分析 -> 地域分析 -> AI报告
"""
import streamlit as st
from openai import OpenAI
import os
from dotenv import load_dotenv
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from docx import Document
from docx.shared import Inches
from io import BytesIO
import re

load_dotenv()

DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY")

CHART_COLORS = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd',
                '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf']


def safe_write_image(fig, filename):
    try:
        fig.write_image(filename)
    except Exception:
        pass


def style_chart(fig):
    fig.update_layout(
        paper_bgcolor='white',
        plot_bgcolor='white',
        font=dict(color='black', size=14),
        xaxis=dict(tickfont=dict(color='black', size=12), title_font=dict(color='black', size=13)),
        yaxis=dict(tickfont=dict(color='black', size=12), title_font=dict(color='black', size=13)),
        legend=dict(font=dict(color='black', size=12))
    )
    return fig


def extract_province(address):
    """从配送地址中提取省份/直辖市"""
    if not isinstance(address, str):
        return "未知"
    match = re.match(r'^(.+?[省市])', address)
    return match.group(1) if match else str(address)[:3]


def get_ai_report(data, forecast_data=None, customer_data=None, geo_data=None):
    """调用DeepSeek API获取AI分析报告"""
    if not DEEPSEEK_API_KEY:
        return "⚠️ 未配置DEEPSEEK_API_KEY环境变量，无法使用AI分析功能。"

    prompt = f"""
    你是一名销售分析师。
    根据以下销售数据生成分析报告：
    销售额：{data["总销售额"]}
    销量：{data["总销量"]}
    热销产品：{data["热销产品"]}
    """
    if forecast_data:
        prompt += f"""
    环比变化：{forecast_data.get("环比变化", "无数据")}
    同比变化：{forecast_data.get("同比变化", "无数据")}
    下月预测销量（趋势）：{forecast_data.get("趋势预测", "无数据")}
    下月预测销量（移动平均）：{forecast_data.get("移动平均预测", "无数据")}
    """
    if customer_data:
        prompt += f"""
    --- 客户数据 ---
    客户总数：{customer_data["客户总数"]}
    人均消费金额：{customer_data["人均消费"]}
    客单价：{customer_data["客单价"]}
    Top5高价值客户：{customer_data["top5客户"]}
    高价值客户收入占比：{customer_data["高价值占比"]}
    购买频次分布：{customer_data["频次分布"]}
    """
    if geo_data:
        prompt += f"""
    --- 地域数据 ---
    覆盖地区数：{geo_data["地区数"]}
    Top5销售地区：{geo_data["top5地区"]}
    地区集中度（Top3地区销售额占比）：{geo_data["地区集中度"]}
    """

    prompt += """
    请严格按照下面格式输出：
    ## 📌 销售总结
    总结当前销售情况。
    ## 🔍 核心发现
    列出3条数据发现：
    1.
    2.
    3.
    ## 👥 客户分析
    分析客户价值分层、高价值客户特征、客户留存建议。
    ## 🗺️ 地域分析
    分析各地区销售差异、地域扩张建议、区域产品偏好。
    ## 📈 趋势与预测分析
    分析同比环比变化和需求预测结果，给出供应链备货建议。
    ## 💡 优化建议
    列出3条可执行建议：
    1.
    2.
    3.
    要求：
    - 使用中文
    - 简洁专业
    - 像商业报告
    - 如果没有客户或地域数据，跳过对应章节
    """
    try:
        client = OpenAI(api_key=DEEPSEEK_API_KEY, base_url="https://api.deepseek.com")
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[{"role": "user", "content": prompt}]
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"AI分析出错：{e}"


# ============================================================
# 主程序
# ============================================================
st.title("📊 AI销售分析助手")
st.write("上传Excel文件，自动分析销售数据、客户画像、地域分布，AI生成报告")

# 侧边栏
with st.sidebar:
    st.header("📋 使用说明")
    st.markdown("""
    **Excel需包含以下列：**
    - 日期
    - 产品
    - 销量
    - 金额
    - 用户ID（可选，用于客户分析）
    - 配送地址（可选，用于地域分析）

    **系统自动生成：**
    - 销售概况与产品排行
    - 同比环比与需求预测
    - 客户价值分层与消费排行
    - 地域销售分布与热力图
    - AI智能分析报告
    """)
    st.markdown("---")
    st.markdown("💡 可下载测试数据体验")

# 下载测试数据
import os as _os
_test_data_path = _os.path.join(_os.path.dirname(__file__), "销售数据_测试.xlsx") if "__file__" in dir() else "销售数据_测试.xlsx"
if _os.path.exists(_test_data_path):
    with open(_test_data_path, "rb") as f:
        st.sidebar.download_button(
            "📥 下载测试数据",
            f.read(),
            file_name="销售数据_测试.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# 上传文件
file = st.file_uploader("请选择Excel文件", type=["xlsx", "xls"])

if file is None:
    st.info("👆 请上传Excel文件，或下载测试数据体验")
    st.stop()

# 读取数据
try:
    df = pd.read_excel(file)
except Exception as e:
    st.error(f"读取文件出错：{e}")
    st.stop()

# 检查必要列
required_cols = ['日期', '产品', '销量', '金额']
missing = [c for c in required_cols if c not in df.columns]
if missing:
    st.error(f"缺少必要列：{missing}")
    st.stop()

# 检查可选列
has_customer_data = '用户ID' in df.columns
has_geo_data = '配送地址' in df.columns

if has_customer_data:
    st.success("✅ 检测到用户ID列，将生成客户分析")
if has_geo_data:
    st.success("✅ 检测到配送地址列，将生成地域分析")

st.write(f"文件名：{file.name} | 共 {len(df)} 条记录")

# ==============================================
# 基础分析
# ==============================================
total_sales = int(df['金额'].sum())
total_volume = int(df['销量'].sum())
product_sales = df.groupby("产品")["销量"].sum().sort_values(ascending=False)
hot_product = product_sales.index[0]

# ==============================================
# 1. 销售概况
# ==============================================
st.subheader("📊 销售分析报告")

st.write("### 💰 销售概况")
col1, col2, col3 = st.columns(3)
with col1:
    st.metric("总销售额", f"¥{total_sales:,}")
with col2:
    st.metric("总销量", f"{total_volume:,} 件")
with col3:
    st.metric("热销产品", hot_product)

if has_customer_data:
    col4, col5 = st.columns(2)
    with col4:
        st.metric("客户总数", f"{df['用户ID'].nunique()} 人")
    with col5:
        avg_customer_spending = int(total_sales / df['用户ID'].nunique())
        st.metric("人均消费", f"¥{avg_customer_spending:,}")

# ==============================================
# 2. 产品分析
# ==============================================
st.write("### 🏆 产品销售分析")
col1, col2 = st.columns(2)

with col1:
    st.write("📦 销量排行")
    ps_df = product_sales.reset_index()
    fig_sales = px.bar(ps_df, x="产品", y="销量", title="产品销量",
                       color="产品", color_discrete_sequence=CHART_COLORS)
    style_chart(fig_sales)
    safe_write_image(fig_sales, "sales_chart.png")
    st.plotly_chart(fig_sales, use_container_width=True)

with col2:
    st.write("💰 金额排行")
    pm_df = df.groupby("产品")["金额"].sum().reset_index().sort_values(by="金额", ascending=False)
    fig_money = px.bar(pm_df, x="产品", y="金额", title="产品销售额",
                       color="产品", color_discrete_sequence=CHART_COLORS)
    style_chart(fig_money)
    safe_write_image(fig_money, "money_chart.png")
    st.plotly_chart(fig_money, use_container_width=True)

st.write("### 🥧 产品销量占比")
fig2 = px.pie(ps_df, names="产品", values="销量", title="产品销售占比",
              color="产品", color_discrete_sequence=CHART_COLORS)
style_chart(fig2)
safe_write_image(fig2, "pie_chart.png")
st.plotly_chart(fig2, use_container_width=True)

# ==============================================
# 3. 销售趋势
# ==============================================
date_col = '日期'
df[date_col] = pd.to_datetime(df[date_col])
trend = df.groupby(date_col)["销量"].sum().reset_index()

st.write("### 📈 销售趋势")
fig_line = px.line(trend, x=date_col, y="销量", title="销售数量趋势")
style_chart(fig_line)
safe_write_image(fig_line, "line_chart.png")
st.plotly_chart(fig_line, use_container_width=True)

# 按月汇总
df['月份'] = df[date_col].dt.to_period('M')
monthly_sales = df.groupby('月份')['销量'].sum().sort_index()

# ==============================================
# 4. 同比环比分析
# ==============================================
st.write("### 📊 同比环比分析")
forecast_data = None

if len(monthly_sales) >= 2:
    current_month = monthly_sales.iloc[-1]
    last_month = monthly_sales.iloc[-2]
    mom_change = ((current_month - last_month) / last_month * 100)

    col_m1, col_m2, col_m3 = st.columns(3)
    with col_m1:
        st.metric("本月销量", f"{current_month} 件")
    with col_m2:
        st.metric("上月销量", f"{last_month} 件")
    with col_m3:
        st.metric("环比变化", f"{mom_change:+.1f}%")

    yoy_change = None
    last_year_sales = None
    current_idx = monthly_sales.index[-1]
    last_year_idx = pd.Period(f"{current_idx.year - 1}-{current_idx.month:02d}", freq='M')
    if last_year_idx in monthly_sales.index:
        last_year_sales = monthly_sales.loc[last_year_idx]
        yoy_change = ((current_month - last_year_sales) / last_year_sales * 100)
        col_y1, col_y2, col_y3 = st.columns(3)
        with col_y1:
            st.metric("本月销量", f"{current_month} 件")
        with col_y2:
            st.metric("去年同期销量", f"{last_year_sales} 件")
        with col_y3:
            st.metric("同比变化", f"{yoy_change:+.1f}%")
    else:
        st.info("暂无去年同期数据，无法计算同比")

    monthly_df = monthly_sales.reset_index()
    monthly_df['月份'] = monthly_df['月份'].astype(str)
    fig_monthly = px.bar(monthly_df, x='月份', y='销量', title="月度销量对比",
                         color='月份', color_discrete_sequence=CHART_COLORS)
    style_chart(fig_monthly)
    safe_write_image(fig_monthly, "monthly_chart.png")
    st.plotly_chart(fig_monthly, use_container_width=True)

    forecast_data = {
        "环比变化": f"{mom_change:+.1f}%",
        "同比变化": f"{yoy_change:+.1f}%" if yoy_change is not None else "无去年同期数据",
    }
else:
    st.info("数据不足，至少需要两个月的数据才能做环比分析")

# ==============================================
# 5. 需求预测
# ==============================================
st.write("### 🔮 需求预测")
if len(monthly_sales) >= 3:
    x = np.arange(len(monthly_sales))
    y = monthly_sales.values.astype(float)
    coeffs = np.polyfit(x, y, 1)

    future_labels = []
    trend_preds = []
    last_idx = monthly_sales.index[-1]
    for i in range(1, 4):
        future_x = len(monthly_sales) - 1 + i
        pred = max(int(coeffs[0] * future_x + coeffs[1]), 0)
        trend_preds.append(pred)
        future_labels.append(str(last_idx + i))

    ma_pred = int(monthly_sales.iloc[-3:].mean())

    col_p1, col_p2 = st.columns(2)
    with col_p1:
        st.metric("下月预测（趋势线）", f"{trend_preds[0]} 件")
    with col_p2:
        st.metric("下月预测（移动平均）", f"{ma_pred} 件")

    st.write(f"未来3个月趋势预测：{trend_preds[0]} → {trend_preds[1]} → {trend_preds[2]} 件")

    hist_months = [str(idx) for idx in monthly_sales.index]
    hist_values = list(monthly_sales.values)

    fig_forecast = go.Figure()
    fig_forecast.add_trace(go.Scatter(
        x=hist_months, y=hist_values, mode='lines+markers',
        name='历史销量', line=dict(color='#1f77b4', width=2)
    ))
    fig_forecast.add_trace(go.Scatter(
        x=[hist_months[-1]] + future_labels,
        y=[hist_values[-1]] + trend_preds,
        mode='lines+markers', name='趋势预测',
        line=dict(color='#ff7f0e', width=2, dash='dash'),
        marker=dict(size=10)
    ))
    fig_forecast.update_layout(title="销量趋势与预测")
    style_chart(fig_forecast)
    safe_write_image(fig_forecast, "forecast_chart.png")
    st.plotly_chart(fig_forecast, use_container_width=True)

    st.caption("💡 趋势预测基于线性回归，移动平均基于最近3个月数据。仅供参考，实际备货需结合季节性等因素。")

    if forecast_data:
        forecast_data["趋势预测"] = f"{trend_preds[0]} 件"
        forecast_data["移动平均预测"] = f"{ma_pred} 件"
else:
    st.info("数据不足，至少需要三个月的数据才能做需求预测")

# ==============================================
# 6. 客户分析（新增）
# ==============================================
customer_data = None
if has_customer_data:
    st.write("### 👥 客户分析")

    customer_stats = df.groupby('用户ID').agg(
        消费总额=('金额', 'sum'),
        购买次数=('日期', 'count'),
        购买件数=('销量', 'sum')
    ).sort_values('消费总额', ascending=False)

    total_customers = len(customer_stats)
    avg_spending = int(customer_stats['消费总额'].mean())
    avg_order_value = int(df['金额'].mean())

    # KPI
    col_c1, col_c2, col_c3 = st.columns(3)
    with col_c1:
        st.metric("客户总数", f"{total_customers} 人")
    with col_c2:
        st.metric("人均消费", f"¥{avg_spending:,}")
    with col_c3:
        st.metric("客单价", f"¥{avg_order_value:,}")

    # Top 10 客户排行
    st.write("#### 🏆 Top 10 高价值客户")
    top10 = customer_stats.head(10).reset_index()
    fig_top_customer = px.bar(
        top10, x='用户ID', y='消费总额', title="Top 10 客户消费金额",
        color='消费总额', color_continuous_scale='Blues'
    )
    style_chart(fig_top_customer)
    fig_top_customer.update_layout(showlegend=False)
    safe_write_image(fig_top_customer, "customer_top10.png")
    st.plotly_chart(fig_top_customer, use_container_width=True)

    # 客户价值分层
    st.write("#### 📊 客户价值分层")
    top_20_pct = max(int(total_customers * 0.2), 1)
    high_value = customer_stats.head(top_20_pct)
    mid_end = top_20_pct + int(total_customers * 0.3)
    mid_value = customer_stats.iloc[top_20_pct:mid_end]
    low_value = customer_stats.iloc[mid_end:]

    high_revenue = int(high_value['消费总额'].sum())
    mid_revenue = int(mid_value['消费总额'].sum())
    low_revenue = int(low_value['消费总额'].sum())
    high_pct = round(high_revenue / total_sales * 100, 1)

    tier_df = pd.DataFrame({
        '客户层级': ['高价值（前20%）', '中价值（中间30%）', '低价值（后50%）'],
        '客户数': [len(high_value), len(mid_value), len(low_value)],
        '消费总额': [high_revenue, mid_revenue, low_revenue]
    })

    col_t1, col_t2 = st.columns(2)
    with col_t1:
        fig_tier = px.pie(
            tier_df, names='客户层级', values='消费总额',
            title="各层级消费金额占比",
            color='客户层级',
            color_discrete_map={
                '高价值（前20%）': '#E74C3C',
                '中价值（中间30%）': '#F39C12',
                '低价值（后50%）': '#3498DB'
            }
        )
        style_chart(fig_tier)
        safe_write_image(fig_tier, "customer_tier.png")
        st.plotly_chart(fig_tier, use_container_width=True)

    with col_t2:
        st.write("**分层详情：**")
        st.dataframe(tier_df, use_container_width=True, hide_index=True)
        st.info(f"💡 前20%的客户贡献了 **{high_pct}%** 的销售额")

    # 购买频次分布
    st.write("#### 🔄 客户购买频次分布")
    freq_bins = pd.cut(customer_stats['购买次数'], bins=[0, 2, 5, 10, 999],
                       labels=['1-2次', '3-5次', '6-10次', '10次以上'])
    freq_dist = freq_bins.value_counts().sort_index()
    freq_df = pd.DataFrame({'频次区间': freq_dist.index.astype(str), '客户数': freq_dist.values})

    fig_freq = px.bar(freq_df, x='频次区间', y='客户数', title="客户购买频次分布",
                      color='频次区间', color_discrete_sequence=CHART_COLORS)
    style_chart(fig_freq)
    safe_write_image(fig_freq, "customer_freq.png")
    st.plotly_chart(fig_freq, use_container_width=True)

    # 帕累托图（累计收入占比）
    st.write("#### 📉 客户收入贡献帕累托图")
    customer_stats['累计收入'] = customer_stats['消费总额'].cumsum()
    customer_stats['累计占比'] = (customer_stats['累计收入'] / total_sales * 100).round(1)
    customer_stats['排名'] = range(1, len(customer_stats) + 1)

    fig_pareto = make_subplots(specs=[[{"secondary_y": True}]])
    fig_pareto.add_trace(
        go.Bar(x=customer_stats['排名'], y=customer_stats['消费总额'],
               name='客户消费金额', marker_color='#3498DB'),
        secondary_y=False
    )
    fig_pareto.add_trace(
        go.Scatter(x=customer_stats['排名'], y=customer_stats['累计占比'],
                   name='累计占比(%)', line=dict(color='#E74C3C', width=2)),
        secondary_y=True
    )
    fig_pareto.add_hline(y=80, line_dash="dash", line_color="gray",
                         annotation_text="80%线", secondary_y=True)
    fig_pareto.update_layout(title="客户收入贡献帕累托图（二八法则）")
    fig_pareto.update_xaxes(title_text="客户排名")
    fig_pareto.update_yaxes(title_text="消费金额", secondary_y=False)
    fig_pareto.update_yaxes(title_text="累计占比(%)", secondary_y=True)
    style_chart(fig_pareto)
    safe_write_image(fig_pareto, "customer_pareto.png")
    st.plotly_chart(fig_pareto, use_container_width=True)

    top5_str = ", ".join([f"{idx}({val}元)" for idx, val in customer_stats.head(5)['消费总额'].items()])
    freq_dist_str = ", ".join([f"{idx}{val}人" for idx, val in freq_dist.items()])

    customer_data = {
        "客户总数": f"{total_customers}人",
        "人均消费": f"¥{avg_spending}",
        "客单价": f"¥{avg_order_value}",
        "top5客户": top5_str,
        "高价值占比": f"{high_pct}%",
        "频次分布": freq_dist_str
    }
else:
    st.write("### 👥 客户分析")
    st.info("当前数据没有「用户ID」列，无法生成客户分析。请在Excel中添加「用户ID」列体验此功能。")

# ==============================================
# 7. 地域分析（新增）
# ==============================================
geo_data = None
if has_geo_data:
    st.write("### 🗺️ 地域分析")

    df['省份'] = df['配送地址'].apply(extract_province)

    geo_stats = df.groupby('省份').agg(
        销售额=('金额', 'sum'),
        销量=('销量', 'sum'),
        订单数=('日期', 'count'),
        客户数=('用户ID', 'nunique') if has_customer_data else ('日期', 'count')
    ).sort_values('销售额', ascending=False)

    total_regions = len(geo_stats)
    top3_revenue = geo_stats.head(3)['销售额'].sum()
    top3_pct = round(top3_revenue / total_sales * 100, 1)

    # KPI
    col_g1, col_g2 = st.columns(2)
    with col_g1:
        st.metric("覆盖地区数", f"{total_regions} 个")
    with col_g2:
        st.metric("Top3地区销售额占比", f"{top3_pct}%")

    # 各省销售额排行
    st.write("#### 📊 各地区销售额排行")
    geo_df = geo_stats.reset_index()
    fig_geo = px.bar(geo_df, x='省份', y='销售额', title="各地区销售额",
                     color='销售额', color_continuous_scale='Viridis')
    style_chart(fig_geo)
    fig_geo.update_layout(showlegend=False)
    safe_write_image(fig_geo, "geo_sales.png")
    st.plotly_chart(fig_geo, use_container_width=True)

    # 各省销售额 vs 销量
    st.write("#### 📊 各地区销售额与销量对比")
    fig_geo2 = make_subplots(specs=[[{"secondary_y": True}]])
    fig_geo2.add_trace(
        go.Bar(x=geo_df['省份'], y=geo_df['销售额'], name='销售额', marker_color='#3498DB'),
        secondary_y=False
    )
    fig_geo2.add_trace(
        go.Scatter(x=geo_df['省份'], y=geo_df['销量'], name='销量',
                   line=dict(color='#E74C3C', width=2), mode='lines+markers'),
        secondary_y=True
    )
    fig_geo2.update_layout(title="各地区销售额与销量对比")
    fig_geo2.update_xaxes(title_text="省份")
    fig_geo2.update_yaxes(title_text="销售额(元)", secondary_y=False)
    fig_geo2.update_yaxes(title_text="销量(件)", secondary_y=True)
    style_chart(fig_geo2)
    safe_write_image(fig_geo2, "geo_compare.png")
    st.plotly_chart(fig_geo2, use_container_width=True)

    # 地区 × 产品交叉热力图
    st.write("#### 🔥 地区 × 产品销量热力图")
    cross = df.pivot_table(values='销量', index='省份', columns='产品', aggfunc='sum', fill_value=0)
    fig_heat = px.imshow(cross, title="地区×产品销量分布",
                         color_continuous_scale='YlOrRd', aspect='auto')
    fig_heat.update_layout(
        xaxis_title="产品", yaxis_title="省份",
        font=dict(color='black', size=12)
    )
    safe_write_image(fig_heat, "geo_heatmap.png")
    st.plotly_chart(fig_heat, use_container_width=True)

    # 地域明细表
    st.write("#### 📋 地域明细")
    display_geo = geo_stats.copy()
    display_geo.columns = ['销售额', '销量', '订单数', '客户数'] if has_customer_data else ['销售额', '销量', '订单数', '订单数']
    if not has_customer_data:
        display_geo = display_geo.drop(columns=['订单数.1'])
    st.dataframe(display_geo.style.format({'销售额': '¥{:,.0f}', '销量': '{:,.0f}'}),
                 use_container_width=True)

    top5_geo_str = ", ".join([f"{idx}({int(val)}元)" for idx, val in geo_stats.head(5)['销售额'].items()])

    geo_data = {
        "地区数": f"{total_regions}个",
        "top5地区": top5_geo_str,
        "地区集中度": f"{top3_pct}%"
    }
else:
    st.write("### 🗺️ 地域分析")
    st.info("当前数据没有「配送地址」列，无法生成地域分析。请在Excel中添加「配送地址」列体验此功能。")

# ==============================================
# 8. AI 分析报告
# ==============================================
st.write("### 🤖 AI分析报告")
with st.spinner("AI正在生成分析报告..."):
    ai_report = get_ai_report(
        {"总销售额": f"¥{total_sales:,}", "总销量": f"{total_volume:,}件", "热销产品": hot_product},
        forecast_data, customer_data, geo_data
    )

with st.container(border=True):
    st.markdown(ai_report)

# ==============================================
# 9. 导出Word报告
# ==============================================
def create_word_report(text):
    doc = Document()
    doc.add_heading("销售分析报告", level=1)
    doc.add_heading("📊 销售数据图表", level=2)

    all_images = [
        "sales_chart.png", "money_chart.png", "pie_chart.png",
        "line_chart.png", "monthly_chart.png", "forecast_chart.png",
        "customer_top10.png", "customer_tier.png", "customer_freq.png",
        "customer_pareto.png", "geo_sales.png", "geo_compare.png", "geo_heatmap.png"
    ]
    for img in all_images:
        try:
            doc.add_picture(img, width=Inches(5))
        except Exception:
            pass

    doc.add_heading("📝 AI分析内容", level=2)
    for line in text.split("\n"):
        if line.startswith("##"):
            doc.add_heading(line.replace("#", "").strip(), level=2)
        elif line.strip():
            line = line.replace("**", "")
            doc.add_paragraph(line)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


word_file = create_word_report(ai_report)
st.download_button(
    label="📄 导出Word报告",
    data=word_file,
    file_name="销售分析报告.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)
