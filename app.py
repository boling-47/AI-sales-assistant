import streamlit as st
from openai import OpenAI
import os
from dotenv import load_dotenv
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from docx import Document
from docx.shared import Inches
from io import BytesIO

load_dotenv()

client = OpenAI(
    api_key=os.getenv("DEEPSEEK_API_KEY"),
    base_url="https://api.deepseek.com"
)


def safe_write_image(fig, filename):
    """安全导出图片，如果 kaleido 不可用就跳过，不影响页面显示"""
    try:
        fig.write_image(filename)
    except Exception:
        pass


def style_chart(fig):
    """设置图表样式：白色背景、黑色文字"""
    fig.update_layout(
        paper_bgcolor='white',
        plot_bgcolor='white',
        font=dict(color='black', size=14),
        xaxis=dict(tickfont=dict(color='black', size=12), title_font=dict(color='black', size=13)),
        yaxis=dict(tickfont=dict(color='black', size=12), title_font=dict(color='black', size=13)),
        legend=dict(font=dict(color='black', size=12))
    )
    return fig


CHART_COLORS = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf']


def get_ai_report(data, forecast_data=None):
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
    prompt += """
    请严格按照下面格式输出：
    ## 📌 销售总结
    总结当前销售情况。
    ## 🔍 核心发现
    列出3条数据发现：
    1.
    2.
    3.
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
    """
    response = client.chat.completions.create(
        model="deepseek-chat",
        messages=[
            {
                "role": "user",
                "content": prompt
            }
        ]
    )
    return response.choices[0].message.content


st.title("📊 AI销售分析助手")
st.write("上传Excel文件，AI帮你分析销售数据")

file = st.file_uploader(
    "请选择Excel文件",
    type=["xlsx"]
)

if file:
    st.success("文件上传成功！")
    st.write("文件名：", file.name)

    if st.button("开始分析"):
        with open("temp.xlsx", "wb") as f:
            f.write(file.getbuffer())

        from analysis_tools import analyze_sales
        result = analyze_sales("temp.xlsx")

        df = pd.read_excel("temp.xlsx")

        st.subheader("📊 销售分析报告")

        st.write("### 💰 销售概况")
        st.metric(
            "总销售额",
            f"¥{result['总销售额']}"
        )
        st.metric(
            "总销量",
            f"{result['总销量']} 件"
        )

        st.write("### 🏆 热销产品")
        st.success(
            f"当前热销产品是：{result['热销产品']}"
        )

        st.write("### 🏆 产品销售分析")
        col1, col2 = st.columns(2)

        # 左边：销量排行
        with col1:
            st.write("📦 销量排行")
            product_sales = (
                df.groupby("产品")["销量"]
                .sum()
                .reset_index()
                .sort_values(
                    by="销量",
                    ascending=False
                )
            )
            fig_sales = px.bar(
                product_sales,
                x="产品",
                y="销量",
                title="产品销量",
                color="产品",
                color_discrete_sequence=CHART_COLORS
            )
            style_chart(fig_sales)
            safe_write_image(fig_sales, "sales_chart.png")
            st.plotly_chart(
                fig_sales,
                use_container_width=True
            )

        # 右边：金额排行
        with col2:
            st.write("💰 金额排行")
            product_money = (
                df.groupby("产品")["金额"]
                .sum()
                .reset_index()
                .sort_values(
                    by="金额",
                    ascending=False
                )
            )
            fig_money = px.bar(
                product_money,
                x="产品",
                y="金额",
                title="产品销售额",
                color="产品",
                color_discrete_sequence=CHART_COLORS
            )
            style_chart(fig_money)
            safe_write_image(fig_money, "money_chart.png")
            st.plotly_chart(
                fig_money,
                use_container_width=True
            )

        st.write("### 🥧 产品销量占比")
        fig2 = px.pie(
            product_sales,
            names="产品",
            values="销量",
            title="产品销售占比",
            color="产品",
            color_discrete_sequence=CHART_COLORS
        )
        style_chart(fig2)
        safe_write_image(fig2, "pie_chart.png")
        st.plotly_chart(fig2, use_container_width=True)

        # ====== 日期处理 ======
        st.write("### 📈 销售趋势")
        date_columns = ["日期", "时间", "date", "Date"]
        date_col = None
        for col in date_columns:
            if col in df.columns:
                date_col = col
                break

        monthly_sales = None
        forecast_data = None

        if date_col:
            df[date_col] = pd.to_datetime(df[date_col])
            trend = (
                df.groupby(date_col)["销量"]
                .sum()
                .reset_index()
            )
            fig_line = px.line(
                trend,
                x=date_col,
                y="销量",
                title="销售数量趋势"
            )
            style_chart(fig_line)
            safe_write_image(fig_line, "line_chart.png")
            st.plotly_chart(fig_line, use_container_width=True)

            # 按月汇总
            df['月份'] = df[date_col].dt.to_period('M')
            monthly_sales = df.groupby('月份')['销量'].sum().sort_index()
        else:
            st.info("当前Excel没有日期数据，暂无法生成销售趋势图")

        # ====== 同比环比分析 ======
        st.write("### 📊 同比环比分析")
        if monthly_sales is not None and len(monthly_sales) >= 2:
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

            # 同比（检查是否有去年同期数据）
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

            # 月度对比图
            monthly_df = monthly_sales.reset_index()
            monthly_df['月份'] = monthly_df['月份'].astype(str)
            fig_monthly = px.bar(
                monthly_df,
                x='月份',
                y='销量',
                title="月度销量对比",
                color='月份',
                color_discrete_sequence=CHART_COLORS
            )
            style_chart(fig_monthly)
            safe_write_image(fig_monthly, "monthly_chart.png")
            st.plotly_chart(fig_monthly, use_container_width=True)

            forecast_data = {
                "环比变化": f"{mom_change:+.1f}%",
                "同比变化": f"{yoy_change:+.1f}%" if yoy_change is not None else "无去年同期数据",
            }
        elif monthly_sales is not None:
            st.info("数据不足，至少需要两个月的数据才能做环比分析")
        else:
            st.info("当前Excel没有日期数据，暂无法做同比环比分析")

        # ====== 需求预测 ======
        st.write("### 🔮 需求预测")
        if monthly_sales is not None and len(monthly_sales) >= 3:
            # 线性回归预测
            x = np.arange(len(monthly_sales))
            y = monthly_sales.values.astype(float)
            coeffs = np.polyfit(x, y, 1)

            # 预测未来3个月
            future_labels = []
            trend_preds = []
            last_idx = monthly_sales.index[-1]
            for i in range(1, 4):
                future_x = len(monthly_sales) - 1 + i
                pred = int(coeffs[0] * future_x + coeffs[1])
                pred = max(pred, 0)
                trend_preds.append(pred)
                next_period = last_idx + i
                future_labels.append(str(next_period))

            # 移动平均预测（最近3个月）
            ma_pred = int(monthly_sales.iloc[-3:].mean())

            col_p1, col_p2 = st.columns(2)
            with col_p1:
                st.metric("下月预测（趋势线）", f"{trend_preds[0]} 件")
            with col_p2:
                st.metric("下月预测（移动平均）", f"{ma_pred} 件")

            st.write(f"未来3个月趋势预测：{trend_preds[0]} → {trend_preds[1]} → {trend_preds[2]} 件")

            # 预测趋势图（历史 + 预测）
            hist_months = [str(idx) for idx in monthly_sales.index]
            hist_values = list(monthly_sales.values)

            fig_forecast = go.Figure()
            # 历史销量
            fig_forecast.add_trace(go.Scatter(
                x=hist_months,
                y=hist_values,
                mode='lines+markers',
                name='历史销量',
                line=dict(color='#1f77b4', width=2)
            ))
            # 趋势预测（从最后一个历史点延伸）
            fig_forecast.add_trace(go.Scatter(
                x=[hist_months[-1]] + future_labels,
                y=[hist_values[-1]] + trend_preds,
                mode='lines+markers',
                name='趋势预测',
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
        elif monthly_sales is not None:
            st.info("数据不足，至少需要三个月的数据才能做需求预测")
        else:
            st.info("当前Excel没有日期数据，暂无法做需求预测")

        # ====== AI 分析建议 ======
        ai_report = get_ai_report(result, forecast_data)

        st.markdown("### 🤖 AI分析建议")
        with st.container(border=True):
            st.markdown(ai_report)

        def create_word_report(text):
            doc = Document()
            doc.add_heading("销售分析报告", level=1)
            doc.add_heading("📊 销售数据图表", level=2)
            for img in ["sales_chart.png", "money_chart.png", "pie_chart.png", "line_chart.png", "monthly_chart.png", "forecast_chart.png"]:
                try:
                    doc.add_picture(img, width=Inches(5))
                except Exception:
                    pass
            for line in text.split("\n"):
                if line.startswith("##"):
                    doc.add_heading(
                        line.replace("#", "").strip(),
                        level=2
                    )
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
