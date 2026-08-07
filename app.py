import streamlit as st
from openai import OpenAI
import os
from dotenv import load_dotenv
import pandas as pd
import plotly.express as px
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
        yaxis=dict(tickfont=dict(color='black', size=12), title_font=dict(color='black', size=13))
    )
    return fig


CHART_COLORS = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf']


def get_ai_report(data):
    prompt = f"""
    你是一名销售分析师。
    根据以下销售数据生成分析报告：
    销售额：{data["总销售额"]}
    销量：{data["总销量"]}
    热销产品：{data["热销产品"]}
    请严格按照下面格式输出：
    ## 📌 销售总结
    总结当前销售情况。
    ## 🔍 核心发现
    列出3条数据发现：
    1.
    2.
    3.
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
        ai_report = get_ai_report(result)

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

        st.write("### 📈 销售趋势")
        date_columns = [
            "日期",
            "时间",
            "date",
            "Date"
        ]
        date_col = None
        for col in date_columns:
            if col in df.columns:
                date_col = col
                break

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
        else:
            st.info("当前Excel没有日期数据，暂无法生成销售趋势图")

        st.markdown("### 🤖 AI分析建议")
        with st.container(border=True):
            st.markdown(ai_report)

        def create_word_report(text):
            doc = Document()
            doc.add_heading("销售分析报告", level=1)
            doc.add_heading("📊 销售数据图表", level=2)
            for img in ["sales_chart.png", "money_chart.png", "pie_chart.png", "line_chart.png"]:
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
