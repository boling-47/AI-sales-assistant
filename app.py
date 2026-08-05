import streamlit as st
from openai import OpenAI
import os
from dotenv import load_dotenv

load_dotenv()

client = OpenAI(
    api_key=os.getenv("DEEPSEEK_API_KEY"),
    base_url="https://api.deepseek.com"
)

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


        st.write("### 🤖 AI分析建议")


        advice = f"""
        根据当前销售数据：

        1. 总销售额达到 ¥{result['总销售额']}。
        2. 总销量为 {result['总销量']} 件。
        3. {result['热销产品']} 是目前销售表现最好的产品。

        建议：
        - 可以重点推广 {result['热销产品']}。
        - 增加该产品库存准备。
        - 后续可以继续分析销售趋势和利润情况。
        """

        st.info(advice)