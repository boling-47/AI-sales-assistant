import pandas as pd


def analyze_sales(file_path):

    df = pd.read_excel(file_path)

    total_money = df["金额"].sum()

    total_sales = df["销量"].sum()

    best_product = (
        df.groupby("产品")["销量"]
        .sum()
        .idxmax()
    )

    return {
    "总销售额": int(total_money),
    "总销量": int(total_sales),
    "热销产品": best_product
    }

