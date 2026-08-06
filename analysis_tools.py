import pandas as pd


def analyze_sales(file_path):

    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        return {
            "错误": str(e)
        }

    # 自动寻找金额列
    money_columns = ["金额", "销售额", "销售金额", "amount", "money"]

    sales_columns = ["销量", "数量", "销售数量", "sales", "quantity"]

    product_columns = ["产品", "商品", "产品名称", "product"]

    money_col = None
    sales_col = None
    product_col = None


    def find_column(columns):
        for col in columns:
            if col in df.columns:
                return col
        return None


    money_col = find_column(money_columns)
    sales_col = find_column(sales_columns)
    product_col = find_column(product_columns)


    result = {}


    if money_col:
        result["总销售额"] = int(df[money_col].sum())
    else:
        result["总销售额"] = "未找到"


    if sales_col:
        result["总销量"] = int(df[sales_col].sum())
    else:
        result["总销量"] = "未找到"


    if product_col and sales_col:

        best_product = df.groupby(product_col)[sales_col].sum().idxmax()

        result["热销产品"] = best_product

    else:

        result["热销产品"] = "未找到"

    return result
