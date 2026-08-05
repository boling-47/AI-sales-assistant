def calculator(expression):
    result = eval(expression)
    return result
expenses = []

import json
from datetime import datetime

def add_expense(money, note):

    expense = {
        "type": "expense",
        "money": money,
        "note": note,
        "date": datetime.now().strftime("%Y-%m-%d")
    }

    with open("expenses.json", "a", encoding="utf-8") as f:
        f.write(json.dumps(expense, ensure_ascii=False) + "\n")
        f.write(json.dumps(expense, ensure_ascii=False) + "\n")

    return "记录成功"


def get_expenses():

    total = 0

    with open("expenses.json", "r", encoding="utf-8") as f:
        for line in f:
            item = json.loads(line)
            total += item["money"]

    return f"总消费：{total}元"

def get_expenses():

    total = 0
    records = []

    with open("memory.json", "r", encoding="utf-8") as f:
        for line in f:
            try:
                item = json.loads(line)

                if "money" in item:
                    total += item["money"]
                    records.append(item)

            except:
                pass


    return {
        "total": total,
        "records": records
    }
def add_income(money, note):

    income = {
        "type":"income",
        "money":money,
        "note":note,
        "date":datetime.now().strftime("%Y-%m-%d")
    }

    with open("memory.json","a",encoding="utf-8") as f:
        f.write(
            json.dumps(
                income,
                ensure_ascii=False
            )+"\n"
        )

    return "收入记录成功"

def get_balance():

    balance = 0

    with open("memory.json","r",encoding="utf-8") as f:

        for line in f:

            if line.strip():

                item=json.loads(line)

                if item.get("type")=="income":
                    balance += item["money"]

                elif item.get("type")=="expense":
                    balance -= item["money"]

    return f"你的余额是{balance}元"

from analysis_tools import analyze_sales
sales_tool = {
    "type": "function",
    "function": {
        "name": "analyze_sales",
        "description": "分析Excel销售数据，包括总销售额、总销量和热销产品",
        "parameters": {
            "type": "object",
            "properties": {}
        }
    }
}

tools = [
    sales_tool
]