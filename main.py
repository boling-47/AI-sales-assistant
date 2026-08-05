from openai import OpenAI
import os
import json
from dotenv import load_dotenv
from tools import (
    add_expense,
    get_expenses,
    add_income,
    get_balance
)
from analysis_tools import analyze_sales
load_dotenv()


client = OpenAI(
    api_key=os.getenv("DEEPSEEK_API_KEY"),
    base_url="https://api.deepseek.com"
)

tools = [
    {
        "type": "function",
        "function": {
            "name": "add_expense",
            
        }
    },

    {
        "type": "function",
        "function": {
            "name": "get_expenses",
            
        }
    },

    {
        "type": "function",
        "function": {
            "name": "add_income",
            "description": "记录收入，例如工资到账",
            "parameters": {
                "type": "object",
                "properties": {
                    "money": {
                        "type": "number"
                    },
                    "note": {
                        "type": "string"
                    }
                },
                "required": [
                    "money",
                    "note"
                ]
            }
        }
    }
    ,
{
    "type": "function",
    "function": {
        "name": "analyze_sales",
        "description": "分析销售数据文件data.xlsx，包括总销售额、总销量和热销产品",
        "parameters": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Excel文件路径，例如 data.xlsx"
                }
            },
            "required": [
                "file_path"
            ]
        }
    }
}
]

class Agent:

    def __init__(self):
        self.messages = []


    def add_message(self, role, content):
        self.messages.append(
            {
                "role": role,
                "content": content
            }
        )
    def think(self):

        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=self.messages,
            tools=tools
        )

        answer = response.choices[0].message.content

        message = response.choices[0].message

        if message.tool_calls:

            self.messages.append(
                {
                    "role": "assistant",
                    "tool_calls": message.tool_calls
                }
            )

            tool = message.tool_calls[0]
            if tool.function.name == "add_expense":

                args = json.loads(tool.function.arguments)
                print(args)

                result = add_expense(
                    args["money"],
                    args["note"]
                )


            elif tool.function.name == "get_expenses":

                result = get_expenses()


            elif tool.function.name == "add_income":

                args = json.loads(tool.function.arguments)

                result = add_income(
                    args["money"],
                    args["note"]
                )


            elif tool.function.name == "get_balance":

                result = get_balance()
            elif tool.function.name == "analyze_sales":

                args = json.loads(tool.function.arguments)

                result = analyze_sales(
                args["file_path"]
                )
            elif tool.function.name == "get_expenses":

                result = get_expenses()
            self.messages.append(
                {
                    "role": "tool",
                    "tool_call_id": tool.id,
                    "content": json.dumps(result, ensure_ascii=False)
                }
            )

            return self.think()
    
        else:
            answer = message.content
            self.add_message(
                "assistant",
                answer
            )

            return answer
agent = Agent()

while True:
    user = input("你：")

    agent.add_message(
        "user",
        user
    )

    answer = agent.think()

    print("AI:", answer)