#!/usr/bin/python
# -- coding: utf-8 --

import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog
import openpyxl
from openai import OpenAI
import threading


class TestCaseGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AI测试用例生成器")
        self.client = OpenAI(
            api_key="sk-dkhpejkynrjezgudanlggabdwhcejpngknccxnfirpdflrbs",
            base_url="https://api.siliconflow.cn/v1"
        )

        # 创建界面组件
        self.create_widgets()

    def create_widgets(self):
        # 输入区域
        input_frame = ttk.LabelFrame(self.root, text="需求输入")
        input_frame.pack(padx=10, pady=5, fill="both", expand=True)

        self.txt_input = scrolledtext.ScrolledText(input_frame, height=10)
        self.txt_input.pack(padx=5, pady=5, fill="both", expand=True)

        # 按钮区域
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(pady=5)

        self.btn_generate = ttk.Button(btn_frame, text="生成测试用例", command=self.start_generation)
        self.btn_generate.pack(side=tk.LEFT, padx=5)

        self.btn_export = ttk.Button(btn_frame, text="导出Excel", command=self.export_excel, state=tk.DISABLED)
        self.btn_export.pack(side=tk.LEFT, padx=5)

        # 状态栏
        self.status = ttk.Label(self.root, text="就绪", relief=tk.SUNKEN)
        self.status.pack(side=tk.BOTTOM, fill=tk.X)

        # 存储测试用例
        self.test_cases = []

    def start_generation(self):
        self.btn_generate.config(state=tk.DISABLED)
        self.btn_export.config(state=tk.DISABLED)
        self.status.config(text="正在生成测试用例...")

        # 使用线程避免界面冻结
        threading.Thread(target=self.generate_test_cases).start()

    def generate_test_cases(self):
        try:
            requirement = self.txt_input.get("1.0", tk.END).strip()
            if not requirement:
                raise ValueError("请输入需求内容")

            response = self.client.chat.completions.create(
                model="deepseek-ai/DeepSeek-V3",
                messages=[{
                    'role': 'user',
                    'content': """请根据以下需求文档生成详细的测试用例，每个功能点都需要详细的用例，要求包含以下要素：
                    1. 测试用例编号
                    2. 测试场景
                    3. 前置条件
                    4. 测试步骤
                    5. 预期结果
                    6. 优先级（高/中/低）

                    需求文档内容：
                    {requirement}

                    请用以下格式返回：
                    ### 测试用例1
                    - 场景：[具体场景]
                    - 前置条件：[条件]
                    - 步骤：1. [步骤1] 2. [步骤2]
                    - 预期结果：[结果]
                    - 优先级：[优先级]"""
                }],
                stream=True
            )

            full_response = ""
            for chunk in response:
                content = chunk.choices[0].delta.content or ""
                full_response += content

            self.parse_response(full_response)
            self.status.config(text="生成完成！")
            self.btn_export.config(state=tk.NORMAL)
            print(full_response)
        except Exception as e:
            self.status.config(text=f"错误：{str(e)}")
        finally:
            self.btn_generate.config(state=tk.NORMAL)

    def parse_response(self, response):
        self.test_cases = []
        current_case = {}

        for line in response.split('\n'):
            line = line.strip()
            if line.startswith("###"):
                if current_case:
                    self.test_cases.append(current_case)
                    current_case = {}
                    print(f"当前测试用例: {current_case}")
                current_case["编号"] = line.split()[-1]
            elif line.startswith("-"):
                # 检查并提取各个字段
                if "场景：" in line:
                    current_case["场景"] = line.split("场景：", 1)[1].strip()
                elif "前置条件：" in line:
                    current_case["前置条件"] = line.split("前置条件：", 1)[1].strip()
                elif "步骤：" in line:
                    current_case["步骤"] = line.split("步骤：", 1)[1].strip()
                elif "预期结果：" in line:
                    current_case["预期结果"] = line.split("预期结果：", 1)[1].strip()
                elif "优先级：" in line:
                    current_case["优先级"] = line.split("优先级：", 1)[1].strip()

        if current_case:
            self.test_cases.append(current_case)

    def export_excel(self):
        try:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel文件", "*.xlsx")]
            )

            if not file_path:
                return

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "测试用例"

            # 写入表头
            headers = ["编号", "场景", "前置条件", "步骤", "预期结果", "优先级"]
            ws.append(headers)

            # 写入数据
            for case in self.test_cases:
                row = [
                    case.get("编号", ""),
                    case.get("场景", ""),
                    case.get("前置条件", ""),
                    case.get("步骤", ""),
                    case.get("预期结果", ""),
                    case.get("优先级", "")
                ]
                ws.append(row)

            # 调整列宽
            for col in ws.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                ws.column_dimensions[column].width = adjusted_width

            wb.save(file_path)
            self.status.config(text=f"文件已保存至：{file_path}")

        except Exception as e:
            self.status.config(text=f"导出失败：{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = TestCaseGeneratorApp(root)
    root.geometry("600x400")
    root.mainloop()
