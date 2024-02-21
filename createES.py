import Status
import Company
from dotenv import load_dotenv

load_dotenv()

from langchain_google_genai import ChatGoogleGenerativeAI

llm = ChatGoogleGenerativeAI(model='gemini-pro', temperature=0)

status = Status.Status()
company = Company.Company()

import openpyxl

wb = openpyxl.load_workbook(r"C:\Users\bothm\dev_Python\createES_RPA\create_entry_sheet.xlsm")
company_sheet = wb['createES']
status_sheet = wb['Status']

company.setName(company_sheet['C4'].value)
company.setJob(company_sheet['C8'].value)
status.setCharactorLimit(company_sheet['C12'].value)
status.setStrength(status_sheet['C4'].value)
status.setExperience(status_sheet['C8'].value, status_sheet['C12'].value)

strengthPrompt = status.getStrength(0)
for strengthNum in range(1, status.getStrengthNum()):
    strengthPrompt += "、" + status.getStrength(strengthNum)

experienceKeys = list(status.experience.keys())
experiencePrompt = ""
for experienceNum in (0, len(experienceKeys) - 1):
    experiencePrompt += experienceKeys[experienceNum] + "では" + status.experience.get(experienceKeys[experienceNum]) + "を学びました。"

firstPrompt = "あなたは素晴らしいAIです。次の質問に日本語で答えてください。"
mainPrompt = (
    company.getName()
    + "のHPを参照し"
    + company.getJob()
    + "職の志望動機を"
    + str(status.getCharactorLimit())
    + "以内で回答してください。会社名を言うときは、「貴社」という言葉を使用してください。また、私の強みは"
    + strengthPrompt
    + "です。また、経験として"
    + experiencePrompt
    + "これらの強みと経験が、企業の強みとどのようにマッチングするかを明確にしてください。なお、回答の文字数は"
    + str(status.getCharactorLimit())
    + "字の9割以上にする必要があり、"
    + str(status.getCharactorLimit())
    + "字を絶対に超えてはいけません。"
)

result = llm.invoke(str(mainPrompt))

# saveメソッドで同じ名前で上書き保存
company_sheet['J4'] = result.content.replace("\n", "")
company_sheet['J4'].alignment = openpyxl.styles.Alignment(wrapText=True)
wb.save(r"C:\Users\bothm\dev_Python\createES_RPA\{}.xlsx".format(company_sheet['C4'].value))