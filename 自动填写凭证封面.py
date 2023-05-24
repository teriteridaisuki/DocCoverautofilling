from win32com.client import Dispatch
import openpyxl

def typing(box,text):
    doc.Shapes.Range(box).Select()
    s.Text=text
def cleaning():
    typing("Text Box 4", "")
    typing("Text Box 11", "")
    typing("Text Box 12", "")
    typing("Text Box 13", "")
    typing("Text Box 14", "")
    typing("Text Box 15", "")
    typing("Text Box 16", "")
    typing("Text Box 17", "")
    typing("Text Box 18", "")
    typing("Text Box 19", "")
    typing("Text Box 20", "")
    typing("Text Box 21", "")
    typing("Text Box 22", "")
def testsaving():
    if ws.cell(2 + i, 7).value != ws.cell(3 + i, 7).value:
        doc.SaveAs('E:\\untitle\\输出\\凭证封面第' + str(ws.cell(2 + i, 7).value) + '本.docx')
        global centernum
        centernum = 1
        cleaning()
ws = openpyxl.load_workbook('凭证编号.xlsx')["Sheet1"]
path=r'E:\untitle\备用程序\凭证侧边与封面\凭证封面打印版.docx'
app = Dispatch('Word.Application')
# 新建word文档
doc = app.Documents.Open(FileName=path)
app.Visible = False
s=app.Selection
cleaning()
centernum = 1  # 表示这本凭证里是第几个利润中心
for i in range(ws.max_row):
    try:
        typing("Text Box 4", ws.cell(2+i,7).value)#现在是第几本
        if centernum==1:
            typing("Text Box 11", ws.cell(2+i,1).value+" "+ws.cell(2+i,2).value)
            typing("Text Box 15", ws.cell(2 + i, 4).value)
            typing("Text Box 19", ws.cell(2 + i, 5).value)
            centernum += 1
            print(ws.cell(2+i,1).value,centernum,ws.cell(2+i,7).value)
            testsaving()
            continue
        if centernum==2:
            typing("Text Box 12", ws.cell(2+i,1).value+" "+ws.cell(2+i,2).value)
            typing("Text Box 16", ws.cell(2 + i, 4).value)
            typing("Text Box 20", ws.cell(2 + i, 5).value)
            centernum += 1
            print(ws.cell(2+i,1).value,centernum,ws.cell(2+i,7).value)
            testsaving()
            continue
        if centernum==3:
            typing("Text Box 13", ws.cell(2 + i, 1).value + " " + ws.cell(2 + i, 2).value)
            typing("Text Box 17", ws.cell(2 + i, 4).value)
            typing("Text Box 21", ws.cell(2 + i, 5).value)
            centernum += 1
            print(ws.cell(2+i,1).value,centernum,ws.cell(2+i,7).value)
            testsaving()
            continue
        if centernum==4:
            typing("Text Box 14", ws.cell(2 + i, 1).value + " " + ws.cell(2 + i, 2).value)
            typing("Text Box 18", ws.cell(2 + i, 4).value)
            typing("Text Box 22", ws.cell(2 + i, 5).value)
            centernum += 1
            print(ws.cell(2+i,1).value,centernum,ws.cell(2+i,7).value)
            testsaving()
            continue

    except:
        pass
doc.Close()
app.Quit()
