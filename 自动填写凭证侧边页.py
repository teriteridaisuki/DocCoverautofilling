from win32com.client import Dispatch
import openpyxl


def typing(box, text):
    doc.Shapes.Range(box).Select()
    s.Text = text


def cleaning():
    typing("Text Box 13", "")
    typing("Text Box 2", "")
    typing("Text Box 1", "")
    typing("Text Box 7", "")
    typing("Text Box 8", "")
    typing("Text Box 11", "")


def testsaving():
    if ws.cell(2 + i, 7).value != ws.cell(3 + i, 7).value:
        doc.SaveAs('E:\\untitle\\输出\\凭证侧边第' + str(ws.cell(2 + i, 7).value) + '本.docx')
        global centernum
        centernum = 1
        cleaning()


ws = openpyxl.load_workbook('凭证编号.xlsx')["Sheet1"]
path = r'E:\untitle\备用程序\凭证侧边与封面\凭证侧边打印版.docx'
app = Dispatch('Word.Application')
# 新建word文档
doc = app.Documents.Open(FileName=path)
app.Visible = False
s = app.Selection
cleaning()
centernum = 1  # 表示这本凭证里是第几个利润中心
for i in range(ws.max_row):
    try:
        typing("Text Box 11", ws.cell(2 + i, 7).value)
        typing("Text Box 1", ws.cell(2 + i, 1).value[1:5])
        if centernum == 1:
            typing("Text Box 13", ws.cell(2 + i, 1).value)
            centernum += 1
            print(ws.cell(2 + i, 1).value, centernum, ws.cell(2 + i, 7).value)
            testsaving()
            continue
        if centernum == 2:
            typing("Text Box 2", ws.cell(2 + i, 1).value)
            centernum += 1
            print(ws.cell(2 + i, 1).value, centernum, ws.cell(2 + i, 7).value)
            testsaving()
            continue
        if centernum == 3:
            typing("Text Box 7", ws.cell(2 + i, 1).value)
            centernum += 1
            print(ws.cell(2 + i, 1).value, centernum, ws.cell(2 + i, 7).value)
            testsaving()
            continue
        if centernum == 4:
            typing("Text Box 8", ws.cell(2 + i, 1).value)
            centernum += 1
            print(ws.cell(2 + i, 1).value, centernum, ws.cell(2 + i, 7).value)
            testsaving()
            continue

    except:
        pass
doc.Close()
app.Quit()

'''
    ActiveDocument.Shapes.Range(Array("Text Box 15")).Select
    Selection.WholeStory
    Selection.TypeText Text:="2021"
    ActiveDocument.Shapes.Range(Array("Text Box 14")).Select
    Selection.WholeStory
    Selection.TypeText Text:="08"
    ActiveDocument.Shapes.Range(Array("Text Box 13")).Select
    Selection.WholeStory
    Selection.TypeText Text:="L80030003"
    ActiveDocument.Shapes.Range(Array("Text Box 2")).Select
    Selection.WholeStory
    Selection.TypeText Text:="L80020002"
    ActiveDocument.Shapes.Range(Array("Text Box 7")).Select
    Selection.WholeStory
    Selection.TypeText Text:="L80010001"
    ActiveDocument.Shapes.Range(Array("Text Box 8")).Select
    Selection.WholeStory
    Selection.TypeText Text:="L8801880"
    ActiveDocument.Shapes.Range(Array("Text Box 11")).Select
    Selection.WholeStory
    Selection.TypeText Text:="12"
    ActiveDocument.Save
'''
