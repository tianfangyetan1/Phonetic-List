
import win32com.client as win32

def insert_pinyin_text(doc, p, text, pinyin):
  # 使用Word的拼音指南功能
  range = doc.Paragraphs(p + 1).Range.Characters.Last
  range.Font.Name = 'MiSans'
  range.Font.Size = 13
  range.Text = text
  range.Select()
  word.Selection.Range.LanguageID = win32.constants.wdSimplifiedChinese
  word.Selection.Range.NoProofing = True
  word.Selection.Range.PhoneticGuide(
    Text = pinyin,
    Alignment = win32.constants.wdPhoneticGuideAlignmentCenter,
    Raise = 0, # 这个参数有bug用不了
    FontSize = 10,
    FontName = 'MiSans'
  )

def insert_last(doc, p, t):
  doc.Paragraphs(p + 1).Range.Characters.Last.InsertAfter(t)

# 启动Word应用程序并新建文档
word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = True  # 让Word可见
doc = word.Documents.Add()

# 打开markdown文件
with open('assets.md', 'r', encoding='utf-8') as f:
  content = f.readlines()

insert_pinyin_text(doc, 0, '字', 'zì')  
doc.Content.Paragraphs.Add()

# 调用函数插入含有拼音的文字
for index, l in zip(range(1, len(content) + 1), content):
  skip = 0
  l2 = list(l)
  for i in range(len(l) - 1):
    c = l2[i]
    if skip > 0:
      skip -= 1
      continue
    if l2[i + 1] == '(':
      pinyin = ''.join(l2[l2.index('(') + 1:l2.index(')')])
      insert_pinyin_text(doc, index, c, pinyin)
      skip = len(pinyin) + 2
      l2[l2.index('(')] = ' '
      l2[l2.index(')')] = ' '
    else:
      insert_last(doc, index, c)

  # 添加新段落
  doc.Content.Paragraphs.Add()

# 保存并关闭文档
doc.SaveAs('output.docx')   # 文件保存在 `C:\Users\{用户名}\Documents` 下
doc.Close()
word.Quit()
