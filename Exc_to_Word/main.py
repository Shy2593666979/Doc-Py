import pandas as pd
from docx import Document

excel_file_path = "嵌入式二班基本信息表.xlsx"

word_file_path = "XXX2128X24XXX.docx"

your_class = "专业年级：21级嵌入式开发二班"

def deal_Excel():
    df = pd.read_excel(excel_file_path)
 
    name_list = []
    id_list = []
    for name in df["姓名"]:
        name_list.append(name)

    for id in df["学号"]:
        id_list.append(id)

    # 班级花名册的姓名、学号读取
    deal_Word(name_list,id_list)
    
def deal_Word(name_list:list, id_list:list):
    #读取word文档
    doc = Document(word_file_path)
    for i in range(0,len(name_list)):
        save_path = name_list[i] + str(id_list[i]) + ".docx"
        new_doc = doc

        new_doc.paragraphs[2].text = your_class + "   学号：" + str(id_list[i]) + "   姓名：" + name_list[i]

        new_doc.save(save_path)


if __name__ == "__main__":
    deal_Excel()
