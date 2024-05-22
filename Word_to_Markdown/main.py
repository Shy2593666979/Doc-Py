import time
import mammoth
import markdownify
# 转存Word文档内的图片
def convert_imgs(image):
    with image.open() as image_bytes:
        file_suffix = image.content_type.split("/")[1]
        path_file = "./img/{}.{}".format(str(time.time()),file_suffix)
        with open(path_file, 'wb') as f:
            f.write(image_bytes.read())


    return {"src":path_file}


# 读取Word文件
with open(r"C:\xxx\xxx\Desktop\XXXX.docx", "rb") as docx_file:
    # 转化Word文档为HTML
    result = mammoth.convert_to_html(docx_file,convert_image = mammoth.images.img_element(convert_imgs))
    # 获取HTML内容
    html = result.value
    # 转化HTML为Markdown
    md = markdownify.markdownify(html,heading_style="ATX")
    print(md)
    with open("./docx_to_html.html",'w',encoding='utf-8') as html_file,open("./docx_to_md.md","w",encoding='utf-8') as md_file:
        html_file.write(html)
        md_file.write(md)
    messages = result.messages
