import fitz
import pandas as pd

Filename = pd.read_excel('Files.xlsx')
search = pd.read_excel('Search.xlsx')
l=[]
l=Filename['File_Name'].tolist()
chg=search['Changes'].tolist()
for i in l:
    pdfIn = fitz.open(i)
    
    for page in pdfIn:
        print(page)
        texts = search['Text_to_Redline'].tolist()
        text_instances = [page.search_for(text) for text in texts]    
        for inst,j in zip(text_instances,chg):
               
            annot = page.add_highlight_annot(inst)
            annot = page.add_strikeout_annot(inst)
            #page.add_freetext_annot(rect=(100,100,500,500),text="Rashid")
            #annot = page.add_rect_annot(inst)
               
            info = annot.info
            info["title"] = "Replace With"
            info["content"] = j
            annot.set_info(info)
            annot.update()
            #highlight = page.add_highlight_annot(inst)
            #highlight.update()
    pdfIn.save("Redline_"+i)