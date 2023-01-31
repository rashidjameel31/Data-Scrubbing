from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import openpyxl
from openpyxl import Workbook
import PySimpleGUI as sg
from sys import exit

sg.theme('Topanga')
layout = [  [sg.Text('Enter your Agile user name'), sg.InputText()],
            [sg.Text('Enter you Agile password'), sg.InputText(key='password', password_char = '*')],
            [sg.Text('Select the excel file consisting label name')],
            [sg.Input(), sg.FileBrowse(key="exlo")],
            #[sg.Text('Enter the name of the output data excel'), sg.InputText()],
            #[sg.Checkbox('Check if excel consist header', default=False, key="-IN-")],
            [sg.Button('Ok'), sg.Button('Cancel')] ]
window = sg.Window('Data Scraping from Agile', layout)
event, values = window.read()
password = values['password']
window.close()
if event == "Ok":
    inst = webdriver.Chrome()
    try:
        Filename = pd.read_excel(values["exlo"])
    except:
        layout1 = [  [sg.Text('Unable to open the selected excel. Please check the file')],
                    [sg.Button('Ok')] ]
        window1 = sg.Window('File Error', layout1)
        newevent,newvalues = window1.read()
        window1.close()
        inst.close()
        exit()
    l=Filename['File_Name'].tolist()
    book = Workbook()
    sheet = book.active
    inst.get('https://agileprod.jnj.com/Agile/default/login-cms.jsp')
    window_before = inst.window_handles[0]
    usr=inst.find_element('id','j_username')
    usr.send_keys(values[0])
    pas=inst.find_element('id','j_password')
    pas.send_keys(password)
    print(password)
    inst.find_element('id','login').click()
    inst.implicitly_wait(2)
    wb = openpyxl.Workbook()
    ws = wb.active
    header=[]
    header.append("Number")
    header.append("Document Type")
    header.append("Lifecycle Phase")
    header.append("Description")
    header.append("Class /Document Category")
    header.append("Document Security")
    header.append("Product Line(s)")
    header.append("Rev Incorp Date")
    header.append("Rev Release Date")
    header.append("EFFECTIVE FROM")
    header.append("Product Family")
    header.append("Material Group")
    header.append("Material Type")
    header.append("Responsible Location")
    header.append("Plant")
    header.append("Language")
    header.append("Formula")
    header.append("Transfer to LMS?")
    header.append("Service Library")
    header.append("Point(s) of Use")
    header.append("Notes")
    header.append("Create User")
    header.append("Document Owner")
    ws.append(header)
    wb.save('Output.xlsx')
    for i in l:
        err=[]
        f=[]
        try:
            window_after = inst.window_handles[1]
        except:
            layout1 = [  [sg.Text('Wrong Credentials. Try again')],
                    [sg.Button('Ok')] ]
            window1 = sg.Window('Credential Error', layout1)
            newevent,newvalues = window1.read()
            window1.close()
            inst.close()
            exit()
        inst.switch_to.window(window_after)
        inst.implicitly_wait(2)
        try:
            WebDriverWait(inst, 20).until(EC.element_to_be_clickable((By.ID, 'QUICKSEARCH_STRING')))
        except:
            layout1 = [  [sg.Text('Unable to load website. Please rerun')],
                         [sg.Button('Ok')] ]
            window1 = sg.Window('Automatic Scripts', layout1)
            newevent,newvalues = window1.read()
            window1.close()
            inst.quit()
        inst.find_element('id','QUICKSEARCH_STRING').clear()
        search=inst.find_element('id','QUICKSEARCH_STRING')
        search.send_keys(i)
        inst.find_element(By.ID, "selector_elm").click()
        inst.find_element(By.LINK_TEXT, "Items").click()
        inst.find_element(By.CSS_SELECTOR, ".quick_search").click()
        inst.implicitly_wait(2)
        try:
            inst.find_element(By.ID, "col_1001")
        except:
            try:
                inst.find_element(By.LINK_TEXT,i).click()
            except:
                err.append(i)
                err.append("Label not found")
                ws.append(err)
                wb.save('Output.xlsx')
                pass
                continue
        try:
            f.append(inst.find_element(By.ID, "col_1001").text)
        except :
            pass
        try:
            f.append(inst.find_element(By.ID, "col_1081").text)
        except :
            pass
        try:
            f.append(inst.find_element(By.ID, "col_1084").text)
        except :
            pass
        try:
            f.append(inst.find_element(By.ID, "col_1002").text)
        except:
            pass
        try:
            f.append(inst.find_element(By.ID, "col_1082").text)
        except:
            pass
        try:
            f.append(inst.find_element(By.ID, "col_1068").text)
        except :
            pass
        try:
            f.append(inst.find_element(By.ID, "col_1004").text)
        except :
            pass
        try:
            f.append(inst.find_element(By.ID, "col_1017").text)
        except :
            pass
        try:
            f.append(inst.find_element(By.ID, "col_1016").text)
        except :
            pass
        try:
            f.append(inst.find_element(By.ID, "col_12089").text)
        except :
            pass
        try:
            f.append(inst.find_element(By.ID, "col_2000008074").text)
        except :
            pass
        try:
            f.append(inst.find_element(By.ID, "col_2023").text)
        except :
            pass
        try:
            f.append(inst.find_element(By.ID, "col_2029").text)
        except :
            pass
        try:
            f.append(inst.find_element(By.ID, "col_2092").text)
        except :
            pass
        try:
            f.append(inst.find_element(By.ID, "col_2090").text)
        except :
            pass
        try:
            f.append(inst.find_element(By.ID, "col_2091").text)
        except :
            pass
        try:
            f.append(inst.find_element(By.ID, "col_2007").text)
        except :
            pass
        try:
            f.append(inst.find_element(By.ID, "col_2024").text)
        except :
            pass
        try:
            f.append(inst.find_element(By.ID, "col_2021").text)
        except :
            pass
        try:
            f.append(inst.find_element(By.ID, "col_2000008063").text)
        except :
            pass
        try:
            f.append(inst.find_element(By.ID, "col_1080").text)
        except :
            pass
        try:
            f.append(inst.find_element(By.ID, "col_1420").text)
        except :
            pass
        try:
            f.append(inst.find_element(By.ID, "col_1271").text)
        except :
            pass
        ws.append(f)
        wb.save('Output.xlsx')
    inst.implicitly_wait(5)
    layout1 = [  [sg.Text('Program Complete')],
                [sg.Button('Ok')] ]
    window1 = sg.Window('Automatic Scripts', layout1)
    newevent,newvalues = window1.read()
    window1.close()
    inst.quit()
    exit()
elif event == sg.WIN_CLOSED or event=="Exit":
    exit()

    