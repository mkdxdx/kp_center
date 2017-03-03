from xlwt import Workbook, XFStyle, Borders, easyxf
from grab import Grab
import re
import tkinter as tk
import os

list_url = [
    'http://center-bespeki.com/catalog/products/analogovoe_video_nablyudenie/ahvr0804f/,2',
    'http://center-bespeki.com/catalog/products/hd_cvi_videonablyudenie/blc_s2mp30ir/,2',
    'http://center-bespeki.com/catalog/products/analogovoe_video_nablyudenie/ac_a251ir1/,10',
    'http://center-bespeki.com/catalog/products/analogovoe_video_nablyudenie/ac_a254ir5/',
    'Установка внешней камеры,10,200',
    'Выезд,1,300']

book_filename = 'kp.xls'

# gui declaration section
mainw = tk.Tk()
mainw.geometry("800x600")
mainw.title("Коммерческое предложение по ЦБ")
mainw.columnconfigure(0,weight=1)
mainw.rowconfigure(0,weight=1)
put_url = tk.BooleanVar()
put_url.set("0")

# parse engine section

column_offset = 1
row_offset = 2


def strip_currency(str):
    return re.sub(r'[А-Яа-я\s]*', '', str)


def swap_delimiter(str):
    return re.sub(r'\.', ',', str)


def strip_dollarfor(str):
    return re.sub(r' за[\w\W]*$', '', str)


def walk_url(url):
    gr = Grab()
    gr.go(url)
    cont = gr.doc.select(
        '//div[contains(@class,"container tb10")]//div[contains(@class,"wrapper")]//div[contains(@class,"container")]')
    hw_name = strip_dollarfor(cont.select('.//div[contains(@class,"span_5_of_5")]//h1').text())
    hw_price = strip_currency(cont.select('.//div[contains(@class,"emarket-current-price emarket-format-price")]').text())
    try:
        hw_desc_caption = cont.select('.//div[contains(@class,"emarket-detail-area-container")]//h2').text()
    except IndexError:
        hw_desc_caption = hw_name
    print(hw_name + '/' + hw_price)
    result = [hw_desc_caption, hw_price, url]
    return result


def alignwrite(sh,ix,iy,text,borders):
    if borders == True:
        sh.write(ix, iy, text, style=easyxf('border: top medium, right medium, bottom medium, left medium;'))
    elif borders == False:
        sh.write(ix, iy, text, style=easyxf('border: top no_line, right no_line, bottom no_line, left no_line;'))
    if len(text) * 256 > sh.col(iy).width:
        sh.col(iy).width = len(text) * 256


def write_book(header ,columns, data, grid_limit):
    book = Workbook(encoding="utf-8")
    sh = book.add_sheet('Список оборудования и услуг')
    alignwrite(sh, row_offset, column_offset, header, False)

    ci = 0
    ri = 2
    for cname in columns:
        alignwrite(sh, ri + row_offset, ci + column_offset, cname, False)
        ci = ci + 1

    ri = ri + 1

    for rd in data:
        ci = 0
        for cd in rd:
            if ri<=grid_limit+row_offset and not cd == "":
                alignwrite(sh, ri + row_offset, ci + column_offset, re.sub(r'\'', '', cd),True)
            else:
                alignwrite(sh, ri + row_offset, ci + column_offset, re.sub(r'\'', '', cd),False)
            ci = ci + 1
        ri = ri + 1
    book.save(book_filename)
    os.startfile(book_filename)
    pass



tb_linklist = tk.Text(mainw)
tb_linklist.grid(row=0,column=0,sticky=tk.N+tk.W+tk.E+tk.S,columnspan=3)

en_client = tk.Entry(mainw)
en_client.grid(row=1,column=0,sticky=tk.W+tk.N+tk.E)

def fill_example():
    for url in list_url:
        tb_linklist.insert(tk.INSERT,url+"\n")
    en_client.insert(0,'Клиент, т. 111 222 33 44')


def process_textbox():
    if (len(tb_linklist.get('1.0','end-1c')) == 0):
        fill_example()

    header = "Предложение для: "+en_client.get()
    columns = ['Название', 'Цена', 'Кол-во, ед', 'Стоимость']

    if put_url == True:
        columns.append('Ссылка')
    rows = []
    prepay = 0
    sum = 0
    for entry in tb_linklist.get('1.0','end-1c').splitlines():
        if entry:
            row = ['','0','1','0','']
            ldata = entry.split(',')
            if "://" in ldata[0]:
                # its an url
                page_data = walk_url(ldata[0])
                row[0] = page_data[0]
                row[1] = swap_delimiter(page_data[1])
                price = float(page_data[1])
                try:
                    row[2] = ldata[1]
                    price = float(ldata[1])*float(page_data[1])
                except IndexError:
                    pass

                row[3] = swap_delimiter(str(price))
                prepay = prepay + price
                sum = sum + price
                if put_url.get() == 1:
                    row[4] = page_data[2]
            else:
                # its a plain text
                ind = 0
                row[0] = ldata[0]
                try:
                    row[1] = swap_delimiter(ldata[1])
                    price = float(ldata[1])
                    try:
                        price = float(ldata[1]) * float(ldata[2])
                        row[3] = swap_delimiter(str(price))
                    except IndexError:
                        pass
                    sum = sum + price
                except IndexError:
                    pass
                for coldata in ldata:
                    row[ind] = coldata
                    ind = ind + 1

            rows.append(row)
    styled_rows = len(rows)
    rows.append(['','','Предоплата:', swap_delimiter(str(prepay))])
    rows.append(['','','Всего:', swap_delimiter(str(sum))])
    write_book(header,columns,rows,styled_rows)
    pass

b_process = tk.Button(mainw, text="Сформировать документ", command=process_textbox)
b_process.grid(row=1,column=1,sticky=tk.W+tk.N)

cb_includelink = tk.Checkbutton(mainw,text="Включать ссылки на страницу ЦБ",variable=put_url)
cb_includelink.grid(row=1,column=2,sticky=tk.NW)

mainw.mainloop()