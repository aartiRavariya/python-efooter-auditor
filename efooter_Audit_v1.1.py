from tkinter import *
import os
import xlsxwriter
import cv2
from bs4 import BeautifulSoup


def check_values(result):
    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook('Audit_file.xlsx')
    worksheet = workbook.add_worksheet()

    header_format = workbook.add_format({'bold': 1})
    headers = ("Languages", "URLs", "Is Img_Width matched ", "Is Img_Height matched", "Title", "Is Img_src matched")

    # Write some data headers.
    worksheet.write_row(0, 0, headers, header_format)

    # Start from the first cell below the headers.
    row_no = 1
    cell_format = workbook.add_format({'font_color': 'red'})
    cell_format1 = workbook.add_format()
    cell_format1.set_bg_color('red')
    for row in result:
        lang, url, html, width, height, title = row[0],  row[1],  row[2],  row[3], row[4], row[5]
        str_width = 'width="' + str(width) + '"'
        str_height = 'height="' + str(height) + '"'
        str_name = 'src="' + title + '.jpg"'

        checked_width = 'No' if html.find(str_width) < 0 else 'Yes'
        checked_height = 'No' if html.find(str_height) < 0 else 'Yes'
        checked_name = 'No' if html.find(str_name) < 0 else 'Yes'

        worksheet.write(row_no, 0, row[0])

        if url == "":
            worksheet.write(row_no, 1, row[1], cell_format1)
        else:
            worksheet.write(row_no, 1, row[1])

        if checked_width == "Yes":
            worksheet.write(row_no, 2, checked_width)
        else:
            worksheet.write(row_no, 2, checked_width, cell_format)

        if checked_height == "Yes":
            worksheet.write(row_no, 3, checked_height)
        else:
            worksheet.write(row_no, 3, checked_height, cell_format)

        if title == "":
            worksheet.write(row_no, 4, row[5], cell_format1)
        else:
            worksheet.write(row_no, 4, row[5])

        if checked_name == "Yes":
            worksheet.write(row_no, 5, checked_name)
        else:
            worksheet.write(row_no, 5, checked_name, cell_format)
        row_no += 1

    workbook.close()


def submit():
    src = src_dir.get()
    src_dir.set("")

    result = []
    lang = {
        'ar': 'AR', 'de': 'DE', 'en': 'EN', 'es': 'ES', 'fr': 'FR', 'it': 'IT', 'nl': 'NL', 'no': 'NO',
        'pl': 'PL', 'pt': 'PT', 'ru': 'RU', 'se': 'SE', 'cn': 'CN'
    }

    for subdir, dirs, files in os.walk(src):
        for sub_dir in dirs:
            if sub_dir in lang:
                full_path = os.path.join(subdir, sub_dir)
                for item in os.listdir(full_path):
                    if item.endswith('.html'):
                        with open(full_path + '\\' + item, 'r') as f:
                            html = f.read()
                            soup = BeautifulSoup(html, 'lxml')
                            link = soup.find('a')['href']
                            title = soup.find('title').text
                    if item.endswith('.jpg'):
                        src = os.path.join(full_path, item)
                        img = cv2.imread(src)
                        h, w, c = img.shape
                        result.append([lang[sub_dir], link, html, w, h, title])
    check_values(result)


if __name__ == '__main__':

    top = Tk()
    top.geometry("450x300")
    top.title("efooter Auditor")
    top.configure(background="Dark gray")

    input_path = Label(top, text="Enter Input Path:", bg="Dark gray").place(x=40, y=60)

    src_dir = StringVar()
    input_path_entry_area = Entry(top, textvariable=src_dir, width=30).place(x=150, y=60)

    submit_button = Button(top, text="Submit", command=submit).place(x=200, y=150)
    top.mainloop()
