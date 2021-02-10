#!/usr/bin/env python3.7
from pdfminer.layout import LAParams, LTTextBox
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from docx import Document
from docx.enum.text import WD_BREAK
from docx.shared import Pt
import tkinter as tk

sku = []
fnsku = []
pieces = []
units = []
cases = []
total = []

def read_pdf(file):
    print(file)
    fp = open(file, 'rb')
    rsrcmgr = PDFResourceManager()
    laparams = LAParams()
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    pages = PDFPage.get_pages(fp)

    for page in pages:
        print('Processing next page...')
        interpreter.process_page(page)
        layout = device.get_result()
        for lobj in layout:
            if isinstance(lobj, LTTextBox):
                x, y, text = lobj.bbox[0], lobj.bbox[3], lobj.get_text()
                # print('At %r is text: %s' % ((x, y), text))

                # sku
                if x == 42.50000165316652:
                    if len(text) >= 6:
                        sku.append(text)
                        # print('At %r is text: %s' % ((x, y), text))

                # fnsku
                if x == 113.75000324441633:
                    if text[-11:][0:2] == 'X0' or text[-11:][0:2] == 'B0':
                        fnsku.append(text[-11:])
                        # print('At %r is text: %s' % ((x, y), text))

                # pieces
                import re
                if x == 42.50000165316652:
                    if len(text) >= 6:
                        seg = text.split('-')
                        pieces.append(re.sub("[^0-9]", "", seg[1]))

                # If SKU does not contain pieces, use below
                # if 'pack' in text:
                #     pos = text.index('pack')
                #     pieces.append(text[pos-3:pos])
                # elif 'pcs' in text:
                #     pos = text.index('pcs')
                #     pieces.append(text[pos-3:pos])
                # elif 'pack of' in text:
                #     pos = text.index('pack of')
                #     pieces.append(text[pos + 8:pos + 9])
                # else:
                #     pieces.append(1)

                # units
                if x == 435.2422104524071 or x == 440.24611681416104:
                    units.append(text)
                    # print('At %r is text: %s' % ((x, y), text))

                # cases
                if x == 492.74611462666104:
                    cases.append(text)
                    # print('At %r is text: %s' % ((x, y), text))

                # total
                if x == 518.4921997836575 or x == 523.4961098954115:
                    total.append(text)
                    # print('At %r is text: %s' % ((x, y), text))


def make_doc(order):
    document = Document()

    for i in range(len(sku)):
        run = document.add_paragraph().add_run(order)
        run.font.size = Pt(24)
        temp_sku = sku[i].replace('\n', '') + '\n'
        temp_fnsku = fnsku[i].replace('\n', '') + '\n'
        temp_units = pieces[i].replace('\n', '') + ' PCS/UNIT\n'
        temp_cases = str(int(int(units[i].replace('\n', '')) / int(cases[i].replace('\n', '')))) + ' UNITS/CASE\n'
        temp_total = 'TOTAL: ' + cases[i].replace('\n', '') + ' CASES, ' + units[i].replace('\n', '') + ' UNITS'
        text = temp_sku + temp_fnsku + temp_units + temp_cases + temp_total
        run = document.add_paragraph().add_run(text)
        run.font.size = Pt(24)
        run.add_break(WD_BREAK.PAGE)

    document.save('demo.docx')


def main():
    root = tk.Tk()
    frame = tk.Frame(master=root, width=300, height=300)
    frame.grid(row=3, column=2)
    frame.pack()

    v1 = tk.StringVar()
    label1 = tk.Label(master=frame, text='Shipment Number: ')
    entry1 = tk.Entry(master=frame, width=50, textvariable=v1)
    label1.grid(row=1, column=1)
    entry1.grid(row=1, column=2)

    v2 = tk.StringVar()
    label2 = tk.Label(master=frame, text='Shipment PDF with extension: ')
    entry2 = tk.Entry(master=frame, width=50, textvariable=v2)
    label2.grid(row=2, column=1)
    entry2.grid(row=2, column=2)

    def update():
        order = v1.get()
        file = v2.get()
        read_pdf(file)
        make_doc(order)

    button = tk.Button(master=frame, text='Make Doc', command=update)
    button.grid(row=3, column=2)

    root.mainloop()

main()

