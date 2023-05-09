import configparser
import pyodbc
import os
import time
import win32api
import win32print
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from datetime import datetime
from barcode import Code128
from barcode.writer import ImageWriter

def Connect_Product():
    config = configparser.ConfigParser()
    config.read("config.txt")

    driver = config["PRODUCT SERVER"]["driver"]
    server_name = config["PRODUCT SERVER"]["server_name"]
    database_name = config["PRODUCT SERVER"]["database_name"]
    username = config["PRODUCT SERVER"]["username"]
    password = config["PRODUCT SERVER"]["password"]

    connectionString = "DRIVER=" + driver + ";SERVER=" + server_name + ";DATABASE=" + database_name + ";UID=" + username + ";PWD=" + password
    return connectionString

class toPDF():
    def pdfFile(self):
        def split_text(text, style):
            lines = text.split('\n')
            return [Paragraph(line, style) for line in lines]

        config = configparser.ConfigParser()
        config.read("config.txt")

        printer_width = float(config["USE PAPER SIZE"]["printer_width"])
        printer_height = float(config["USE PAPER SIZE"]["printer_height"])
        page_width = float(config["USE PAPER SIZE"]["page_width"])
        page_height = float(config["USE PAPER SIZE"]["page_height"])
        table_width = float(config["USE PAPER SIZE"]["table_width"])
        top_margin = float(config["USE PAPER SIZE"]["top_margin"])
        bottom_margin = float(config["USE PAPER SIZE"]["bottom_margin"])
        left_margin = float(config["USE PAPER SIZE"]["left_margin"])
        right_margin = float(config["USE PAPER SIZE"]["right_margin"])
        font_size = float(config["USE PAPER SIZE"]["font_size"])
        font_leading = float(config["USE PAPER SIZE"]["font_leading"])
        printer_name = config["USE PAPER SIZE"]["printer_name"]
        folder_path = config["USE PAPER SIZE"]["folder_path"]

        # Set up the document
        fd = folder_path
        now = datetime.now()
        today = now.strftime('%Y-%m-%d')
        current_time = now.strftime('%H-%M-%S')
        filename = f"{today} {current_time}.pdf"
        file_path = os.path.join(fd, filename)
        doc = SimpleDocTemplate(file_path, pagesize=(
            page_width * inch,
            page_height * inch),
                                leftMargin=left_margin * inch,
                                rightMargin=right_margin * inch,
                                topMargin=top_margin * inch,
                                bottomMargin=bottom_margin * inch)

        pdfmetrics.registerFont(TTFont('Impact', 'impact.ttf'))

        style = getSampleStyleSheet()["Normal"]
        # style.fontName = "Impact"
        style.fontSize = font_size
        style.leading = font_leading  # This is the line spacing
        style.alignment = 1  # This sets the alignment to center

        header_style = getSampleStyleSheet()['Normal']
        header_style.fontName = "Impact"
        # header_style.alignment = 2
        header_style.fontSize = 9
        header_style.leading = 7  # This is the line spacing

        conn = pyodbc.connect(Connect_Product())
        cursor = conn.cursor()
        cursor.execute(f"SELECT * FROM bcodeprinter")
        data = cursor.fetchone()

        number = '4800011121512'
        my_code = Code128(number, writer=ImageWriter())
        my_code.save("barcode")

        cursor.execute(f"select * from items where itemcode = '{number}'")
        item = cursor.fetchone()
        itemname = item[2].strip()
        itemprice = round(item[4], 2)

        def add_footer(canvas, doc):
            canvas.saveState()
            # Define the header text for each column
            header_text = f"<b>PHP {itemprice}</b>"
            # Create a Paragraph object with the respective text and style
            P = Paragraph(header_text, header_style)
            # Define the column positions and offsets
            column_width = doc.width / 3
            column_positions = [(0.5 * inch), (0.69 * inch) + column_width, (0.89 * inch) + 2 * column_width]
            # Loop through each column and draw the header text
            for i, position in enumerate(column_positions):
                # Calculate the width and height of the Paragraph object
                w, h = P.wrap(column_width, doc.bottomMargin)
                # Draw the Paragraph object on the canvas with appropriate positioning
                P.drawOn(canvas, position, doc.height - h - 60)
            canvas.restoreState()

        #Create the table data
        column_text = [f"<b>ZANKPOS</b>"
                       f"\n<i><b>{itemname[:46]}</b></i>"
                       f"\n--\n--\n--\n<img src='barcode.png' width='1.2in' height='0.42in'/>"] * 3
        data = [[split_text(text, style) for text in column_text]]
        # ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        table_style = TableStyle([('ALIGN', (0, 0), (0, -1), 'CENTER'),
                                  ('TOPPADDING', (0, 0), (-1, -1), 0),
                                  ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
                                  ('LEFTPADDING', (0, 0), (-1, -1), 0),
                                  ('RIGHTPADDING', (0, 0), (-1, -1), 0)])
        table_width = table_width * inch

        # Calculate the column widths for the table
        num_cols = len(data[0])
        col_width = table_width / num_cols
        col_widths = [col_width] * num_cols

        # Create the table and add it to the document
        table = Table(data, colWidths=col_widths)
        table.setStyle(table_style)
        doc.build([table], onFirstPage=add_footer)

        path = file_path
        printer_name = printer_name
        printer_handle = win32print.OpenPrinter(printer_name)

        try:
            # Set the printer paper size
            properties = win32print.GetPrinter(printer_handle, 2)
            properties["pDevMode"].PaperSize = 256  # custom paper size
            properties["pDevMode"].PaperWidth = int(printer_width * 254)
            properties["pDevMode"].PaperLength = int(printer_height * 254)

            # Print the PDF file to the specified printer
            win32api.ShellExecute(
                0,
                "printto",
                path,
                f'"{printer_name}"',
                ".",
                0
            )
            print(f"Printing {path} to {printer_name}...")
            time.sleep(5)
        except Exception as e:
            print(f"Error: {e}")
        finally:
            # Close the printer handle
            win32print.ClosePrinter(printer_handle)
            pdf_reader = "Acrobat.exe"
            os.system(f"taskkill /f /im {pdf_reader}")
            print("Adobe Acrobat closed.")

t = toPDF()
t.pdfFile()