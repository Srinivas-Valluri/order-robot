from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
import csv
from RPA.PDF import PDF
import PyPDF2
from PyPDF2 import PdfReader, PdfWriter
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from io import BytesIO
from PIL import Image
import shutil
import zipfile
import os

pdf = PDF()


@task
def robot_spare_bin():
    browser.configure(
        slowmo=100,
    )
    open_the_website()
    give_consent()
    download_excel_file()
    csv_data = read_data_from_csv()
    fill_form_using_data(csv_data)
    zip_up('./output/orders', './output/orders.zip')
    clean_up('./output/orders')


def open_the_website():
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def give_consent():
    page = browser.page()
    page.click("button:text('OK')")
    
def download_excel_file():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def read_data_from_csv():
    csv_file_path = './orders.csv'
    csv_data = []
    with open(csv_file_path, mode='r', newline='') as csvfile:
        csv_reader = csv.reader(csvfile)
        next(csv_reader)

        
        for row in csv_reader:
            csv_data.append(row)
    
    return csv_data

def add_png_to_pdf(pdf_path, output_path, *image_paths ):
    pdf_reader = PdfReader(pdf_path)
    pdf_writer = PdfWriter()

    for page_num in range(len(pdf_reader.pages)):
        page = pdf_reader.pages[page_num]
        pdf_writer.add_page(page)

    for image_path in image_paths:
        img = Image.open(image_path)
        img_width, img_height = img.size

        image_width = 200  # Image width
        image_height = img_height * (image_width / img_width) 
        image_x = (letter[0] - image_width) / 2  
        image_y = 50  

        packet = BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)
        can.drawImage(image_path, image_x, image_y, width=image_width, height=image_height)
        can.save()

        packet.seek(0)
        new_pdf = PdfReader(packet)

        pdf_writer.add_page(new_pdf.pages[0])

    with open(output_path, 'wb') as output_file:
        pdf_writer.write(output_file)

    print(f"Images added to '{output_path}' successfully.")

def clean_up(folder_path):

    try:
        shutil.rmtree(folder_path)
        print(f"Deleted folder '{folder_path}' successfully.")
    except Exception as e:
        print(f"Error deleting folder '{folder_path}': {e}")

def fill_form_using_data(csv_data):
    page = browser.page()
    for row in csv_data:
        print(row)
        page.select_option("#head", row[1])
        body_id= "#id-body-"+row[2]
        page.click(body_id)
        page.locator(".mb-3 input[placeholder='Enter the part number for the legs']").fill(row[3])
        page.locator(".mb-3 input[placeholder='Shipping address']").fill(row[4])
        while True:
            page.click("#order")
            try:
                page.locator(".alert, .alert-danger").inner_html()
            except:
                break
        order_page_path = f"output/orders/sales_summary_{row[0]}.pdf"
        robot_head_page_path = f"output/parts/sales_summary_robot_head_{row[0]}.png"
        robot_body_page_path = f"output/parts/sales_summary_robot_body_{row[0]}.png"
        robot_legs_page_path = f"output/parts/sales_summary_robot_legs_{row[0]}.png"
        recept = str(page.locator("#receipt").inner_html())
        head = str(page.locator("#robot-preview-image img[alt='Head']").screenshot(path=robot_head_page_path))
        body = str(page.locator("#robot-preview-image img[alt='Body']").screenshot(path=robot_body_page_path))
        legs = str(page.locator("#robot-preview-image img[alt='Legs']").screenshot(path=robot_legs_page_path))
        pdf.html_to_pdf(recept, order_page_path)

        add_png_to_pdf(order_page_path, order_page_path, robot_head_page_path, robot_body_page_path, robot_legs_page_path)
        clean_up('./output/parts')
        
        
        page.click("#order-another")
        give_consent()
        print("One Done")

def zip_up(folder_path, output_zip):
    try:
        with zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(folder_path):
                for file in files:
                    file_path = os.path.join(root, file)
                    zipf.write(file_path, os.path.relpath(file_path, folder_path))

        print(f"Folder '{folder_path}' successfully zipped to '{output_zip}'.")
    except Exception as e:
        print(f"Error zipping folder '{folder_path}': {e}")