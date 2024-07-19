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
        # Create a CSV reader object
        csv_reader = csv.reader(csvfile)
        next(csv_reader)

        
        for row in csv_reader:
            csv_data.append(row)
    
    return csv_data

def add_png_to_pdf(pdf_path, output_path, *image_paths ):
    pdf_reader = PdfReader(pdf_path)
    pdf_writer = PdfWriter()

    # Iterate through each page of the PDF and add it to the new PDF
    for page_num in range(len(pdf_reader.pages)):
        page = pdf_reader.pages[page_num]
        pdf_writer.add_page(page)

    # Calculate dimensions for each image and add them to the PDF
    for image_path in image_paths:
        img = Image.open(image_path)
        img_width, img_height = img.size

        # Define position and size for the image on the page
        image_width = 200  # Image width
        image_height = img_height * (image_width / img_width)  # Maintain aspect ratio
        image_x = (letter[0] - image_width) / 2  # Centered horizontally
        image_y = 50  # Position at the bottom with some margin

        # Create a new page and draw the image on it
        packet = BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)
        can.drawImage(image_path, image_x, image_y, width=image_width, height=image_height)
        can.save()

        # Move to the beginning of the StringIO buffer
        packet.seek(0)
        new_pdf = PdfReader(packet)

        # Add the new page with the image to the existing PDF
        pdf_writer.add_page(new_pdf.pages[0])

    # Write the combined PDF to a new file
    with open(output_path, 'wb') as output_file:
        pdf_writer.write(output_file)

    print(f"Images added to '{output_path}' successfully.")

def clean_up(folder_path):

    try:
        # Attempt to delete the folder and its contents
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
        
        # pdf_reader = PdfReader(order_page_path)
        # pdf_writer = PdfWriter()

        # # Add existing pages to new PDF
        # for page_num in range(len(pdf_reader.pages)):
        #     page = pdf_reader.pages[page_num]
        #     pdf_writer.add_page(page)

        # # Create a new page with the image
        # packet = BytesIO()
        # can = canvas.Canvas(packet, pagesize=letter)
        
        # # Load image and get dimensions
        # img = Image.open(robot_page_path)
        # img_width, img_height = img.size
        
        # # Define position and size for the image on the page
        # image_x = 100  # X coordinate (from left)
        # image_y = 100  # Y coordinate (from bottom)
        # image_width = 200  # Image width
        # image_height = img_height * (image_width / img_width)  # Maintain aspect ratio
        
        # # Draw the image on the canvas
        # can.drawImage(head, image_x, image_y, width=image_width, height=image_height)
        # can.save()

        # # Move to the beginning of the StringIO buffer
        # packet.seek(0)
        # new_pdf = PdfReader(packet)

        # # Add the new page with the image to the existing PDF
        # pdf_writer.addPage(new_pdf.getPage(0))

        # # Write the combined PDF to a new file
        # with open(order_page_path, 'wb') as output_file:
        #     pdf_writer.write(order_page_path)

        # print(f"Image added to '{order_page_path}' successfully.")


        # page.screenshot(path=order_page_path)
        page.click("#order-another")
        give_consent()
        print("One Done")


    #----Ignore this this is hard-coded stuff----
    # page = browser.page()
    # page.select_option("#head", "1")
    # body_id= "#id-body-"+"1"
    # page.click("#id-body-1")
    # page.locator(".mb-3 input[placeholder='Enter the part number for the legs']").fill("2")
    # page.locator(".mb-3 input[placeholder='Shipping address']").fill("HYD")
    # while True:
    #     page.click("#order")
    #     try:
    #         page.locator(".alert, .alert-danger").inner_html()
    #     except:
    #         break
    # print("NO")

def zip_up(folder_path, output_zip):
    try:
        # Initialize ZipFile object
        with zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # Walk through all files and folders in the directory
            for root, _, files in os.walk(folder_path):
                for file in files:
                    # Create complete file path
                    file_path = os.path.join(root, file)
                    # Add file to zip
                    zipf.write(file_path, os.path.relpath(file_path, folder_path))

        print(f"Folder '{folder_path}' successfully zipped to '{output_zip}'.")
    except Exception as e:
        print(f"Error zipping folder '{folder_path}': {e}")