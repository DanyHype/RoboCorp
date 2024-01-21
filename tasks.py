from robocorp import browser
from robocorp.tasks import task
from RPA.PDF import PDF
from robot.api import logger
import os
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.Excel.Files import Files
from RPA.FileSystem import FileSystem
from RPA.Archive import Archive
import csv



@task
def order_robots_from_RobotSpareBin():

    start()

    open_robot_order_website()

    download_order_data()

    complete_robot_form()

    save_receipts()

   


def open_robot_order_website():
    # TODO: Implement your function here
    page = browser.page()
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    close_annoying_modal()


def complete_robot_form():
    """complete parts form"""
    page = browser.page()
    worksheet = get_order()

    for row in worksheet:
         fill_and_submit(row)

    save_receipts()
    

def fill_and_submit(row):
    page = browser.page()

    page.select_option("#head", str(row["Head"]))
    rad = "#id-body-"+str(row["Body"])
    page.set_checked(rad, True)
    page.get_attribute
    page.keyboard.press('Tab')
    page.keyboard.press(str(row["Legs"]))

    page.fill("#address", row["Address"])    

    page.click("#preview")
   

    error = True

    while error:
        page.click("#order")
        error = page.locator("//div[@class='alert alert-danger']").is_visible()

        if not error:
            break

       
    store_receipt_as_pdf(row["Order number"])
    
    page.click("#order-another")
    close_annoying_modal()
         

def get_order():
     library = Tables()
     data =  library.read_table_from_csv("orders.csv")
     return data





def close_annoying_modal():
    page = browser.page()
    page.click("button:text('Yep')")


def download_order_data():
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)


def store_receipt_as_pdf(order_number):

    page = browser.page()
    pdf = PDF()
    pdf_path = "output/receipts/order_"+order_number+".pdf"
   
    order_receipt_html = page.locator("#receipt").inner_html()
    pdf.html_to_pdf(order_receipt_html,  pdf_path)

    #take screenshot
    screenshot_robot(order_number)
   
    #embed screenshot in pdf
    embed_screenshot_to_receipt("output/receipts/order_"+order_number+".png", pdf_path)
    

def screenshot_robot(order_number):
    page = browser.page()
    page.locator(selector="#robot-preview-image").screenshot(path="output/receipts/order_"+order_number+".png")


def embed_screenshot_to_receipt(screenshot, pdf_file):
    pdf = PDF()
    
    pdf.add_files_to_pdf(
        files=[screenshot],
        target_document=pdf_file,
        append=True
    )


def save_receipts():
    """Download files"""
    lib = Archive()
    lib.archive_folder_with_zip("output/receipts",'receipts.zip',recursive=True)


def start():
    """Initiate robot"""
    output_folder()

def output_folder():
    """Set save location"""
    os.makedirs("output/receipts", exist_ok=True)
