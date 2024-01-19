from robocorp import browser
from robocorp.tasks import task
from RPA.PDF import PDF
import os
from fpdf import FPDF
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.FileSystem import FileSystem
import time
import zipfile

@task
def order_robots_from_RobotSpareBin():
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

    url = "https://robotsparebinindustries.com/orders.csv"
    file_path = "output/orders.csv"
    download_orders_file(url, file_path)

    orders = read_orders_file(file_path)

    for order in orders:
        close_annoying_modal()
        fill_the_form(order)

        # Pause for elements to appear
        

        # Print receipt
        receipt_pdf_path = store_receipt_as_pdf(order['Order number'])
        robot_screenshot_path = screenshot_robot(order['Order number'])
        if robot_screenshot_path:
            embed_screenshot_to_receipt(robot_screenshot_path, receipt_pdf_path)

        # Force reload the page to reset the state for the next order
        page = browser.page()
        page.reload()

        # If the navigation to the order form is necessary, uncomment the next line
        # browser.goto("https://robotsparebinindustries.com/#/robot-order")

    

# ... rest of your functions ...



def download_orders_file(url, file_path):
    http = HTTP()
    http.download(url, file_path, overwrite=True)

def read_orders_file(file_path):
    tables = Tables()
    return tables.read_table_from_csv(file_path, header=True)

def close_annoying_modal():
    page = browser.page()

    # Remove modal
    if page.is_visible("css=.btn.btn-dark", timeout=5000):  # Timeout can be adjusted
        page.click("css=.btn.btn-dark")
    else:
        print("Modal not found or already closed.")


def fill_the_form(order):
    page = browser.page()
    page.select_option("#head", order["Head"])  
    body_option_id = f"id-body-{order['Body']}"
    page.click(f"#{body_option_id}")
    page.fill("css=.form-control", order["Legs"])  
    page.fill("#address", order["Address"]) 
    page.click("#preview")  

    # Retry mechanism for order submission
    max_attempts = 10
    for attempt in range(max_attempts):
        page.click("#order")  
        # Wait for a short duration for the error message to potentially appear
        time.sleep(2)

        # Check if an error message is displayed on the screen
        if page.is_visible("#alert alert-danger"):  
            print(f"Error in submission, retrying... (Attempt {attempt + 1} of {max_attempts})")
            continue
        else:
           
            break


def store_receipt_as_pdf(order_number):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt=f"Order Receipt: {order_number}", ln=True, align='C')

    

    # Save the PDF
    pdf_output_path = f"output/receipts/receipt_{order_number}.pdf"
    pdf.output(pdf_output_path)
    return pdf_output_path

def screenshot_robot(order_number):
    page = browser.page()

  
    robot_preview_selector = "#robot-preview-image"  
    if page.wait_for_selector(robot_preview_selector, timeout=30000):
        screenshot_path = f"output/screenshots/robot_{order_number}.png"
        page.screenshot(path=screenshot_path, full_page=True)
        return screenshot_path
    else:
        print(f"Robot preview not found for order {order_number}, screenshot skipped.")
        return None



def embed_screenshot_to_receipt(screenshot_path, pdf_file_path):
    pdf = FPDF()
    pdf.add_page()
    pdf.image(screenshot_path, x=10, y=8, w=100)
    pdf.output(pdf_file_path)

   

def archive_receipts():
    output_dir = 'output/receipts'
    zip_filename = 'output/receipts_archive.zip'

    try:
        with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(output_dir):
                for file in files:
                    if file.endswith('.pdf'):
                        # Create a complete file path
                        pdf_path = os.path.join(root, file)
                        # Write the file to the zip archive
                        zipf.write(pdf_path, arcname=file)

        print(f"Archive created at {zip_filename}")
    except Exception as e:
        print(f"An error occurred: {e}")
