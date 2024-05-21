import os

from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
# from RPA.Assistant import Assistant

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(
        slowmo=100
    )
    open_robot_order_website()
    # user_input_task()
    close_annoying_modal()
    orders = get_orders()
    fill_the_form(orders)
    archive_receipts()

def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

# """Robot Assistant task to get user input"""
# def open_robot_order_website(url):
#     """Navigates to the given URL"""
#     browser.goto(url)

#     """Robot Assistant task to get user input"""
# def user_input_task():
#     assistant = Assistant()
#     assistant.add_heading("Input from user")
#     assistant.add_text_input("text_input", placeholder="Please enter URL")
#     assistant.add_submit_buttons("Submit", default="Submit")
#     result = assistant.run_dialog()
#     url = result.text_input
#     open_robot_order_website(url)

def close_annoying_modal():
    page = browser.page()
    page.click("button:text('OK')")

def get_orders():
    """Gets the orders file, reads it as a table, and returns the data"""

    """Downloads csv file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

    """Read data from csv as a table"""
    library = Tables()
    worksheet = library.read_table_from_csv("orders.csv", columns=["Order number", "Head", "Body", "Legs", "Address"])

    """"Return the data"""
    rows = []
    for row in worksheet:
        rows.append(row)
    return rows

def fill_the_form(orders):
    """Fills in the order form"""
    page = browser.page()

    for order in orders:
        if orders.index(order) != 0:
            close_annoying_modal()
        order_number = order["Order number"]
        fill_order(page, order)
        store_receipt_as_pdf(order_number)
        screenshot_robot(order_number)
        embed_screenshot_to_receipt(f"output/receipts/{order_number}.png", f"output/receipts/{order_number}.pdf")
        page.click("text=Order another robot")

def fill_order(page, order):
    """
    Fills the order form, retrying indefinitely until successful or max_time is reached.
    """
    xpathHead = '//select[@id="head"]'
    xpathLegs = '//input[@placeholder="Enter the part number for the legs"]'
    xpathAddress = '//input[@id="address"]'
    xpathButtonPreview = '//button[@id="preview"]'

    page.locator(xpathHead).select_option(order["Head"])
    page.locator(f'//input[@id="id-body-{str(order["Body"])}"]').click()
    page.locator(xpathLegs).fill(order["Legs"])
    page.locator(xpathAddress).fill(order["Address"])
    page.locator(xpathButtonPreview).click()
    while True:
        try:
            xpathButtonOrder = '//button[@id="order"]'
            page.locator(xpathButtonOrder).click()
            # Check for the presence of an error message
            if not page.is_visible("div.alert.alert-danger"):
                print("Order placed successfully.")
                return True
            else:
                print("Error encountered. Retrying...")
        except Exception as e:
            print(f"Exception encountered: {e}. Retrying...")
      
def store_receipt_as_pdf(order_number):
    page = browser.page()
    xpathReceipt = '//div[@id="receipt"]'
    order_receipt_html = page.locator(xpathReceipt).inner_html()
    pdf = PDF()
    pdf.html_to_pdf(order_receipt_html, f"output/receipts/{order_number}.pdf")

def screenshot_robot(order_number):
    page = browser.page()
    xpathRobotImage = '//div[@id="robot-preview-image"]'
    page.locator(xpathRobotImage).screenshot(path=f"output/receipts/{order_number}.png")

def embed_screenshot_to_receipt(screenshot_file_location, pdf_file_location):
    pdf = PDF()
    list_of_files = [screenshot_file_location]
    pdf.add_files_to_pdf(
        files=list_of_files,
        target_document=pdf_file_location,
        append=True
    )
    delete_file(screenshot_file_location)

def delete_file(file_path):
    try:
        os.remove(file_path)
        print(f"File {file_path} has been deleted successfully.")
    except FileNotFoundError:
        print(f"File {file_path} not found.")
    except PermissionError:
        print(f"Permission denied to delete {file_path}.")
    except Exception as e:
        print(f"Error occurred while deleting file {file_path}: {e}")

def archive_receipts():
    lib = Archive()
    lib.archive_folder_with_zip('./output/receipts', 'receipts.zip', recursive=True)