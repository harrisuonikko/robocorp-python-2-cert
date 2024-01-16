from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
import os
import time

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
        slowmo=100,
    )
    open_robot_order_website()
    close_popup_if_exists() 
    orders = get_orders()
    for row in orders:
        fill_the_form(row)
        screenshot = store_robot_picture(row["Order number"])
        submit_order()
        receipt = store_receipt_as_pdf(row["Order number"])
        embed_screenshot_to_receipt(screenshot, receipt)
        order_another_robot()
    archive_receipts()
    cleanup()


def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def close_popup_if_exists():
    """Close popup if it exists."""
    page = browser.page()
    try:
        page.eval_on_selector(
            ".modal-content", "el => el.querySelector('button').click()"
        )
    except Exception:
        # Popup did not exist
        pass

def get_orders():
    """Download and store orders"""
    http = HTTP()
    tables = Tables()
    file_name = "orders.csv"
    url = "https://robotsparebinindustries.com/orders.csv"
    http.download(url, overwrite=True)
    orders = tables.read_table_from_csv(file_name)
    return orders

def fill_the_form(row):
    """Fills in the data and click the 'Preview' button"""
    page = browser.page()
    page.select_option("#head", str(row["Head"]))
    page.click(f"#id-body-" + str(row["Body"]))
    page.fill("input[placeholder='Enter the part number for the legs']", str(row["Legs"]))
    page.fill("#address", row["Address"])
    page.click("text=Preview")
    
def submit_order():
    page = browser.page()
    retries = 0
    while retries < 5:
        page.click("#order")  # Click the button by ID
        if page.query_selector("#receipt"):
            return  # Element found, exit the function

        retries += 1  # Increment retry count
        time.sleep(1)
    raise Exception("Maximum retries reached, #receipt not found")

def store_robot_picture(order_number):
    """Take a screenshot of the robot"""
    page = browser.page()
    image = page.locator("#robot-preview-image")
    screenshot_path = f"output/orders/{order_number}.png"
    image = browser.screenshot(image)
    with open(screenshot_path, "wb") as file:
        file.write(image)
    return screenshot_path

def store_receipt_as_pdf(order_number):
    pdf = PDF()
    path = f"output/receipts/{order_number}.pdf"
    page = browser.page()
    receipt = page.inner_html("#receipt")
    pdf.html_to_pdf(receipt, path)
    return path

def order_another_robot():
    page = browser.page()
    page.click("#order-another")
    close_popup_if_exists()

def embed_screenshot_to_receipt(screenshot, pdf_file):
    pdf = PDF()
    pdf.add_files_to_pdf(
        files=[
            pdf_file,
            screenshot + ":align=center",
        ],
        target_document=pdf_file,
    )

def archive_receipts():
    """Archive receipts to zip file."""
    archive = Archive()
    date = time.strftime("%Y-%m-%d-%H-%M-%S")
    archive.archive_folder_with_zip("output/receipts", f"output/receipts-{date}.zip")

def cleanup():
    """Cleanup output folders."""
    for file in os.listdir("output/receipts"):
        os.remove(f"output/receipts/{file}")
    for file in os.listdir("output/orders"):
        os.remove(f"output/orders/{file}")