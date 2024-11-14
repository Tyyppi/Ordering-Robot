import shutil
from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(slowmo=100)
    open_the_order_website()
    download_orders_list()
    #close_annoying_modal()
    read_orders_from_files()
    archive_receipts()
    remove_temp_files()
    

def open_the_order_website():
    """Opens the website for placing orders"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def download_orders_list():
    """Downloads the orders listing from the website"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def read_orders_from_files():
    """Reads the orders csv file and loop through them"""
    csv = Tables()
    orders = csv.read_table_from_csv(
        "orders.csv", columns=["Order number", "Head", "Body", "Legs", "Address"], header=True
    )

    for order in orders:
        print(order)
        fill_the_order_form(order)

def close_annoying_modal():
    """Close the pop-up window asking about rights"""
    page = browser.page()
    page.get_by_text('OK').click()
    #page.click("button:text('OK')")

def fill_the_order_form(order):
    """Fills in the order information to the website"""
    close_annoying_modal()
    page = browser.page()
    #page.select_option("#head", str(order['Head']))
    page.locator("#head").select_option(order['Head'])
    #page.click("#id-body-"+str(order['Body']))
    page.locator("#id-body-"+str(order['Body'])).click()

    page.get_by_placeholder("Enter the part number for the legs").fill(order['Legs'])
    #page.fill("#address", order['Address'])
    page.locator("#address").fill(order['Address'])
    screenshot = screenshot_robot(order['Order number'])
    
    alert = True
    while alert:
        page.click("button:text('order')")
        alert = page.locator("[class='alert alert-danger']").is_visible()
        if not alert:
            break

    order_pdf = store_order_receipt_as_pdf(order['Order number'])
    embed_screenshot_to_receipt(screenshot, order_pdf)
    page.click("#order-another")

def store_order_receipt_as_pdf(order_number):
    """Saves the order receipt from the website as PDF"""
    page = browser.page()
    order_receipt = page.locator("#receipt").inner_html()

    pdf = PDF()
    filename = "output/orders/receipt_order_"+order_number+".pdf"
    pdf.html_to_pdf(order_receipt, filename)

    return filename

def screenshot_robot(order_number):
    """Takes a screenshot of the ordered robot"""
    page = browser.page()
    page.click("#preview")
    filename = "output/screenshots/robot_order_"+order_number+".png"
    page.locator("#robot-preview-image").screenshot(path=filename)

    return filename

def embed_screenshot_to_receipt(robot_screenshot, order_pdf):
    """Add the screenshot of the robot to its order receipt"""
    pdf = PDF()
    files = [order_pdf, robot_screenshot]
    pdf.add_files_to_pdf(files=files, target_document=order_pdf, append=False)

def archive_receipts():
    """Add order pdfs to a zip archive"""
    lib = Archive()
    lib.archive_folder_with_zip("output/orders", "output/receipts.zip", recursive=True)

def remove_temp_files():
    """Cleans up unneeded pdf and png files"""
    #os.remove("output/screenshots/*")
    #os.rmdir("output/screenshots")
    shutil.rmtree("output/screenshots")
    #os.remove("output/orders/*")
    #os.rmdir("output/orders")
    shutil.rmtree("output/orders")
