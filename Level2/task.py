"""Orders robot from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images."""

#importing libraries
from RPA.Browser.Selenium import Selenium
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
from RPA.Dialogs import Dialogs
from RPA.Robocorp.Vault import Vault

#create objects from the imported libraries
browser = Selenium()
http = HTTP()
tab = Tables()
pdf = PDF()
archive = Archive()
dialogs = Dialogs()
valut = Vault()

#path of the zip file with all receipts
receipts_zip_path = "output/receipts.zip"

#ask to user the path of the .csv file
#"https://robotsparebinindustries.com/orders.csv"
def ask_for_csv():
    dialogs.add_text_input("url", label="Insert the csv address")
    response = dialogs.run_dialog()
    return response.url

#open the browser with the given url
def open_browser(url):
    browser.open_available_browser(url)

#download the resource orders.csv from the site
def download_orders_csv(url):
    http.download(url, overwrite=True)
    return tab.read_table_from_csv("orders.csv")

#fill the orders with the fields of a given row and check if receipt is created
def fill_orders(row):
    browser.select_from_list_by_index("id:head", row["Head"])
    browser.click_element("id:id-body-" + str(row["Body"]))
    browser.input_text("class:form-control", str(row["Legs"]))
    browser.input_text("id:address", row["Address"])
    browser.click_button("id:preview")
    browser.click_button("id:order")
    while not browser.is_element_visible("id:receipt"):
        browser.click_button("id:order")

#save the screenshot of the robot preview
def save_preview_screenshot(order_nr):
    browser.screenshot("id:robot-preview-image", "output/robot" + order_nr + ".png")

#save the receipt from html to pdf and add the preview image of the robot to pdf file
def save_receipt(order_nr):
    receipt = browser.get_element_attribute("id:receipt", "outerHTML")
    pdf.html_to_pdf(receipt, "output_receipts/receipt" + order_nr + ".pdf")
    #pdf.open_pdf("output_receipts/receipt" + order_nr + ".pdf")
    list_to_add = ["output/robot" + order_nr + ".png"]
    pdf.add_files_to_pdf(files=list_to_add, target_document="output_receipts/receipt" + order_nr + ".pdf", append=True)
    #pdf.close_pdf("output/receipt" + order_nr + ".pdf")

#create ZIP archive with the receipts
def archive_to_zip():
    archive.archive_folder_with_zip("output_receipts", receipts_zip_path)

def task():
    #some problems with the versions here - many times library returns error
    secret_url = valut.get_secret("website_url")
    #secret_url = {"url": "https://robotsparebinindustries.com/#/robot-order"}
    open_browser(secret_url["url"])
    csv_url = ask_for_csv()
    orders_table = download_orders_csv(csv_url)
    for row in orders_table:
        if(browser.is_element_visible("class:btn.btn-warning")):
            browser.click_button("class:btn.btn-warning")
        browser.wait_until_element_is_visible("id:head")
        fill_orders(row)
        browser.wait_until_element_is_visible("id:receipt")
        save_preview_screenshot(row["Order number"])
        save_receipt(row["Order number"])
        browser.click_button("id:order-another")
    archive_to_zip()
    browser.close_browser()
    print("Done.")

if __name__ == "__main__":
    task()


