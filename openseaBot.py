from selenium import webdriver
from time import sleep
from selenium.webdriver.chrome.options import Options
from shutil import which
from scrapy import Selector
import xlsxwriter
import datetime


targetURL = "https://opensea.io/activity/theo-nft"


months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sept", "Oct", "Nov", "Dec"]
format = '%m-%d-%Y %I:%M:%S %p'

userEnteredDate = input("Enter Max Limit of Date(e.g 29-02-2021): ")

workbook = xlsxwriter.Workbook(userEnteredDate + '.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write(0, 0, "Item")
worksheet.write(0, 1, "Unit Price")
worksheet.write(0, 2, "Quantity")
worksheet.write(0, 3, "From")
worksheet.write(0, 4, "To")
worksheet.write(0, 5, "Date")
rowN = 1


userEnteredDate = userEnteredDate + " 01:01:01 AM"
userEnteredDate = datetime.datetime.strptime(userEnteredDate, "%d-%m-%Y %I:%M:%S %p")

scrappedURLs = []


def botInitialization():
    # Initialize the Bot
    chromeOptions = Options()
    # chromeOptions.add_argument("--headless")
    chromePath = which("chromedriver")
    driver = webdriver.Chrome(executable_path=chromePath, options=chromeOptions)
    driver.maximize_window()
    driver.execute_script("window.open()")
    return driver


driver = botInitialization()

driver.get(targetURL)
sleep(5)

while True:
    print("Scrolling & getting Results... ")
    rowLastPrev = ""
    response = Selector(text=driver.page_source)
    oldRows = response.css("div.EventHistory--row")
    for i in range(5):
        rowLast = driver.find_elements_by_css_selector("div.EventHistory--row")[-1]
        driver.execute_script("return arguments[0].scrollIntoView();", rowLast)
        sleep(3)

    response = Selector(text=driver.page_source)
    rows = response.css("div.EventHistory--row")
    if len(rows) == len(oldRows):
        workbook.close()
        driver.quit()
        exit(0)

    driver.switch_to.window(driver.window_handles[1])
    for row in rows:
        date = row.css("div[data-testid='EventTimestamp'] > a.ca-DaLy::attr(href)").extract_first()
        if date not in scrappedURLs:
            item = row.css("span.AssetCell--name::text").extract_first()
            price = row.css("div.Price--amount::text").extract_first()
            quantity = row.css("div.EventHistory--quantity-col::text").extract_first()
            accounts = row.css("a.AccountLink--ellipsis-variant-both > span::text").extract()
            From = accounts[0]
            To = accounts[1]

            driver.get(date)
            scrappedURLs.append(date)

            resp = Selector(text=driver.page_source)
            date = resp.css("#ContentPlaceHolder1_divTimeStamp > div > div.col-md-9::text").extract()[1].strip()
            dateOld = date.split("(")[-1].split(")")[0].split("+")[0].strip()

            dateStr = dateOld
            for i, m in enumerate(months):
                if m in dateStr:
                    dateStr = dateStr.replace(m, str(i + 1))

            date = datetime.datetime.strptime(dateStr, format)

            if date < userEnteredDate:
                workbook.close()
                driver.quit()
                exit(0)

            worksheet.write(rowN, 0, item)
            worksheet.write(rowN, 1, price)
            worksheet.write(rowN, 2, quantity)
            worksheet.write(rowN, 3, From)
            worksheet.write(rowN, 4, To)
            worksheet.write(rowN, 5, dateOld)
            rowN += 1
            print([item, price, quantity, From, To, dateOld])

    driver.switch_to.window(driver.window_handles[0])
