
from playwright.sync_api import sync_playwright
import pandas as pd
import sys


def run(playwright):
    browser = playwright.chromium.launch(headless=True)
    context = browser.new_context(accept_downloads=True)
    reliablePath = 'D:\Salman_Api\Grades_Report_Query\Reports'
 
    # Open ruppin homepage refill username and password
    # Goes to grades page
    page = context.new_page()
    page.goto('https://mtrnews.ruppin.ac.il/yedion/fireflyweb.aspx')
    page.fill('input#R1C1', sys.argv[1])
    page.click('input#R1C2')
    page.fill('input#R1C2', sys.argv[2])
    page.click('button[type=submit]')
    page.click('div[class=ButtonMenuLinkCircle]')  
    page.click('a[accesskey=A]')

    with page.expect_download() as download_info:
        page.click('a#Excel')
    download = download_info.value
    path = download.path()
    download.save_as(reliablePath + f'\\{sys.argv[3]}.xlsx')  
    print(path)

    context.close()
    browser.close()

    read_file = pd.read_excel (f'D:\Salman_Api\Grades_Report_Query\Reports\{sys.argv[3]}.xlsx')
    read_file.to_csv (f'D:\Salman_Api\Grades_Report_Query\Reports\{sys.argv[3]}.csv', encoding='ANSI', index = None, header=True)

with sync_playwright() as playwright:
    run(playwright)