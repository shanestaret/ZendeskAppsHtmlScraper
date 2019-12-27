from selenium import webdriver
from bs4 import BeautifulSoup
import requests
import xlsxwriter
import numpy as np
from urllib.request import Request, urlopen

# Method that will create the excel sheet
def create_excel_spreadsheet():
    # path to the Excel file that will be written
    path = 'C:\\Users\\12158\\Documents\\Shane\\ZendeskAppsNew.xlsx'

    # list that will hold the names of the headers of the Excel file
    headers_list = ['App Name', 'App Price', 'Author', 'Website', 'Email']

    # variable that holds the actual Excel workbook
    workbook = xlsxwriter.Workbook(path)

    # variable that holds the actual Excel worksheet
    worksheet = workbook.add_worksheet()

    # formatting text for header cells
    header_cell_format = workbook.add_format()

    # setting header text to bold
    header_cell_format.set_bold()

    # writing the headers of each column into the worksheet for each header there is
    for i in range(len(headers_list)):
        # writing the header into the worksheet (row, col, text, format)
        worksheet.write_string(0, i, headers_list[i], header_cell_format)

    # Closes the Excel workbook
    workbook.close()

# Method to get a random user agent to use to access the website
def get_random_ua():

    # The random user agent that will be returned
    random_ua = ''

    # The txt file that contains all possible user agents
    ua_file = 'ua_file.txt'
    try:
        # Open the text file and read each line
        with open(ua_file) as f:
            lines = f.readlines()

        # If there are any lines in the text file, choose a random one
        if len(lines) > 0:

            # Get random number
            prng = np.random.RandomState()
            index = prng.permutation(len(lines) - 1)
            idx = np.asarray(index, dtype=np.integer)[0]

            # Assign the randomly picked user agent to the String that will be returned
            random_ua = lines[int(idx)][:-1]

    # If there is an exception, print this
    except Exception as ex:
        print('Exception in random_ua')
        print(str(ex))

    # Always return the String with the user agent
    finally:
        return random_ua

def get_zendesk_apps_info():

    # URL
    zendesk_website = 'https://www.zendesk.com/apps/directory/'

    # Opens the file containing the needed HTML and puts the HTML into a String
    with open('html_file.txt', 'r') as html_file:
        # The String that holds the HTML to parse through
        html_string = html_file.read()

    # sets up Firefox browser to be opened; need to include executable_path
    # driver = webdriver.Firefox(executable_path= 'C:\\Users\\12158\\Documents\\Shane\\geckodriver-v0.26.0-win64\\geckodriver.exe')

    # gets the Zendesk page in Firefox
    # driver.get(zendesk_website)

    # connecting to the Zendesk page
    # result = requests.get(zendesk_website, headers=headers)

    # getting the content from the page
    # content = result.content

    # setting up the parser for the page content
    soup = BeautifulSoup(html_string, "html.parser")

    # list that will hold the names of the apps
    names_list = soup.find_all("span", class_="app-title")

    print(names_list[0].string)

    create_excel_spreadsheet()

if __name__ == "__main__":
    get_zendesk_apps_info()