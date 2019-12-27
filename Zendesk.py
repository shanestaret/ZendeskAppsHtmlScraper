import time

from selenium import webdriver
from bs4 import BeautifulSoup
import xlsxwriter
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.firefox.options import Options

# Method that will create the Excel sheet
def create_excel_spreadsheet(name_list, author_list, price_list, website_list, email_list):
    # path to the Excel file that will be written
    path = 'C:\\Users\\12158\\Documents\\Shane\\ZendeskAppsNew.xlsx'

    # list that will hold the names of the headers of the Excel file
    header_list = ['App Name', 'App Price', 'Author', 'Website', 'Email']

    # variable that holds the actual Excel workbook
    workbook = xlsxwriter.Workbook(path)

    # variable that holds the actual Excel worksheet
    worksheet = workbook.add_worksheet()

    # formatting text for header cells
    header_cell_format = workbook.add_format()

    # setting header text to bold
    header_cell_format.set_bold()

    # writing the headers of each column into the worksheet for each header there is
    for i in range(len(header_list)):
        # writing the header into the worksheet (row, col, text, format)
        worksheet.write_string(0, i, header_list[i], header_cell_format)

    # writing the app name, author, price, website, and email into each row
    for i in range(len(name_list)):
        worksheet.write_string(i + 1, 0, name_list[i])
        worksheet.write_string(i + 1, 1, author_list[i])
        worksheet.write_string(i + 1, 2, price_list[i])
        worksheet.write_string(i + 1, 3, website_list[i])
        worksheet.write_string(i + 1, 4, email_list[i])

    # Closes the Excel workbook
    workbook.close()

# A method that gets the app's URL prefix based on the app's name
def get_url_prefix(name, chat_app_list):

    # if a chat app then it needs to go to the URL for chat apps
    if name in chat_app_list:
        url_prefix = 'https://www.zendesk.com/apps/chat/'

    # if the name does not indicate the app is a chat app then it needs to go to the URL for support apps
    else:
        url_prefix = 'https://www.zendesk.com/apps/'

    return url_prefix


# A method that removes whitespace from an app's name and appropriately replaces the whitespace with another character
def remove_whitespace_from_name(name):

    # the variable that will contain the newly formatted app name without whitespace
    formatted_name = name

    # if there is a space in the app's name, then replace the space with a hyphen
    if ' ' in name:
        formatted_name = formatted_name.replace(' ', '-')

    # if there is a period in the app's name, then remove it
    if '.' in name:
        formatted_name = formatted_name.replace('.', '')

    # if there is an ampersand in the app's name, then remove it
    if '&' in name:
        formatted_name = formatted_name.replace('&', '')

    # if there is an exclamation point in the app's name, then remove it
    if '!' in name:
        formatted_name = formatted_name.replace('!', '')

    # if there is an apostrophe in the app's name, then remove it
    if "'" in name:
        formatted_name = formatted_name.replace("'", '')

    # if there is a colon in the app's name, then remove it
    if ':' in name:
        formatted_name = formatted_name.replace(':', '')

    # if there is an unformatted apostrophe in the app's name, then remove it
    if 'â€™' in name:
        formatted_name = formatted_name.replace('â€™', '')

    # if there is a trademark symbol in the app's name, then remove it
    if '\u2122' in name:
        formatted_name = formatted_name.replace('\u2122', '')

    # if there is a weird left apostrophe in the app's name, then remove it
    if '‘' in name:
        formatted_name = formatted_name.replace('‘', '')

    # if there is a weird left apostrophe in the app's name, then remove it
    if '‘' in name:
        formatted_name = formatted_name.replace('‘', '')

    # if there is a weird right apostrophe in the app's name, then remove it
    if '’' in name:
        formatted_name = formatted_name.replace('’', '')

    # if there is a weird apostrophe in the app's name, then remove it
    if 'â€˜' in name:
        formatted_name = formatted_name.replace('â€˜', '')

    # if there is a weird trademark in the app's name, then remove it
    if 'â„¢' in name:
        formatted_name = formatted_name.replace('â„¢', '')

    # if there is a left parenthesis in the app's name, then remove it
    if '(' in name:
        formatted_name = formatted_name.replace('(', '')

    # if there is a right parenthesis in the app's name, then remove it
    if ')' in name:
        formatted_name = formatted_name.replace(')', '')

    return formatted_name

# A function that checks if an HTML element is loaded on the app's page
def checkLoadElement(element_xpath, driver, wait_time):
    count = 0
    while count < wait_time:
        #grabs element
        try:
            element = driver.find_element_by_xpath(element_xpath)
            return True

        # if element does not exist yet, sleep for 1 second and try again
        except NoSuchElementException:
            count += 1
            time.sleep(1)

        # if element has not loaded in 10 seconds, return False
        if count == wait_time:
            return False

# A method to get the name of each app and return it as a list
def get_name_list():
    # Opens the file containing the needed HTML and puts the HTML into a String
    with open('html_file.txt', 'r') as html_file:
        # The String that holds the HTML to parse through
        html_string = html_file.read()

    # setting up the HTML parser
    soup = BeautifulSoup(html_string, "html.parser")

    # list that will hold the names of the apps
    name_list = soup.find_all("span", class_="app-title")

    # Storing only the name of the app in the name_list, not any other HTML surrounding the name
    name_list = list(map(lambda name: name.string, name_list))

    return name_list


# A method to get the URL of each app and return it as a list
def get_url_list(name_list, chat_app_list):
    # list that will hold the individual URL for each app's page on Zendesk; these URLs will be used to get the app's
    # price, author, email, and website
    url_list = []

    # for each app, designate its individual URL on Zendesk
    for i in range(len(name_list)):
        # The app's individual URL
        app_url = get_url_prefix(name_list[i], chat_app_list) + remove_whitespace_from_name(name_list[i]).lower()

        # Adding the app's individual URL to the list of URLs
        url_list.append(app_url)

    return url_list


# A method that sets up the selenium driver to retrieve app data
def get_selenium_driver():
    # options object to make Firefox browser headless
    options = Options()

    # setting Firefox browser to be headless
    options.headless = True

    # sets up Firefox browser to be opened; need to include options and executable_path
    driver = webdriver.Firefox(options=options, executable_path='C:\\Users\\12158\\Documents\\Shane\\geckodriver-v0.26.0-win64\\geckodriver.exe')

    return driver

# A method that gets a list of all chat apps
def get_chat_app_list(driver):
    # constant for the maximum amount of time spent checking if a webpage loaded
    MAX_WEBPAGE_LOAD_TIME = 5

    # constant for the maximum number of apps that can display on page before pressing more apps button
    MAX_APPS_ON_PAGE = 36

    # constant that holds the URL to the Zendesk chat apps
    CHAT_APPS_URL = 'https://www.zendesk.com/apps/directory/?q=&idx=appsIndex&p=0&hFR%5Bproducts%5D%5B0%5D=chat'

    # constant that holds the xpath that checks if the website loaded
    WEBSITE_LOAD_XPATH = '/html/body/div[1]/article/section[2]/div[2]/div[2]/div/ul/li[1]/a/figure/div/div'

    # constant that holds the xpath to the total number of chat apps
    TOTAL_CHAT_APPS_XPATH = '/html/body/div[1]/article/section[2]/div[2]/div[3]/div[1]/span[2]'

    # constant that holds the xpath of the button that loads more chat apps (if necessary)
    MORE_APPS_BUTTON_XPATH = '//*[@id="view-more-apps"]'

    # Going to the URL that contains all chat apps
    driver.get(CHAT_APPS_URL)

    # checking if the website loaded
    website_loaded = checkLoadElement(WEBSITE_LOAD_XPATH, driver, MAX_WEBPAGE_LOAD_TIME)

    # if the chat apps page loaded properly, load all of the chat apps
    # if website_loaded == True:


    # if the webpage has not loaded, indicate this
    # else:
        # print('Page did not load correctly. URL: ' + CHAT_APPS_URL)


# A method that gets data (author, price, website, email) about each app
def get_app_data(driver, url_list):
    # constant for the maximum amount of time spent checking if a webpage loaded
    MAX_WEBPAGE_LOAD_TIME = 5

    # constant for the maximum amount of time spent checking if an element on a webpage is present
    MAX_ELEMENT_LOAD_TIME = 2

    # constant that holds the xpath that checks if the website loaded
    WEBSITE_LOAD_XPATH = '/html/body/article/section[2]/div[2]/section[2]/section/section/ul/li[1]/span'

    # constant that holds the xpath of the author
    AUTHOR_XPATH = '/html/body/article/section[2]/div[2]/section[2]/section/section/ul/li[1]/span[2]'

    # constant that holds the xpath of the price
    PRICE_XPATH = '/html/body/article/section[2]/div[2]/section[2]/section/section/ul/li[2]/span[2]'

    # constant that holds the xpath of the website
    WEBSITE_XPATH = '/html/body/article/section[2]/div[2]/section[2]/section/section/ul/li[3]/span[2]/a[2]'

    # constant that holds the xpath of the email
    EMAIL_XPATH = '/html/body/article/section[2]/div[2]/section[2]/section/section/ul/li[3]/span[2]/a[1]'

    # list that contains the author of each app
    author_list = []

    # list that contains the price of each app
    price_list = []

    # list that contains the website of each app
    website_list = []

    # list that contains the email of each app
    email_list = []

    # for each app, retrieve its author, price, email, and website
    for i in range(len(url_list)):
        # gets the app page in Firefox
        driver.get(url_list[i])

        # checking if the website loaded
        website_loaded = checkLoadElement(WEBSITE_LOAD_XPATH, driver, MAX_WEBPAGE_LOAD_TIME)

        # if the app's page loaded properly, look for the author, price, website, and email
        if website_loaded == True:

            # checking if the app has an author
            author_loaded = checkLoadElement(AUTHOR_XPATH, driver, MAX_ELEMENT_LOAD_TIME)

            # if the app has an author, add it to the list of authors
            if author_loaded == True:
                author_list.append(driver.find_element_by_xpath(AUTHOR_XPATH).text)

            # if the app does not have an author, indicate this in the list of authors
            else:
                author_list.append("No author listed")

            # checking if the app has a price
            price_loaded = checkLoadElement(PRICE_XPATH, driver, MAX_ELEMENT_LOAD_TIME)

            # if the app has a price, add it to the list of prices
            if price_loaded == True:
                price_list.append(driver.find_element_by_xpath(PRICE_XPATH).text)

            # if the app does not have a price, indicate this in the list of prices
            else:
                price_list.append("No price listed")

            # checking if the app has a website
            website_loaded = checkLoadElement(WEBSITE_XPATH, driver, MAX_ELEMENT_LOAD_TIME)

            # if the app has a website, add it to the list of websites
            if website_loaded == True:
                website_list.append(driver.find_element_by_xpath(WEBSITE_XPATH).get_attribute('href'))

            # if the app does not have a website, indicate this in the list of websites
            else:
                website_list.append("No website listed")

            # checking if the app has an email
            email_loaded = checkLoadElement(EMAIL_XPATH, driver, MAX_ELEMENT_LOAD_TIME)

            # if the app has an email, add it to the list of emails
            if email_loaded == True:
                email_list.append(driver.find_element_by_xpath(EMAIL_XPATH).get_attribute('href'))

            # if the app does not have an email, indicate this in the list of emails
            else:
                email_list.append("No email listed")

        # if the webpage has not loaded, indicate this
        else:
            print('Page did not load correctly. URL: ' + str(url_list[i]))

        print(author_list[i])
        print(price_list[i])
        print(website_list[i])
        print(email_list[i])

    # close the selenium driver
    driver.quit()

    return author_list, price_list, website_list, email_list

# The main method that will drive the execution of the script
def get_zendesk_apps_info():

    # get the name of each app
    name_list = get_name_list()

    print(name_list)

    # get the driver that will be used to get the app data
    driver = get_selenium_driver()

    chat_app_list = get_chat_app_list(driver)

    print(chat_app_list)

    # get the URL of each app
    url_list = get_url_list(name_list, chat_app_list)

    print(url_list)

    # gets the author, price, website, and email for each app and puts them in their own list
    author_list, price_list, website_list, email_list = get_app_data(driver, url_list)

    print(author_list)
    print(price_list)
    print(website_list)
    print(email_list)

    # creates Excel sheet with all data for all Zendesk apps
    create_excel_spreadsheet(name_list, author_list, price_list, website_list, email_list)

if __name__ == "__main__":
    get_zendesk_apps_info()