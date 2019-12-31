import time

from selenium import webdriver
from bs4 import BeautifulSoup
import xlsxwriter
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.firefox.options import Options
import sys
import urllib.request

# Method that will create the Excel sheet
def create_excel_spreadsheet(name_list, icon_link_list, author_list, price_list, website_list, email_list, chat_app_list, analytics_and_reporting_app_list, cti_providers_app_list, channels_app_list, collaboration_app_list, compose_and_edit_app_list, ecommerce_and_crm_app_list, email_and_social_media_app_list, it_and_project_management_app_list, knowledge_and_content_app_list, productivity_and_time_tracking_app_list, surveys_and_feedback_app_list, telephony_and_sms_app_list, zendesk_labs_app_list):

    # path to the Excel file that will be written
    path = 'PATH/TO/EXCEL/sheet.xlsx'

    # list that will hold the names of the headers of the Excel file
    header_list = ['Name', 'Icon Link', 'Type', 'Price', 'Author', 'Website', 'Email', 'Analytics & Reporting', 'CTI Providers', 'Channels', 'Collaboration', 'Compose & Edit', 'E-commerce & CRM', 'Email & Social Media', 'IT & Product Management', 'Knowledge & Content', 'Productivity & Time-tracking', 'Surveys & Feedback', 'Telephony & SMS', 'Zendesk Labs']

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

    # a constant that indicates the column for the name
    NAME_COLUMN = 0

    # a constant that indicates the column for the icon link
    ICON_LINK_COLUMN = 1

    # a constant that indicates the column for the type
    TYPE_COLUMN = 2

    # a constant that indicates the column for the price
    PRICE_COLUMN = 3

    # a constant that indicates the column for the author
    AUTHOR_COLUMN = 4

    # a constant that indicates the column for the website
    WEBSITE_COLUMN = 5

    # a constant that indicates the column for the email
    EMAIL_COLUMN = 6

    # a constant that indicates the column for the analytics and reporting
    ANALYTICS_AND_REPORTING_COLUMN = 7

    # a constant that indicates the column for the CTI providers
    CTI_PROVIDERS_COLUMN = 8

    # a constant that indicates the column for the channels
    CHANNELS_COLUMN = 9

    # a constant that indicates the column for the collaboration
    COLLABORATION_COLUMN = 10

    # a constant that indicates the column for the compose and edit
    COMPOSE_AND_EDIT_COLUMN = 11

    # a constant that indicates the column for the ecommerce and CRM
    ECOMMERCE_AND_CRM_COLUMN = 12

    # a constant that indicates the column for the email and social media
    EMAIL_AND_SOCIAL_MEDIA_COLUMN = 13

    # a constant that indicates the column for the IT and project management
    IT_AND_PROJECT_MANAGEMENT_COLUMN = 14

    # a constant that indicates the column for the knowledge and content
    KNOWLEDGE_AND_CONTENT_COLUMN = 15

    # a constant that indicates the column for the productivity and time tracking
    PRODUCTIVITY_AND_TIME_TRACKING_COLUMN = 16

    # a constant that indicates the column for the surveys and feedback
    SURVEYS_AND_FEEDBACK_COLUMN = 17

    # a constant that indicates the column for the telephony and sms
    TELEPHONY_AND_SMS_COLUMN = 18

    # a constant that indicates the column for the zendesk labs
    ZENDESK_LABS_COLUMN = 19

    # writing the app name, author, price, website, and email into each row
    for i in range(len(name_list)):

        # the row of the Excel sheet to write into (constant throughout for loop)
        row = i + 1

        worksheet.write_string(row, NAME_COLUMN, name_list[i])
        worksheet.write_string(row, ICON_LINK_COLUMN, icon_link_list[i])
        worksheet.write_string(row, PRICE_COLUMN, price_list[i])
        worksheet.write_string(row, AUTHOR_COLUMN, author_list[i])
        worksheet.write_string(row, WEBSITE_COLUMN, website_list[i])
        worksheet.write_string(row, EMAIL_COLUMN, email_list[i])

        # if this is a chat app, indicate so in the Excel sheet
        if name_list[i] in chat_app_list:
            worksheet.write_string(row, TYPE_COLUMN, "chat")

        # otherwise, indicate that it is a support app
        else:
            worksheet.write_string(row, TYPE_COLUMN, "support")

        if name_list[i] in analytics_and_reporting_app_list:
            worksheet.write_string(row, ANALYTICS_AND_REPORTING_COLUMN, '\u2713')

        if name_list[i] in cti_providers_app_list:
            worksheet.write_string(row, CTI_PROVIDERS_COLUMN, '\u2713')

        if name_list[i] in channels_app_list:
            worksheet.write_string(row, CHANNELS_COLUMN, '\u2713')

        if name_list[i] in collaboration_app_list:
            worksheet.write_string(row, COLLABORATION_COLUMN, '\u2713')

        if name_list[i] in compose_and_edit_app_list:
            worksheet.write_string(row, COMPOSE_AND_EDIT_COLUMN, '\u2713')

        if name_list[i] in ecommerce_and_crm_app_list:
            worksheet.write_string(row, ECOMMERCE_AND_CRM_COLUMN, '\u2713')

        if name_list[i] in email_and_social_media_app_list:
            worksheet.write_string(row, EMAIL_AND_SOCIAL_MEDIA_COLUMN, '\u2713')

        if name_list[i] in it_and_project_management_app_list:
            worksheet.write_string(row, IT_AND_PROJECT_MANAGEMENT_COLUMN, '\u2713')

        if name_list[i] in knowledge_and_content_app_list:
            worksheet.write_string(row, KNOWLEDGE_AND_CONTENT_COLUMN, '\u2713')

        if name_list[i] in productivity_and_time_tracking_app_list:
            worksheet.write_string(row, PRODUCTIVITY_AND_TIME_TRACKING_COLUMN, '\u2713')

        if name_list[i] in surveys_and_feedback_app_list:
            worksheet.write_string(row, SURVEYS_AND_FEEDBACK_COLUMN, '\u2713')

        if name_list[i] in telephony_and_sms_app_list:
            worksheet.write_string(row, TELEPHONY_AND_SMS_COLUMN, '\u2713')

        if name_list[i] in zendesk_labs_app_list:
            worksheet.write_string(row, ZENDESK_LABS_COLUMN, '\u2713')

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

    # if there is a restricted character in the app's name, then remove it
    if 'Â®' in name:
        formatted_name = formatted_name.replace('Â®', '')

    # if there is an at character in the app's name, then remove it
    if '@' in name:
        formatted_name = formatted_name.replace('@', '')

    # if there is an a forward slash in the app's name, then remove it
    if '/' in name:
        formatted_name = formatted_name.replace('/', '')

    # if there is an accented a in the app's name, then remove it
    if 'Ã¡' in name:
        formatted_name = formatted_name.replace('Ã¡', '')

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

# A method to get the link of each app icon and return it as a list
def get_icon_link_list():

    # Opens the file containing the needed HTML and puts the HTML into a String
    with open('html_file.txt', 'r') as html_file:
        # The String that holds the HTML to parse through
        html_string = html_file.read()

    # setting up the HTML parser
    soup = BeautifulSoup(html_string, "html.parser")

    # list that will hold the icon links of the apps
    icon_link_list = soup.find_all("div", class_="lazyloaded")

    # Storing only the icon link of the app in the icon_link_list, not any other HTML surrounding the icon link
    icon_link_list = list(map(lambda icon_link: icon_link.get('data-bgset'), icon_link_list))

    return icon_link_list

# A method that retrieves and downloads the icons to the correct location
def get_icons(name_list, icon_link_list):

    # sets up object to open a URL
    opener = urllib.request.URLopener()

    # user agent needed so 403 error does not occur
    opener.addheader('User-Agent', 'Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.9.0.7) Gecko/2009021910 Firefox/3.0.7')

    # Go through each icon link and download the icon
    for i in range(len(icon_link_list)):

        # Grab each icon and save it to the specified folder
        opener.retrieve(icon_link_list[i], '/PATH/TO/APP/ICONS' + str(name_list[i]) + '.png')

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
    #
    options.headless = True

    # sets up Firefox browser to be opened; need to include options and executable_path
    driver = webdriver.Firefox(options=options, executable_path='PATH/TO/geckodriver.exe')

    return driver

# A method that gets a list of all apps at a specific URL
def get_app_list(url):
    # constant for the maximum amount of time spent checking if a webpage loaded
    MAX_WEBPAGE_LOAD_TIME = 5

    # constant for the maximum number of apps that can display on page before pressing more apps button
    MAX_APPS_ON_PAGE = 36

    # constant that holds the xpath that checks if the website loaded
    WEBSITE_LOAD_XPATH = '/html/body/div[1]/article/section[2]/div[2]/div[2]/div/ul/li[1]/a/figure/div/div'

    # constant that holds the xpath to the total number of apps
    TOTAL_APPS_XPATH = '/html/body/div[1]/article/section[2]/div[2]/div[3]/div[1]/span[2]'

    # the app list that will be returned
    app_list = []

    # get the driver that will be used to get the app data
    driver = get_selenium_driver()

    # Going to the URL that contains apps
    driver.get(url)

    # checking if the website loaded
    website_loaded = checkLoadElement(WEBSITE_LOAD_XPATH, driver, MAX_WEBPAGE_LOAD_TIME)

    # if the apps page loaded properly, load all of the apps, and get the name of each app
    if website_loaded == True:

        # The total number of apps will have loaded by this point but it may not be fully functional, False by default
        total_loaded_correctly = False

        # counting the number of times there has been a ValueError, if more than 10, then likely Zendesk Labs page
        count = 0

        # maximum times ValueError can occur
        max_count = 10

        # Wait for the total to load right
        while (not total_loaded_correctly) and count < max_count:
            try:
                # the total number of apps
                total_apps = int(driver.find_element_by_xpath(TOTAL_APPS_XPATH).text)

                total_loaded_correctly = True

            except ValueError:
                # sleep 1 second to make sure everything loads
                time.sleep(1)
                count += 1

        # if it took too many tries to try to find a value, it likely doesn't exist because it is the Zendesk Labs page
        # so just grab the total for Zendesk Labs at another location
        if count >= max_count:
            # the total number of apps for Zendesk Labs
            total_apps = int(driver.find_element_by_xpath('/html/body/div[1]/article/section[2]/div[2]/section/div[1]/div[3]/div[2]/div/div/div/div/div[13]/div/label/span').text)

        # load all of the chat apps
        load_all_apps(driver, MAX_APPS_ON_PAGE, total_apps)


        # for every app, get its name
        for i in range(total_apps):

            # The app will have loaded by this point but it may not be fully functional, False by default
            app_loaded_correctly = False

            # Wait for the app to load right
            while (not app_loaded_correctly):
                try:
                    # add to the list the name of the app
                    app_list.append(driver.find_element_by_xpath('/html/body/div[1]/article/section[2]/div[2]/div[2]/div[' + str(int(i / MAX_APPS_ON_PAGE) + 1) + ']/ul/li[' + str(int(i % MAX_APPS_ON_PAGE) + 1) + ']/a/div/span[1]').text)

                    app_loaded_correctly = True

                except NoSuchElementException:
                    # sleep 1 second to make sure everything loads
                    time.sleep(1)

        # quit driver
        driver.quit()

    # if the webpage has not loaded, indicate this
    else:
        # quit driver
        driver.quit()

        sys.exit('Page did not load correctly. URL: ' + url)

    return app_list


# A method that properly loads all apps on the apps webpage
def load_all_apps(driver, MAX_APPS_ON_PAGE, total_chat_apps):
    # constant for the maximum amount of time spent checking if an element on a webpage is present
    MAX_ELEMENT_LOAD_TIME = 3

    # constant that holds the xpath of the button that loads more chat apps (if necessary)
    MORE_APPS_BUTTON_XPATH = '//*[@id="view-more-apps"]'

    # the total number of times the more apps button should be pressed based on how many apps there are
    num_of_times_to_press_more_apps_button = int(total_chat_apps / MAX_APPS_ON_PAGE)

    # press the more apps button the number of times it should be pressed
    for i in range(num_of_times_to_press_more_apps_button):

        # check if more apps button loaded
        more_apps_button_loaded = checkLoadElement(MORE_APPS_BUTTON_XPATH, driver, MAX_ELEMENT_LOAD_TIME)

        # if the more apps button loaded, then press it and continue
        if more_apps_button_loaded == True:

            # The button will have loaded by this point but it may not be fully functional, False by default
            button_loaded_correctly = False

            # Wait for the button to load right
            while(not button_loaded_correctly):
                try:

                    # sleep 2 seconds to allow the button to load correctly
                    time.sleep(2)

                    # the more apps button
                    more_apps_button = driver.find_element_by_xpath(MORE_APPS_BUTTON_XPATH)

                    # click on the more apps button
                    more_apps_button.click()

                    button_loaded_correctly = True

                except NoSuchElementException:
                    # sleep 1 second to allow the button to load
                    time.sleep(1)

        # if the more apps button has not loaded, indicate this
        else:
            sys.exit('More apps button did not load correctly.')


# A method that gets data (author, price, website, email) about each app
def get_app_data(url_list):
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

        # get the driver that will be used to get the app data
        driver = get_selenium_driver()

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

        # These app URLs do not work because Zendesk's website is a bit broken
        elif str(url_list[i]) == 'https://www.zendesk.com/apps/chat/mindsay-for-zendesk-chat' or str(url_list[i]) == 'https://www.zendesk.com/apps/mindsay-for-zendesk-support' or str(url_list[i] == 'https://www.zendesk.com/apps/100worte'):
            author_list.append("No author listed")
            price_list.append("No price listed")
            website_list.append("No website listed")
            email_list.append("No email listed")

        # if the webpage has not loaded, indicate this
        else:
            sys.exit('Page did not load correctly. URL: ' + str(url_list[i]))

        print('Retrieved author, price, website, and email of app #' + str(i + 1) + '.')

        # quit the selenium driver
        driver.quit()

    return author_list, price_list, website_list, email_list

# A method that gets a list for every category that an app could belong to
def get_app_category_lists():

    # a constant for the URL to the chat apps
    CHAT_APPS_URL = 'https://www.zendesk.com/apps/directory/?q=&idx=appsIndex&p=0&hFR%5Bproducts%5D%5B0%5D=chat'

    # a constant for the URL to the analytics and reporting apps
    ANALYTICS_AND_REPORTING_APPS_URL = 'https://www.zendesk.com/apps/directory/?q=&idx=appsIndex&p=0&dFR%5Bcategories.name%5D%5B0%5D=Analytics%20%26%20Reporting'

    # a constant for the URL to the CTI providers apps
    CTI_PROVIDERS_APPS_URL = 'https://www.zendesk.com/apps/directory/?q=&idx=appsIndex&p=0&dFR%5Bcategories.name%5D%5B0%5D=CTI%20Providers'

    # a constant for the URL to the channels apps
    CHANNELS_APPS_URL = 'https://www.zendesk.com/apps/directory/?q=&idx=appsIndex&p=0&dFR%5Bcategories.name%5D%5B0%5D=Channels'

    # a constant for the URL to the collaboration apps
    COLLABORATION_APPS_URL = 'https://www.zendesk.com/apps/directory/?q=&idx=appsIndex&p=0&dFR%5Bcategories.name%5D%5B0%5D=Collaboration'

    # a constant for the compose and edit to the chat apps
    COMPOSE_AND_EDIT_APPS_URL = 'https://www.zendesk.com/apps/directory/?q=&idx=appsIndex&p=0&dFR%5Bcategories.name%5D%5B0%5D=Compose%20%26%20Edit'

    # a constant for the URL to the ecommerce and CRM apps
    ECOMMERCE_AND_CRM_APPS_URL = 'https://www.zendesk.com/apps/directory/?q=&idx=appsIndex&p=0&dFR%5Bcategories.name%5D%5B0%5D=E-commerce%20%26%20CRM'

    # a constant for the URL to the email and social media apps
    EMAIL_AND_SOCIAL_MEDIA_APPS_URL = 'https://www.zendesk.com/apps/directory/?q=&idx=appsIndex&p=0&dFR%5Bcategories.name%5D%5B0%5D=Email%20%26%20Social%20Media'

    # a constant for the URL to the IT and project management apps
    IT_AND_PROJECT_MANAGEMENT_APPS_URL = 'https://www.zendesk.com/apps/directory/?q=&idx=appsIndex&p=0&dFR%5Bcategories.name%5D%5B0%5D=IT%20%26%20Project%20Management'

    # a constant for the URL to the knowledge and content apps
    KNOWLEDGE_AND_CONTENT_APPS_URL = 'https://www.zendesk.com/apps/directory/?q=&idx=appsIndex&p=0&dFR%5Bcategories.name%5D%5B0%5D=Knowledge%20%26%20Content'

    # a constant for the URL to the productivity and time tracking apps
    PRODUCTIVITY_AND_TIME_TRACKING_APPS_URL = 'https://www.zendesk.com/apps/directory/?q=&idx=appsIndex&p=0&dFR%5Bcategories.name%5D%5B0%5D=Productivity%20%26%20Time-tracking'

    # a constant for the URL to the surveys and feedback apps
    SURVEYS_AND_FEEDBACK_APPS_URL = 'https://www.zendesk.com/apps/directory/?q=&idx=appsIndex&p=0&dFR%5Bcategories.name%5D%5B0%5D=Surveys%20%26%20Feedback'

    # a constant for the URL to the telephony and SMS apps
    TELEPHONY_AND_SMS_APPS_URL = 'https://www.zendesk.com/apps/directory/?q=&idx=appsIndex&p=0&dFR%5Bcategories.name%5D%5B0%5D=Telephony%20%26%20SMS'

    # a constant for the URL to the zendesk labs apps
    ZENDESK_LABS_APPS_URL = 'https://www.zendesk.com/apps/directory/?q=&idx=appsIndex&p=0&dFR%5Bcategories.name%5D%5B0%5D=Zendesk%20Labs'

    chat_app_list = get_app_list(CHAT_APPS_URL)

    print("Chat list completed.")

    analytics_and_reporting_app_list = get_app_list(ANALYTICS_AND_REPORTING_APPS_URL)

    print("Analytics & Reporting list completed.")

    cti_providers_app_list = get_app_list(CTI_PROVIDERS_APPS_URL)

    print("CTI Providers list completed.")

    channels_app_list = get_app_list(CHANNELS_APPS_URL)

    print("Channels list completed.")

    collaboration_app_list = get_app_list(COLLABORATION_APPS_URL)

    print("Collaboration list completed.")

    compose_and_edit_app_list = get_app_list(COMPOSE_AND_EDIT_APPS_URL)

    print("Compose & Edit list completed.")

    ecommerce_and_crm_app_list = get_app_list(ECOMMERCE_AND_CRM_APPS_URL)

    print("E-commerce & CRM list completed.")

    email_and_social_media_app_list = get_app_list(EMAIL_AND_SOCIAL_MEDIA_APPS_URL)

    print("Email & Social Media list completed.")

    it_and_project_management_app_list = get_app_list(IT_AND_PROJECT_MANAGEMENT_APPS_URL)

    print("IT & Project Management list completed.")

    knowledge_and_content_app_list = get_app_list(KNOWLEDGE_AND_CONTENT_APPS_URL)

    print("Knowledge & Content list completed.")

    productivity_and_time_tracking_app_list = get_app_list(PRODUCTIVITY_AND_TIME_TRACKING_APPS_URL)

    print("Productivity & Time Tracking list completed.")

    surveys_and_feedback_app_list = get_app_list(SURVEYS_AND_FEEDBACK_APPS_URL)

    print("Surveys & Feedback list completed.")

    telephony_and_sms_app_list = get_app_list(TELEPHONY_AND_SMS_APPS_URL)

    print("Telephony & SMS list completed.")

    zendesk_labs_app_list = get_app_list(ZENDESK_LABS_APPS_URL)

    print("Zendesk Labs list completed.")

    return chat_app_list, analytics_and_reporting_app_list, cti_providers_app_list, channels_app_list, collaboration_app_list, compose_and_edit_app_list, ecommerce_and_crm_app_list, email_and_social_media_app_list, it_and_project_management_app_list, knowledge_and_content_app_list, productivity_and_time_tracking_app_list, surveys_and_feedback_app_list, telephony_and_sms_app_list, zendesk_labs_app_list

# The main method that will drive the execution of the script
def get_zendesk_apps_info():

    print('Getting app names...')

    # get the name of each app
    name_list = get_name_list()

    print('Retrieved app names.\nGetting app icon URLs...')

    # get the icon link of each app
    icon_link_list = get_icon_link_list()

    print('Retrieved app icon URLs.\n Saving app icons...')

    # get the icon of each app
    get_icons(name_list, icon_link_list)

    print('Saved app icons.')

    # getting lists that have the name of each app that belongs to the specific category list
    chat_app_list, analytics_and_reporting_app_list, cti_providers_app_list, channels_app_list, collaboration_app_list, compose_and_edit_app_list, ecommerce_and_crm_app_list, email_and_social_media_app_list, it_and_project_management_app_list, knowledge_and_content_app_list, productivity_and_time_tracking_app_list, surveys_and_feedback_app_list, telephony_and_sms_app_list, zendesk_labs_app_list = get_app_category_lists()

    print('Getting app URLs...')

    # get the URL of each app
    url_list = get_url_list(name_list, chat_app_list)

    print('Retrieved app URLs.\nGetting author, price, website, and email info for each app...')

    # gets the author, price, website, and email for each app and puts them in their own list
    author_list, price_list, website_list, email_list = get_app_data(url_list)

    print('Retrieved author, price, website, and email info for each app...\nCreating Excel spreadsheet...')

    # creates Excel sheet with all data for all Zendesk apps
    create_excel_spreadsheet(name_list, icon_link_list, author_list, price_list, website_list, email_list, chat_app_list, analytics_and_reporting_app_list, cti_providers_app_list, channels_app_list, collaboration_app_list, compose_and_edit_app_list, ecommerce_and_crm_app_list, email_and_social_media_app_list, it_and_project_management_app_list, knowledge_and_content_app_list, productivity_and_time_tracking_app_list, surveys_and_feedback_app_list, telephony_and_sms_app_list, zendesk_labs_app_list)

    print('Finished creating Excel spreadsheet.\nFinished!')

if __name__ == "__main__":
    get_zendesk_apps_info()