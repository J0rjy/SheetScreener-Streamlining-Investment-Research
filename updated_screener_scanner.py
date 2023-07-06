# The sheet can be accessed with this link: https://docs.google.com/spreadsheets/d/198Yvl0OtqVG_uArcJC_AwrG9BPO6tF98Qk2Lx08rAd0/edit#gid=0

# Requirements
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from urllib.parse import urlparse
from bs4 import BeautifulSoup
import pandas as pd
from selenium.webdriver.support.ui import Select
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials



#Functions
def login(useername, password):
 """
 Auto-login the screener platform

 Args: Username(email) and Password
 """
 # Find the email and password input fields and fill them in
 email_input = driver.find_element("name", "username")
 email_input.send_keys(username)  # Replace with your email
 password_input = driver.find_element("name", "password")
 password_input.send_keys(password)  # Replace with your password
 
 # Submit the form
 password_input.send_keys(Keys.RETURN)
 
 
def extract_company_url(url):
    """
    Extracts the Company specific Ticker URl part from the Screener company URL

    Args: company URL from Screener
    """
    parsed_url = urlparse(url)
    path_segments = parsed_url.path.split('/')
    company_url = '/'.join(segment for segment in path_segments if segment)
    return company_url

def ratio_clicker_fetcher(company_url):
  '''
  Fetches and clicks on the custom ratios
  So that they can be accessed for dataframe creation

  Args: Company URl: but only the Ticker name and consolidated part from a screener URL
  eg. : company/TCS/consolidated is taken from https://www.screener.in/company/TCS/consolidated/ 
  this is handled by extract_company_url() function
  
  '''
  custom_ratios_url = "https://www.screener.in/user/quick_ratios/?next=/"+company_url
  driver.get(custom_ratios_url)
  #clicking on save quick ratios button
  custom_ratios_btn = driver.find_element("xpath", "/html/body/main/div[2]/form/div[1]/div/button").click()
  company_ratios = driver.find_element("xpath", "//*[@id=\"top-ratios\"]")
  
  return company_ratios

def ratios_extractor(ratios):
 """
 Returns the ratios in the form of a dictionary

 Args: Company ratios in the form of xpath element instance 
 """
 html_code = ratios.get_attribute("innerHTML")

 # Parse the HTML
 soup = BeautifulSoup(html_code, 'html.parser')

 # Find all the list items ('li') with class 'flex flex-space-between'
 list_items = soup.find_all('li', class_='flex flex-space-between')

 # Create an empty dictionary to store the parsed data
 stock_data = {}
 
 # Iterate over the list items and extract the cost and metrics
 for item in list_items:
     metric_name = item.find('span', class_='name').text.strip()
     metric_value = item.find('span', class_='number').text.strip()
     stock_data[metric_name] = metric_value
 return stock_data

def extract_company_name(company_url):
  """Extracts the company name from the given URL."""

  match = re.search(r'company/(.*)/consolidated/', url)
  if match:
    return match.group(1)
  else:
    return None

# Setting up credentials for google account
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('courserapractice-387618-aadc3f1682bb.json', scope)  

# Authorize the client
client = gspread.authorize(credentials)

# Open the Google Sheets document
spreadsheet = client.open('Screener_extracted_sheet') 

#This loop runs the program
while True:
    #Setting up the driver
    options = Options()
    options.add_experimental_option("detach", True)
    #Initializing the driver
    driver = webdriver.Chrome(service = Service(ChromeDriverManager().install()),options=options)
    #Fetching the page
    driver.get("https://www.screener.in/login/")
    #Fetching the page
    driver.get("https://www.screener.in/login/")
    #initializing the security keys
    username = "Supremeburner007@gmail.com"
    password = "Zoro@1029"
    login(username, password)
    
    #click on the search box and serach for the stock now just copy the stocks url and paste it
    stock_screener_url = input("Paste the screener url of the stock: ")
    driver.get(stock_screener_url)
    
    url = stock_screener_url
    company_url = extract_company_url(url)
    
    #Stores all the company ratios
    ratios = ratio_clicker_fetcher(company_url)
    
    #Stores the company ratios in the form of a dictionary
    stock_data_dict = ratios_extractor(ratios)
    
    #converts the dataframe into a pandas Dataframe
    df = pd.DataFrame(list(stock_data_dict.items()),columns = ['Metric','Value'])
    final_df = df.drop([df.index[0], df.index[1], df.index[2], df.index[3], df.index[4], df.index[5], df.index[8]])
    final_df.reset_index(drop=True, inplace=True)
    
     # Extract the company name
    company_name = extract_company_name(company_url)

    # Prepare the data to be written to the worksheet
    data = [[company_name] + final_df['Value'].tolist()]
    worksheet = spreadsheet.get_worksheet(0) 
    # Find the next empty row in the worksheet
    next_row = len(worksheet.col_values(1)) + 1

    # Write the data to the worksheet, starting from the next empty row
    worksheet.update(f'A{next_row}', data)

    # Print the output
    print("Data successfully extracted to Google Sheets!")

    # Ask the user if they want to continue
    choice = input("Do you want to extract more data? (y/n): ")
    if choice.lower() == 'n':
        print("All the changes are saved to : https://docs.google.com/spreadsheets/d/198Yvl0OtqVG_uArcJC_AwrG9BPO6tF98Qk2Lx08rAd0/" )
        break