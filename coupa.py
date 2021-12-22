import time
import pandas as pd
import re
import os
from bs4 import BeautifulSoup
from selenium import webdriver
from credentials import USERNAME, PASSWORD

class Coupa:
    def __init__(self, workbook: str, driver_path: str, dir_path: str) -> None:
        self.username = USERNAME                                                           # Username from the credentials file
        self.password = PASSWORD                                                           # Password from the credentials file
        self.workbook = workbook                                                           # Path to the Excel workbook that holds the SAP Document Numbers and coupa URLs
        self.doc_to_url = dict()                                                           # Dictionary for the SAP Document Numbers and their respective invoice URLs 
        self.dir_path = dir_path                                                           # Directory where the PDF files are located 
        self.options = webdriver.ChromeOptions()                                           # Standard for using Chrome with selenium
        self.options.add_experimental_option('excludeSwitches', ['enable-logging'])        # Get rid of annoying messages related to "DevTools listening on..." 
        self.driver = webdriver.Chrome(executable_path=driver_path, options=self.options)  # Create webdriver for Chrome
    
    def __repr__(self) -> str:
        return 'This script downloads invoices (if available) from coupa URLs and saves them to the default directory (e.g., Downloads folder)'

    def read_excel(self) -> None:
        # Save the contents of the Excel workbook to a pandas DataFrame
        self.df = pd.read_excel(self.workbook)

        # Set this as the initial URL to go to for ease of use
        self.beginning_url = self.df['COUPA URLs'][0] 

    def initial_login(self) -> None:
        # Go to the initial URL
        self.driver.get(url=self.beginning_url)

        time.sleep(4)

        # Log in field
        login = self.driver.find_element_by_xpath(xpath='//*[@id="i0116"]')

        # Type username
        login.send_keys(USERNAME)

        time.sleep(2)

        # Next button
        self.driver.find_element_by_xpath(xpath='//*[@id="idSIButton9"]').click()

        time.sleep(3)

        # Password field
        password = self.driver.find_element_by_xpath(xpath='//*[@id="i0118"]')

        # Type password
        password.send_keys(PASSWORD)

        time.sleep(2)

        # Sign in button
        self.driver.find_element_by_xpath(xpath='//*[@id="idSIButton9"]').click()

        time.sleep(3)

        # Clicking "Yes" to "Stay signed in?"
        self.driver.find_element_by_xpath(xpath='//*[@id="idSIButton9"]').click()

        time.sleep(6)

    def main(self) -> None:
        # Iterate over the coupa URLs
        for i in range(len(self.df['COUPA URLs'])):
            # Go to the coupa URL
            self.driver.get(url=self.df['COUPA URLs'][i])

            time.sleep(5)

            # Get the current page's HTML
            HTML = self.driver.page_source

            # Create a soup object
            soup = BeautifulSoup(HTML, 'html.parser')

            # Assign the invoice PDF's URL to a variable
            try:
                doc_url = soup.find('a', {'href': re.compile(r'https://REDACTED.coupahost.com/attachment/attachment_file/file/.+/.+(\.pdf|\.PDF)')})['href']
            except Exception as e:
                print(str(e))
                print(f"{i+1}: Could not find an invoice for SAP Document Number {self.df['Document Number'][i]}")
                time.sleep(1)
                continue

            # Add the document URL to the dictionary with its SAP Document Number as the value
            self.doc_to_url.setdefault(doc_url.split('/')[-1], str(self.df['Document Number'][i]))

            # Click the PDF file
            self.driver.find_element_by_xpath(xpath='//*[@id="topHalf"]/div/div[1]/div[1]/section/div[11]/span[2]/ul/li/a').click()
            print(f"{i+1}: Invoice downloaded from {self.df['COUPA URLs'][i]}")
          
            time.sleep(3)
    
        # Call the function that closes the web browser
        self.close_browser()

    def close_browser(self) -> None:
        # Close the browser
        self.driver.close()    

    def rename_pdfs(self) -> None:
        # Change directory to the Downloads folder
        os.chdir(path=self.dir_path)

        # Rename the PDF files that are from our download session
        for filename in os.listdir(path=self.dir_path):
            if filename.lower().endswith('.pdf') and filename in self.doc_to_url:
                os.rename(filename, self.doc_to_url.get(filename) + '.pdf')

# Instantiate the class and call the necessary functions
if __name__ == '__main__':
    c = Coupa(workbook=r'C:\Users\REDACTED\Desktop\Documents\Tax\Excel files\Coupa Testing.xlsx',
              driver_path=r'C:\Users\REDACTED\Desktop\Documents\Miscellaneous\python drivers\chromedriver.exe',
              dir_path=r'C:\Users\REDACTED\Downloads'
    )
    c.read_excel()
    c.initial_login()
    c.main()
    c.rename_pdfs()
