from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import time
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from bs4 import BeautifulSoup

# Create the main window
top = tk.Tk()

depot_numre = []
kurtagedf = None


def open_excel_file():
    global depot_numre
    global kurtagedf
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        # Process the selected file
        print("Selected Excel file:", file_path)

        # Read the Excel file using pandas
        excelfile = pd.read_excel(file_path, sheet_name="Review")

        # Extract values from column "B" (index 1) and rows 11 to 20
        depot_numre = excelfile.iloc[10:2000, 1].fillna(0).astype(
            str).tolist()  # 10:20 for rows 11 to 20, 1 for column "B"
        kurtageDKK = excelfile.iloc[10:2000, 28].fillna(0).astype(
            int).tolist()  # 10:20 for rows 11 to 20, 20 for column "U"
        kurtageUSD = excelfile.iloc[10:2000, 29].fillna(0).astype(
            int).tolist()  # 10:20 for rows 11 to 20, 21 for column "V"
        kurtageEUR = excelfile.iloc[10:2000, 30].fillna(0).astype(
            int).tolist()  # 10:20 for rows 11 to 20, 22 for column "W"

        tradingdata = {
            'depot_numre': depot_numre,
            'kurtageDKK': kurtageDKK,
            'kurtageUSD': kurtageUSD,
            'kurtageEUR': kurtageEUR
        }

        kurtagedf = pd.DataFrame(tradingdata)


# Create a button to open the file dialog
open_button = tk.Button(top, text="Open Excel File", command=open_excel_file)
open_button.pack(pady=20)

# Run the Tkinter event loop
top.mainloop()

# Replace with the path to your web browser driver executable
# Download ChromeDriver from https://sites.google.com/a/chromium.org/chromedriver/downloads
chrome_driver_path = 'C://Users//Emil//Miranova FMS AS//MN - Administration//Niels//Chromedriver'

# URL of Nordnet login page
login_url = 'https://classic.nordnet.se/cm/login.html'

# Replace with your Nordnet username and password
nordnet_username = 'mnfmsniehov'
nordnet_password = 'sNsI0Baz'

# Create a new instance of ChromeDriver
driver = webdriver.Chrome()

# Open the Nordnet login page
driver.get(login_url)

# Wait for the page to load
time.sleep(3)

# Find the username and password input fields and fill them
username_input = driver.find_element(By.NAME, 'username')
username_input.send_keys(nordnet_username)

password_input = driver.find_element(By.NAME, 'password')
password_input.send_keys(nordnet_password)

# Submit the form (login)
password_input.send_keys(Keys.ENTER)

# Wait for the login process to complete (you may adjust the time as needed)
time.sleep(15)

for index, row in kurtagedf.iterrows():
    depot_numre = row['depot_numre']
    kurtageDKK = row['kurtageDKK']
    kurtageUSD = row['kurtageUSD']
    kurtageEUR = row['kurtageEUR']

    if len(depot_numre) == 7 or len(depot_numre) == 8:

        depot_input = driver.find_element(By.NAME, 'depot')

        depot_input.send_keys(str(depot_numre))

        time.sleep(2)

        depot_input.send_keys(Keys.ENTER)

        time.sleep(2)

        order_link = driver.find_element(By.XPATH, "//a[contains(text(), 'Exchange')]")
        order_link.click()

        # Find the table element by its XPath
        xpath_expression = '//*[@id="content"]/div[4]/table/tbody'
        table_element = driver.find_element(By.XPATH, xpath_expression)

        # Find all rows (tr elements) in the table
        rows = table_element.find_elements(By.TAG_NAME, 'tr')

        # Initialize an empty list to store the data
        data = []

        # Loop through each row and extract the data from cells (td elements)
        for row in rows:
            cells = row.find_elements(By.TAG_NAME, 'td')
            row_data = [cell.text.strip() for cell in cells]
            data.append(row_data)

        import math

        columns = ['Column1', 'Currency', 'Ticker', 'Amount', 'DKK', 'Debit interest', 'Credit interest',
                   'Capitalised interest(DKK)']
        df = pd.DataFrame(data, columns=columns)
        df = df.drop(columns=["Column1"])

        for row_index in range(len(df)):
            for col in df.columns:
                cell_value = df.at[row_index, col]
                if isinstance(cell_value, str) and cell_value != '' and ',' in cell_value:
                    numeric_value = cell_value.replace(' ', '').replace(',', '.')
                    try:
                        value = float(numeric_value)
                        rounded_value = round(value)
                        df.at[row_index, col] = int(rounded_value)
                    except ValueError:
                        pass  # Ignore cells that cannot be converted to float

        df.set_index('Currency', inplace=True)

        # Locate the table element using its class name
        table_element = driver.find_element(By.CLASS_NAME, 'vaxla')

        # Extract the table HTML content
        table_html = table_element.get_attribute('outerHTML')

        soup = BeautifulSoup(table_html, 'html.parser')

        # Find the table rows
        rows = soup.find_all('tr')

        # Print the table rows
        for row in rows:
            cellsX = row.find_all(['th', 'td'])
            row_data = [cell.get_text(strip=True) for cell in cellsX]

        df1 = df
        df2 = df
        df3 = df

        tbody_element = driver.find_element(By.XPATH, '//table[@class="vaxla"]/tbody')

        # Get the HTML content of the tbody element
        tbody_html = tbody_element.get_attribute('outerHTML')

        # Now you can use BeautifulSoup to parse the HTML and extract the table data
        soup = BeautifulSoup(tbody_html, 'html.parser')

        # Find the table rows
        rows = soup.find_all('tr')

        # Initialize variables to store the values
        EUR_value = None
        USD_value = None

        # Extract the values 7.4509 and 6.7793
        for row in rows:
            cells = row.find_all(['th', 'td'])
            if len(cells) > 1:
                cell_content = cells[1].get_text(strip=True)
                if cell_content == 'EUR' and len(cells) > 2:
                    EUR_value = cells[2].get_text(strip=True)
                elif cell_content == 'USD' and len(cells) > 2:
                    USD_value = cells[2].get_text(strip=True)

        EUR = float(EUR_value)
        USD = float(USD_value)

        if int(df.loc["US dollar", "Amount"]) > 50 and int(df.loc["Euro", "Amount"]) < 0:
            amount = min((int(df.loc["US dollar", "Amount"]) + kurtageUSD - 20) * (USD / EUR),
                         int(df.loc["Euro", "Amount"] * EUR / USD))
            df1.loc["US dollar", "Amount"] = int(df.loc["US dollar", "Amount"]) - (amount * EUR / USD)
            df1.loc["Euro", "Amount"] = int(df.loc["Euro", "Amount"]) + (amount)

            time.sleep(1)

            # Find the dropdown element by its ID
            dropdown_element = driver.find_element(By.ID, 'vaxla')
            dropdown = Select(dropdown_element)
            dropdown.select_by_value('USD')

            dropdown_element_to = driver.find_element(By.ID, 'vaxlaTill')
            dropdown = Select(dropdown_element_to)
            dropdown.select_by_value('EUR')

            time.sleep(1)

            veksel_input = driver.find_element(By.ID, 'price2')
            veksel_input.send_keys(str(amount))

            time.sleep(1)

            Veksl = driver.find_element(By.XPATH, '//*[@id="show_vaxla"]/form/div[2]/span/a')
            Veksl.click()

            time.sleep(1)



        # Hvis USD er positiv og EUR er positiv, så veksl USD til DKK
        else:
            if int(df.loc["US dollar", "Amount"]) > 50:
                amount = int(df.loc["US dollar", "Amount"] - 20 - kurtageUSD)
                df1.loc["Danish kroner", "Amount"] = int(df.loc["Danish kroner", "Amount"]) + (amount * USD)
                df1.loc["US dollar", "Amount"] = int(df.loc["US dollar", "Amount"]) - amount
                time.sleep(1)

                # Find the dropdown element by its ID
                dropdown_element = driver.find_element(By.ID, 'vaxla')
                dropdown = Select(dropdown_element)
                dropdown.select_by_value('USD')

                dropdown_element_to = driver.find_element(By.ID, 'vaxlaTill')
                dropdown = Select(dropdown_element_to)
                dropdown.select_by_value('DKK')

                time.sleep(1)

                veksel_input = driver.find_element(By.ID, 'price1')
                veksel_input.send_keys(str(amount))

                time.sleep(1)

                Veksl = driver.find_element(By.XPATH, '//*[@id="show_vaxla"]/form/div[1]/span/a')
                Veksl.click()

                time.sleep(1)
            else:
                pass

        if int(df1.loc["Euro", "Amount"] - kurtageEUR) > 50 and int(df1.loc["US dollar", "Amount"] - kurtageUSD) < 0:
            EURamount = min(((int(df.loc["US dollar", "Amount"]) + kurtageUSD - 20) * (-EUR / USD)),
                            int(df.loc["US dollar", "Amount"] * USD / EUR))
            df2.loc["Euro", "Amount"] = int(df1.loc["Euro", "Amount"]) - EURamount
            df2.loc["US dollar", "Amount"] = int(df1.loc["US dollar", "Amount"]) + (EURamount * USD / EUR)

            time.sleep(1)

            # Find the dropdown element by its ID
            dropdown_element = driver.find_element(By.ID, 'vaxla')
            dropdown = Select(dropdown_element)
            dropdown.select_by_value('EUR')

            dropdown_element_to = driver.find_element(By.ID, 'vaxlaTill')
            dropdown = Select(dropdown_element_to)
            dropdown.select_by_value('USD')

            time.sleep(1)

            veksel_input = driver.find_element(By.ID, 'price2')
            veksel_input.send_keys(str(amount))

            time.sleep(1)

            Veksl = driver.find_element(By.XPATH, '//*[@id="show_vaxla"]/form/div[2]/span/a')
            Veksl.click()

            time.sleep(1)

            # Hvis USD er positiv og EUR er positiv, så veksl USD til DKK
        else:
            if int(df1.loc["Euro", "Amount"] - kurtageEUR) > 50:
                EURamount = int(df.loc["Euro", "Amount"]) - 20 - kurtageEUR
                df2.loc["Danish kroner", "Amount"] = int(df1.loc["Danish kroner", "Amount"]) + (EURamount * EUR)
                df2.loc["Euro", "Amount"] = int(df1.loc["Euro", "Amount"]) - EURamount

                time.sleep(1)

                # Find the dropdown element by its ID
                dropdown_element = driver.find_element(By.ID, 'vaxla')
                dropdown = Select(dropdown_element)
                dropdown.select_by_value('EUR')

                dropdown_element_to = driver.find_element(By.ID, 'vaxlaTill')
                dropdown = Select(dropdown_element_to)
                dropdown.select_by_value('DKK')

                time.sleep(1)

                veksel_input = driver.find_element(By.ID, 'price1')
                veksel_input.send_keys(str(EURamount))

                time.sleep(1)

                Veksl = driver.find_element(By.XPATH, '//*[@id="show_vaxla"]/form/div[1]/span/a')
                Veksl.click()

                time.sleep(1)



            else:
                pass

        if int(df2.loc["Euro", "Amount"] - kurtageEUR) < 0 and int(
                df2.loc["Danish kroner", "Amount"] - kurtageDKK) > 100:
            DKKamount = min(((-int(df2.loc["Euro", "Amount"]) - int(20)) * EUR + kurtageEUR),
                            int(df2.loc["Danish kroner", "Amount"] - 20))
            df3.loc["Danish kroner", "Amount"] = int(df2.loc["Danish kroner", "Amount"]) - DKKamount
            df3.loc["Euro", "Amount"] = int(df2.loc["Euro", "Amount"]) + (DKKamount / EUR)

            time.sleep(1)

            # Find the dropdown element by its ID
            dropdown_element = driver.find_element(By.ID, 'vaxla')
            dropdown = Select(dropdown_element)
            dropdown.select_by_value('DKK')

            dropdown_element_to = driver.find_element(By.ID, 'vaxlaTill')
            dropdown = Select(dropdown_element_to)
            dropdown.select_by_value('EUR')

            time.sleep(1)

            veksel_input = driver.find_element(By.ID, 'price1')
            veksel_input.send_keys(str(DKKamount))

            time.sleep(1)

            Veksl = driver.find_element(By.XPATH, '//*[@id="show_vaxla"]/form/div[1]/span/a')
            Veksl.click()

            time.sleep(1)



        else:
            pass

        if int(df2.loc["US dollar", "Amount"] - kurtageUSD) < 0 and int(
                df2.loc["Danish kroner", "Amount"] - kurtageDKK) > 100:
            DKKUSDamount = min(((-int(df2.loc["US dollar", "Amount"]) + int(20)) * USD + kurtageUSD),
                               int(df2.loc["Danish kroner", "Amount"]) - 20)
            df3.loc["Danish kroner", "Amount"] = int(df2.loc["Danish kroner", "Amount"]) - DKKamount
            df3.loc["US dollar", "Amount"] = int(df2.loc["US dollar", "Amount"]) + (DKKamount / USD)

            time.sleep(1)

            # Find the dropdown element by its ID
            dropdown_element = driver.find_element(By.ID, 'vaxla')
            dropdown = Select(dropdown_element)
            dropdown.select_by_value('DKK')

            dropdown_element_to = driver.find_element(By.ID, 'vaxlaTill')
            dropdown = Select(dropdown_element_to)
            dropdown.select_by_value('USD')

            time.sleep(1)

            veksel_input = driver.find_element(By.ID, 'price1')
            veksel_input.send_keys(str(DKKUSDamount))

            time.sleep(1)

            Veksl = driver.find_element(By.XPATH, '//*[@id="show_vaxla"]/form/div[1]/span/a')
            Veksl.click()

            time.sleep(1)



        else:
            pass

driver.quit()