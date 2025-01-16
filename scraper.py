import csv
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import re
from openpyxl import load_workbook

def scrape_and_find_prices_for_multiple_codes(product_codes, output_csv_file, output_xlsx_file):
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    
    data = []

    for product_code in product_codes:
        print(f"\nSearching for product code: {product_code.strip()}")
        
        search_url = f"https://jiji.com.gh/search?query={product_code.strip()}"
        driver.get(search_url)
        
        time.sleep(5)
        
        last_height = driver.execute_script("return document.body.scrollHeight")
        
        while True:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height
        
        try:
            result_text = driver.find_element(By.XPATH, "//span[contains(text(),'results for')]").text
            print(f"Result Text: {result_text}")
            
            result_count = re.findall(r'\d+', result_text)
            
            if result_count:
                # Convert to integer to avoid saving as text
                number_of_items = int(result_count[0])
                print(f"\nNumber of items with code {product_code.strip()}: {number_of_items}")
            else:
                print("Couldn't extract the number of results.")
                number_of_items = 0  # Set to 0 if extraction fails
            
            prices = []
            price_elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'GH₵')]")
            
            for price_element in price_elements:
                price_text = price_element.text.strip()
                match = re.search(r'GH₵\s?([\d,]+)', price_text)
                
                if match:
                    price_value = match.group(1).replace(',', '')
                    price_value = float(price_value)
                    prices.append(price_value)
            
            if prices:
                highest_price = max(prices)
                lowest_price = min(prices)
                print(f"\nThe highest price found for {product_code.strip()} is: GH₵{highest_price}")
                print(f"The lowest price found for {product_code.strip()} is: GH₵{lowest_price}")
                
                # Save data in dictionary format with formatted prices
                data.append({
                    "Product Code": product_code.strip(),
                    "Number of Items": number_of_items,  # Ensure this is stored as an integer
                    "Highest Price (GH₵)": f"GH₵ {highest_price:.2f}",  # Add currency symbol and format the value
                    "Lowest Price (GH₵)": f"GH₵ {lowest_price:.2f}"  # Add currency symbol and format the value
                })
            else:
                print("No prices found.")
        
        except Exception as e:
            print(f"Error occurred: {e}")
    
    driver.quit()

    # Save to CSV with the currency symbol added
    with open(output_csv_file, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ["Product Code", "Number of Items", "Highest Price (GH₵)", "Lowest Price (GH₵)"]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(data)
    
    # Save to Excel with the currency symbol added
    df = pd.DataFrame(data)
    df.to_excel(output_xlsx_file, index=False)

    # Open the saved Excel file with openpyxl to adjust column widths
    wb = load_workbook(output_xlsx_file)
    sheet = wb.active

    # Adjust column widths based on the longest value in each column
    for col in sheet.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column name
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        sheet.column_dimensions[column].width = adjusted_width

    # Save the changes to the Excel file
    wb.save(output_xlsx_file)

# Input the product codes separated by commas (e.g., S643W, C211WDG, P1234)
product_codes_input = input("Enter product codes separated by commas: ").strip()
product_codes = [code.strip() for code in product_codes_input.split(',')]

# Specify the output file paths
output_csv_file = "scraped_data/product_prices.csv"
output_xlsx_file = "scraped_data/product_prices.xlsx"

# Run the scraper and save results to CSV and Excel
scrape_and_find_prices_for_multiple_codes(product_codes, output_csv_file, output_xlsx_file)
