from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import re
import csv
import pandas as pd
from openpyxl import load_workbook
from collections import Counter


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
                number_of_items = int(result_count[0])
                print(f"\nNumber of items with code {product_code.strip()}: {number_of_items}")
            else:
                print("Couldn't extract the number of results.")
                number_of_items = 0
            
            prices = []
            locations = []
            price_elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'GH₵')]")
            location_elements = driver.find_elements(By.XPATH, "//div[contains(@class,'location')]")
            
            for price_element in price_elements:
                price_text = price_element.text.strip()
                match = re.search(r'GH₵\s?([\d,]+)', price_text)
                if match:
                    price_value = match.group(1).replace(',', '')
                    price_value = float(price_value)
                    prices.append(price_value)
            
            for location_element in location_elements:
                location_text = location_element.text.strip()
                if ',' in location_text:
                    location_part = location_text.split(',')[1].strip()
                    locations.append(location_part)
            
            if prices:
                highest_price = max(prices)
                lowest_price = min(prices)
                average_price = sum(prices) / len(prices)
                
                price_counter = Counter(prices)
                most_common_price, _ = price_counter.most_common(1)[0]
                
                if locations:
                    location_counter = Counter(locations)
                    top_competitor_location, _ = location_counter.most_common(1)[0]
                else:
                    top_competitor_location = "Not Available"
                
                print(f"\nThe highest price found for {product_code.strip()} is: GH₵{highest_price}")
                print(f"The lowest price found for {product_code.strip()} is: GH₵{lowest_price}")
                print(f"The average price found for {product_code.strip()} is: GH₵{average_price:.2f}")
                print(f"The most common price found for {product_code.strip()} is: GH₵{most_common_price}")
                print(f"The top competitor location for {product_code.strip()} is: {top_competitor_location}")
                
                data.append({
                    "Product Code": product_code.strip(),
                    "Number of Items": number_of_items,
                    "Highest Price": f"GH₵ {highest_price:.2f}",
                    "Lowest Price": f"GH₵ {lowest_price:.2f}",
                    "Average Price": f"GH₵ {average_price:.2f}",
                    "Most Common Price": f"GH₵ {most_common_price:.2f}",
                    "Top Competitor Location": top_competitor_location
                })
            else:
                print("No prices or locations found.")
        
        except Exception as e:
            print(f"Error occurred: {e}")
    
    driver.quit()

    # Save to CSV with all the fields
    with open(output_csv_file, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ["Product Code", "Number of Items", "Highest Price", "Lowest Price", "Average Price", "Most Common Price", "Top Competitor Location"]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(data)
    
    # Save to Excel with all the fields
    df = pd.DataFrame(data)
    df.to_excel(output_xlsx_file, index=False)

    # Adjust column widths
    wb = load_workbook(output_xlsx_file)
    sheet = wb.active
    for col in sheet.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        sheet.column_dimensions[column].width = adjusted_width

    wb.save(output_xlsx_file)

# Input the product codes separated by commas (e.g., S643W, C211WDG, P1234)
product_codes_input = input("Enter product codes separated by commas: ").strip()
product_codes = [code.strip() for code in product_codes_input.split(',')]

# Specify the output file paths
output_csv_file = "scraped_data/research_data.csv"
output_xlsx_file = "scraped_data/research_data.xlsx"

# Run the scraper and save results to CSV and Excel
scrape_and_find_prices_for_multiple_codes(product_codes, output_csv_file, output_xlsx_file)
