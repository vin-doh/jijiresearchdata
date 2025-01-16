from scraper import scrape_and_find_prices_for_multiple_codes
from utils import save_to_csv

def main():
    # Input the product codes separated by commas (e.g., S643W, C211WDG, P1234)
    product_codes_input = input("Enter product codes separated by commas: ").strip()
    
    # Split the input into a list of product codes
    product_codes = [code.strip() for code in product_codes_input.split(',')]

    # Define the CSV output file
    output_csv_file = 'scraped_data/product_prices.csv'

    # Run the scraper for each product code and save the results to CSV
    scrape_and_find_prices_for_multiple_codes(product_codes, output_csv_file)

if __name__ == "__main__":
    main()
