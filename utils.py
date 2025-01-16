import csv

def save_to_csv(product_code, number_of_items, highest_price, lowest_price):
    # Define the CSV output file
    output_csv_file = 'scraped_data/product_prices.csv'

    # Open the CSV file and append the result
    with open(output_csv_file, 'a', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['Product Code', 'Number of Items', 'Highest Price (GH₵)', 'Lowest Price (GH₵)']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        
        # Write the results for the current product code
        writer.writerow({
            'Product Code': product_code.strip(),
            'Number of Items': number_of_items,
            'Highest Price (GH₵)': highest_price,
            'Lowest Price (GH₵)': lowest_price
        })
