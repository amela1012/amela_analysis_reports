
import requests
from bs4 import BeautifulSoup
import json
import pandas as pd


def format_product_names(name):
    # Replace commas and spaces with dashes, then convert to lowercase
    name = name.replace(',', '-').replace(' ', '-').lower()

    # Replace German letters with their ASCII equivalents
    replacements = {
        'ä': 'a',
        'ö': 'oe',
        'ü': 'ue',
        'ß': 'ss',
        'Ä': 'A',
        'Ö': 'OE',
        'Ü': 'UE'
    }

    for key, value in replacements.items():
        name = name.replace(key, value)

    return name


def fetch_page_soup(url):
    response = requests.get(url)
    if response.status_code != 200:
        print(f"Failed to retrieve page. Status code: {response.status_code} for URL: {url}")
        return None
    return BeautifulSoup(response.content, 'html.parser')


def extract_data_layer(soup):
    data = []
    data_layer_inputs = soup.find_all('input', {'name': 'data-layer'})
    for input_tag in data_layer_inputs:
        data_layer_value = input_tag.get('value')
        if data_layer_value:  # Check if value is not None
            data_layer_data = json.loads(data_layer_value)
            data.append(data_layer_data)
    return data


def get_products(base_url, item, max_pages=None):
    page_num = 1
    data = []
    while True:
        url = base_url.format(page_num=page_num)
        soup = fetch_page_soup(url)
        if not soup:
            break

        page_data = extract_data_layer(soup)
        if not page_data:
            print(f"No more products found on page {page_num}. Stopping.")
            break

        data.extend(page_data)
        if max_pages and page_num >= max_pages:
            print(f"Reached max_pages limit: {max_pages}. Stopping.")
            break
        page_num += 1

    df = pd.DataFrame(data)
    print(f"Total rows: {len(df)}.")
    return df


def get_table(soup):
    data = []
    if soup:
        table = soup.find('table', {'class': 'product-detail-properties-table'})
        if table:
            rows = table.find('tbody').find_all('tr')
            for row in rows:
                label = row.find('th').text.strip()
                value = row.find('td').text.strip()
                data.append({"label": label, "value": value})
        else:
            print("No table found on the page.")
    else:
        print("Soup is None. Cannot fetch table.")
    return data


def get_product_details(url, df, max_items=10):
    items = df['item_name_normalized'].unique().tolist()[:max_items]
    data = []

    for item in items:
        url = url.format(item=item)
        soup = fetch_page_soup(url)
        if soup:  # Ensure soup is valid
            table_data = get_table(soup)
            # Flatten the table_data and append item_name_normalized
            for entry in table_data:
                flattened_entry = {
                    'item_name_normalized': item,
                    'label': entry['label'],
                    'value': entry['value']
                }
                data.append(flattened_entry)  # Append the flattened entry
        else:
            print(f"Failed to fetch data for {item}")

    # Create DataFrame from the flattened data
    df = pd.DataFrame(data)
    print(f"Total rows: {len(df)}.")
    return df


def write_to_excel(df, product):
    excel_file_name = 'koro_analysis'
    with pd.ExcelWriter(f"{excel_file_name}.xlsx", engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name=product, index=False)
    print(f"Excel tab {product} created.")


# Snacks
url = "https://www.korodrogerie.de/snacks?order=empfehlung&p={page_num}"
product = 'snacks'
df_snacks = get_products(url, product)
df_snacks['item_name_normalized'] = df_snacks['item_name'].apply(format_product_names)

# Snacks - details
url = "https://www.korodrogerie.de/{item}"
product = 'snacks_items'
df_snacks_items = get_product_details(url=url, df=df_snacks, max_items=20)

# Write the DataFrame to an Excel file
excel_file_name = 'koro_analysis'
with pd.ExcelWriter(f"{excel_file_name}.xlsx", engine="xlsxwriter") as writer:
    df_snacks.to_excel(writer, sheet_name='snacks', index=False)
    df_snacks_items.to_excel(writer, sheet_name='snacks_items', index=False)

    # Format
    worksheet = writer.sheets['snacks']
    worksheet.autofit()
    worksheet = writer.sheets['snacks_items']
    worksheet.autofit()

