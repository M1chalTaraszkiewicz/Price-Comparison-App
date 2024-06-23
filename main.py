import os
import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from lxml import html
import time
import datetime


def get_price(driver, url, price_xpath):
    try:
        driver.get(url)
        time.sleep(3)
        page_source = driver.page_source
        tree = html.fromstring(page_source)
        prices = tree.xpath(price_xpath)
        # Rozdzielanie wartości od waluty bądź innych dodatkowych informacji
        if prices:
            price_text = prices[0].strip()
            if u'\xa0' in price_text:
                price_text = price_text.split(u'\xa0', 1)[0]
            elif ' ' in price_text:
                price_text = price_text.split(' ', 1)[0]
            return price_text.replace(",", ".")
        else:
            return "Not Found"
    except WebDriverException:
        return "-"


# Tworzenie folderu na wyniki jeśli nie istnieje
if not os.path.exists("Wyniki"):
    os.makedirs("Wyniki")


# Wybór pliku/ów do zczytania linków
def choose_file():
    files = [f for f in os.listdir() if f.endswith('.xlsx')]
    if not files:
        print("Brak plików .xlsx w bieżącym katalogu.")
        return None

    print("Wybierz plik do odczytu:")
    print("0. Użyj wszystkich plików")
    for i, file in enumerate(files):
        print(f"{i+1}. {file}")

    while True:
        try:
            choice = int(input("Podaj numer pliku: "))
            if choice == 0:
                return files
            elif 1 <= choice <= len(files):
                return [files[choice - 1]]
            else:
                print("Niepoprawny wybór, spróbuj ponownie.")
        except ValueError:
            print("Proszę podać poprawny numer.")


try:

    # XPathy do cen dla poszczególnych sklepów w formacie 'nazwa sklepu podana w pliku .xlsx': 'xpath'
    price_xpaths = {
        'MyShop': '//div[@class="current-price"]//span[contains(text(), "Cena:")]/following-sibling::text()',
        'Shop1': '//div[@class="product-prices"]//span[contains(text(), "zł") and contains(text(), "m2")]/text()',
        'Shop2': '//div[@id="cena"]/descendant-or-self::text()[not(parent::s)][last()]',
        'Shop3': '//div[@id="cena"]/descendant-or-self::text()[not(parent::s)][last()]',
        'Shop4': '//span[@id="our_price_display" and @class="price"]/text()',
        'Shop5': '//p[@class="our_price_display"]/span[@class="price"]/em/text()',
        'Shop6': '//p[@class="price"]//ins/span[@class="woocommerce-Price-amount amount"]/bdi/text() | //p[@class="price"]/span[@class="wc-measurement-price-calculator-price"]/span[@class="woocommerce-Price-amount amount"]/bdi/text()',
        'Shop7': '//div[@class="price"]//div[contains(@class, "belt price_brutto clearfix") and div[contains(text(), "Cena brutto:")]]//span[@class="price common_color"]/text()',
        'Shop8': '//div[@class="product_price_wrap regular_price"]//span[@class="price"]/strong/text()',
        'Shop9': '//span[contains(@id, "sec_discounted_price_") and @class="ty-price-num"]/text()',
    }

    # Wybór pliku
    selected_files = choose_file()
    if selected_files is None:
        exit()

    for file in selected_files:
        # Konfiguracja WebDrivera
        driver = webdriver.Edge()
        df = pd.read_excel(file)

        # Tworzenie nowego DataFrame i kopiowanie kolumny z nazwami produktów
        df2 = df[['Product Name']].copy()

        # Pobieranie ceny dla każdego produktu ze wszystkich sklepów
        for index, row in df.iterrows():
            promo_code = row['Promo Code'] if 'Promo Code' in row and not pd.isna(row['Promo Code']) else 0
            for shop, xpath in price_xpaths.items():
                price = get_price(driver, row[shop], xpath)
                if price != "Not Found" and price != "-":
                    price = float(price)
                    if shop == 'MyShop':
                        discounted_price = price * (1 - promo_code / 100)
                        df2.at[index, 'MyShop Price'] = f"{price:.2f}"
                        df2.at[index, 'After Promo Code'] = f"{discounted_price:.2f}"
                    else:
                        df2.at[index, shop + ' Price'] = f"{price:.2f}"
                else:
                    df2.at[index, shop + ' Price'] = price
            print(f"{row['Product Name']} Completed!")

        driver.quit()

        # Przeniesienie kolumny 'After Promo Code' w odpowiednie miejsce
        columns = list(df2.columns)
        columns.insert(columns.index('MyShop Price') + 1, columns.pop(columns.index('After Promo Code')))
        df2 = df2[columns]

        # Dodanie kolumny, która będzie określać czy wiersz ma być podświetlony
        df2['Highlight'] = False

        # Sprawdzenie, które wiersze powinny być podświetlone
        for index, row in df2.iterrows():
            after_promo_code = float(row['After Promo Code']) if row['After Promo Code'] != "Not Found" and row['After Promo Code'] != "-" else float('inf')
            for shop in price_xpaths.keys():
                if shop + ' Price' in row and row[shop + ' Price'] != "Not Found" and row[shop + ' Price'] != "-":
                    price = float(row[shop + ' Price'])
                    if price < after_promo_code:
                        df2.at[index, 'Highlight'] = True
                        break

        # Zapisanie wyników do nowego pliku Excel z kolorowaniem
        # Nazwa pliku to Ceny_'nazwa odczytanego pliku'_'aktualna data'
        current_date = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M")
        output_file = os.path.join("Wyniki", f"Ceny_{os.path.splitext(file)[0]}_{current_date}.xlsx")
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            df2.to_excel(writer, index=False, sheet_name='Prices')

            workbook = writer.book
            worksheet = writer.sheets['Prices']

            # Pobieranie liczby wierszy i kolumn
            (max_row, max_col) = df2.shape
            max_row += 1  # Dodaj nagłówek

            # Tworzenie formatów do podświetlania
            format_red = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
            format_green = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
            format_yellow = workbook.add_format({'bg_color': '#FFFF00', 'font_color': '#000000'})

            # Dodanie warunkowego formatowania dla kolumny Highlight
            worksheet.conditional_format(1, max_col-1, max_row-1, max_col-1, {
                'type': 'cell',
                'criteria': '==',
                'value': True,
                'format': format_red
            })
            worksheet.conditional_format(1, max_col-1, max_row-1, max_col-1, {
                'type': 'cell',
                'criteria': '==',
                'value': False,
                'format': format_green
            })

            # Dodanie warunkowego formatowania dla komórek z cenami niższymi niż cena naszego sklepu internetowego
            for index, row in df2.iterrows():
                after_promo_code = float(row['After Promo Code']) if row['After Promo Code'] != "Not Found" and row['After Promo Code'] != "-" else float('inf')
                for col_num, (col_name, cell_value) in enumerate(row.items()):
                    if col_name.endswith('Price') and col_name != 'MyShop Price':
                        try:
                            cell_value = float(cell_value)
                            if cell_value < after_promo_code:
                                worksheet.write(index + 1, col_num, cell_value, format_red)
                            elif cell_value == after_promo_code:
                                worksheet.write(index + 1, col_num, cell_value, format_yellow)
                            else:
                                worksheet.write(index + 1, col_num, cell_value)
                        except ValueError:
                            worksheet.write(index + 1, col_num, cell_value)

    print("Finished")

# Obsługa błędów
except Exception as e:
    print(f"An error occurred: {e}")
    input("Press Enter to exit...")
