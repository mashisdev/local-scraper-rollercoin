import pandas as pd
from bs4 import BeautifulSoup
from dotenv import load_dotenv
import os

load_dotenv()
SHEET_PATH = "sheets/rollercoin-sheet.xlsx"  # Ruta local del archivo Excel
WORKSHEET_NAME = "PythonSheet"  # Nombre de la hoja específica

def convert_power(power_str):
    """Converts power from Gh/s, Th/s, or Ph/s to a numerical value."""

    power_str = power_str.replace(",", "")  # Remove commas for large numbers
    if "Gh/s" in power_str:
        value = float(power_str.replace(" Gh/s", ""))
        return value
    elif "Th/s" in power_str:
        value = float(power_str.replace(" Th/s", "")) * 1000
        return value
    elif "Ph/s" in power_str:
        value = float(power_str.replace(" Ph/s", "")) * 1000000
        return value
    else:
        return 0  # Return 0 if no unit is found


def extract_html_data(html_content):
    # Extracts specific data from all marketplace-buy-item-card elements.

    soup = BeautifulSoup(html_content, "html.parser")
    item_cards = soup.find_all("a", class_="marketplace-buy-item-card")
    results = []

    for card in item_cards:
        try:
            price_str = card.find("p", class_="item-price").text.strip()
            item_price = price_str.replace(" RLT", "")  # Remove RLT from price
            power_str = card.find("span", class_="item-addition-power").text.strip()
            item_addition_power = convert_power(power_str)
            item_addition_bonus = card.find(
                "span", class_="item-addition-bonus"
            ).text.strip()
            item_title_str = card.find("p", class_="item-title")
            rarity = (
                item_title_str.find("span").text.strip()
                if item_title_str.find("span")
                else ""
            )
            item_title = item_title_str.text.replace(rarity, "").strip()

            results.append(
                {
                    "item_title": item_title,
                    "rarity": rarity,
                    "item_addition_power": item_addition_power,
                    "item_addition_bonus": item_addition_bonus,
                    "item_price": item_price,
                }
            )
        except AttributeError:
            print("Elements not found within an item-card.")
            continue
        except Exception as e:
            print(f"Error within an item-card: {e}")
            continue

    return results


def update_excel_sheet(data):
    # Crear el directorio si no existe
    os.makedirs(os.path.dirname(SHEET_PATH), exist_ok=True)
    
    # Verificar si el archivo existe
    if os.path.exists(SHEET_PATH):
        try:
            # Leer el archivo Excel completo
            excel_file = pd.ExcelFile(SHEET_PATH)
            
            # Verificar si la hoja existe
            if WORKSHEET_NAME in excel_file.sheet_names:
                df = pd.read_excel(SHEET_PATH, sheet_name=WORKSHEET_NAME)
            else:
                # Crear nuevo DataFrame si la hoja no existe
                df = pd.DataFrame(columns=["Miner", "Rarity", "Power", "% Bonus", "Price"])
                print(f"Worksheet '{WORKSHEET_NAME}' not found. Creating new one.")
                
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            # Crear nuevo DataFrame si hay error al leer
            df = pd.DataFrame(columns=["Miner", "Rarity", "Power", "% Bonus", "Price"])
    else:
        # Crear nuevo DataFrame si el archivo no existe
        df = pd.DataFrame(columns=["Miner", "Rarity", "Power", "% Bonus", "Price"])
        print(f"Excel file '{SHEET_PATH}' not found. Creating new one.")
    
    # Procesar cada item
    for item in data:
        item_title = item["item_title"].strip()
        item_power = item["item_addition_power"]
        
        # Buscar si ya existe el item
        mask = (df["Miner"].astype(str).str.strip() == item_title) & (df["Power"] == item_power)
        
        if mask.any():
            # Actualizar precio existente
            df.loc[mask, "Price"] = item["item_price"]
            print(f"Price updated for {item_title} ({item_power}): {item['item_price']}")
        else:
            # Agregar nueva entrada
            new_row = {
                "Miner": item["item_title"],
                "Rarity": item["rarity"],
                "Power": item["item_addition_power"],
                "% Bonus": item["item_addition_bonus"],
                "Price": item["item_price"]
            }
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            print(f"New entry added: {item_title} ({item_power})")
    
    # CONVERSIÓN DE TIPOS DE DATOS - PARTE CLAVE
    try:
        # Convertir Power a numérico
        df["Power"] = pd.to_numeric(df["Power"], errors='coerce')
        
        # Convertir % Bonus a porcentaje numérico
        df["% Bonus"] = df["% Bonus"].str.replace('%', '').str.replace(',', '.')
        df["% Bonus"] = pd.to_numeric(df["% Bonus"], errors='coerce') / 100
        
        # Convertir Price a numérico
        df["Price"] = pd.to_numeric(df["Price"], errors='coerce')
        
    except Exception as e:
        print(f"Warning: Error converting data types: {e}")
    
    # Guardar el archivo Excel
    try:
        # Si el archivo ya existe, necesitamos mantener las otras hojas
        if os.path.exists(SHEET_PATH):
            with pd.ExcelWriter(SHEET_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=WORKSHEET_NAME, index=False)
                
                # Obtener la hoja de trabajo para aplicar formatos
                worksheet = writer.sheets[WORKSHEET_NAME]
                
                # Aplicar formato de porcentaje a la columna % Bonus
                for row in range(2, len(df) + 2):  # +2 porque Excel empieza en 1 y tiene headers
                    cell = worksheet.cell(row=row, column=4)  # Columna D (% Bonus)
                    cell.number_format = '0.00%'
                    
                # Aplicar formato numérico a la columna Price
                for row in range(2, len(df) + 2):
                    cell = worksheet.cell(row=row, column=5)  # Columna E (Price)
                    cell.number_format = '0.00'
                    
        else:
            # Si el archivo no existe, crear uno nuevo
            with pd.ExcelWriter(SHEET_PATH, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=WORKSHEET_NAME, index=False)
                
                # Aplicar formatos
                worksheet = writer.sheets[WORKSHEET_NAME]
                
                # Formato de porcentaje para % Bonus
                for row in range(2, len(df) + 2):
                    cell = worksheet.cell(row=row, column=4)
                    cell.number_format = '0.00%'
                    
                # Formato numérico para Price
                for row in range(2, len(df) + 2):
                    cell = worksheet.cell(row=row, column=5)
                    cell.number_format = '0.00'
            
        print(f"Data saved to sheet '{WORKSHEET_NAME}' in {SHEET_PATH}")
        
    except Exception as e:
        print(f"Error saving Excel file: {e}")
        # Intentar guardar como nuevo archivo
        try:
            df.to_excel(SHEET_PATH, sheet_name=WORKSHEET_NAME, index=False)
            print(f"Data saved to new file: {SHEET_PATH}")
        except Exception as e2:
            print(f"Critical error: {e2}")
            # Respaldo como CSV
            backup_path = SHEET_PATH.replace('.xlsx', '_backup.csv')
            df.to_csv(backup_path, index=False)
            print(f"Backup saved to {backup_path}")


if __name__ == "__main__":
    print("Please enter the HTML content (from marketplace-buy-items-list):")
    html_content = input()

    try:
        data = extract_html_data(html_content)
        update_excel_sheet(data)
        print("\nOperation completed.")
    except Exception as e:
        print(f"General error: {e}")