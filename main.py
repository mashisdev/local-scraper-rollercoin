import pandas as pd
from bs4 import BeautifulSoup
from dotenv import load_dotenv
import os

load_dotenv()
SHEET_PATH = "sheets/rollercoin-scraper-sheet.xlsx"
WORKSHEET_NAME = "PythonSheet"

def convert_power(power_str):
    """Converts power from Gh/s, Th/s, or Ph/s to a numerical value."""
    print(f"Converting power: {power_str}")
    power_str = power_str.replace(",", "")
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
        return 0

def extract_html_data(html_content):
    print("Starting HTML extraction...")
    soup = BeautifulSoup(html_content, "html.parser")
    item_cards = soup.find_all("a", class_="marketplace-buy-item-card")
    print(f"Found {len(item_cards)} item cards")
    
    results = []

    for i, card in enumerate(item_cards):
        try:
            print(f"Processing card {i+1}")
            
            # Precio
            price_element = card.find("p", class_="item-price")
            if not price_element:
                print(f"  Price element not found in card {i+1}")
                continue
            price_str = price_element.text.strip()
            item_price = price_str.replace(" RLT", "")
            print(f"  Price: {item_price}")

            # Power
            power_element = card.find("span", class_="item-addition-power")
            if not power_element:
                print(f"  Power element not found in card {i+1}")
                continue
            power_str = power_element.text.strip()
            item_addition_power = convert_power(power_str)
            print(f"  Power: {item_addition_power}")

            # Bonus
            bonus_element = card.find("span", class_="item-addition-bonus")
            if not bonus_element:
                print(f"  Bonus element not found in card {i+1}")
                continue
            item_addition_bonus = bonus_element.text.strip()
            print(f"  Bonus: {item_addition_bonus}")

            # Título y rareza
            item_title_str = card.find("p", class_="item-title")
            if not item_title_str:
                print(f"  Title element not found in card {i+1}")
                continue
                
            rarity_span = item_title_str.find("span")
            rarity = rarity_span.text.strip() if rarity_span else ""
            item_title = item_title_str.text.replace(rarity, "").strip()
            print(f"  Title: {item_title}")
            print(f"  Rarity: {rarity}")

            results.append({
                "item_title": item_title,
                "rarity": rarity,
                "item_addition_power": item_addition_power,
                "item_addition_bonus": item_addition_bonus,
                "item_price": item_price,
            })
            print(f"  ✓ Successfully processed card {i+1}")

        except Exception as e:
            print(f"  ✗ Error processing card {i+1}: {e}")
            continue

    print(f"Extraction completed. Total items: {len(results)}")
    return results

def safe_convert_bonus(bonus_value):
    """Convierte de forma segura el valor de bonus a numérico"""
    try:
        if pd.isna(bonus_value):
            return bonus_value
            
        if isinstance(bonus_value, (int, float)):
            # Si ya es numérico, verificar si necesita división por 100
            if bonus_value > 1:  # Si es porcentaje como 7.2 en lugar de 0.072
                return bonus_value / 100
            return bonus_value
        elif isinstance(bonus_value, str):
            # Si es texto, procesar normalmente
            cleaned = bonus_value.replace('%', '').replace(',', '.')
            numeric_value = pd.to_numeric(cleaned, errors='coerce')
            return numeric_value / 100 if not pd.isna(numeric_value) else numeric_value
        else:
            return bonus_value
    except Exception as e:
        print(f"Error converting bonus value '{bonus_value}': {e}")
        return bonus_value

def update_excel_sheet(data):
    print(f"Updating Excel sheet with {len(data)} items")
    
    # Verificar directorio
    os.makedirs(os.path.dirname(SHEET_PATH), exist_ok=True)
    print(f"Directory ensured: {os.path.dirname(SHEET_PATH)}")
    
    # Verificar si el archivo existe
    file_exists = os.path.exists(SHEET_PATH)
    print(f"Excel file exists: {file_exists}")
    
    if file_exists:
        try:
            excel_file = pd.ExcelFile(SHEET_PATH)
            sheet_exists = WORKSHEET_NAME in excel_file.sheet_names
            print(f"Worksheet '{WORKSHEET_NAME}' exists: {sheet_exists}")
            
            if sheet_exists:
                df = pd.read_excel(SHEET_PATH, sheet_name=WORKSHEET_NAME)
                print(f"Loaded existing sheet with {len(df)} rows")
                print(f"Column dtypes: {df.dtypes}")
            else:
                df = pd.DataFrame(columns=["Miner", "Rarity", "Power", "% Bonus", "Price"])
                print("Created new DataFrame (sheet didn't exist)")
                
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            df = pd.DataFrame(columns=["Miner", "Rarity", "Power", "% Bonus", "Price"])
    else:
        df = pd.DataFrame(columns=["Miner", "Rarity", "Power", "% Bonus", "Price"])
        print("Created new DataFrame (file didn't exist)")
    
    print(f"DataFrame shape before processing: {df.shape}")
    
    # Procesar cada item
    new_items_count = 0
    updated_items_count = 0
    
    for item in data:
        item_title = item["item_title"].strip()
        item_power = item["item_addition_power"]
        
        # Buscar si ya existe el item
        mask = (df["Miner"].astype(str).str.strip() == item_title) & (df["Power"] == item_power)
        
        if mask.any():
            # Convertir el precio a float antes de asignar
            try:
                price_float = float(item["item_price"])
                df.loc[mask, "Price"] = price_float
                updated_items_count += 1
                print(f"Price updated for {item_title} ({item_power}): {item['item_price']}")
            except ValueError as e:
                print(f"Error converting price '{item['item_price']}' to float: {e}")
                continue
        else:
            new_row = {
                "Miner": item["item_title"],
                "Rarity": item["rarity"],
                "Power": item["item_addition_power"],
                "% Bonus": item["item_addition_bonus"],
                "Price": item["item_price"]
            }
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            new_items_count += 1
            print(f"New entry added: {item_title} ({item_power})")
    
    print(f"Processing complete: {new_items_count} new, {updated_items_count} updated")
    print(f"DataFrame shape after processing: {df.shape}")
    
    # Convertir tipos de datos DE FORMA SEGURA
    try:
        print("Converting data types...")
        
        # Power - ya debería ser numérico
        df["Power"] = pd.to_numeric(df["Power"], errors='coerce')
        print("  Power converted")
        
        # % Bonus - conversión segura
        df["% Bonus"] = df["% Bonus"].apply(safe_convert_bonus)
        print("  % Bonus converted")
        
        # Price - convertir a numérico
        df["Price"] = pd.to_numeric(df["Price"], errors='coerce')
        print("  Price converted")
        
        print("Data types converted successfully")
        print(f"Final dtypes: {df.dtypes}")
        
    except Exception as e:
        print(f"Error converting data types: {e}")
    
    # Guardar el archivo Excel
    try:
        if file_exists:
            print("Saving with mode 'a' (append/replace)...")
            # Forzar cierre de cualquier conexión previa
            import gc
            gc.collect()
            
            with pd.ExcelWriter(SHEET_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=WORKSHEET_NAME, index=False)
                print("  File saved successfully")
        else:
            print("Saving with mode 'w' (new file)...")
            with pd.ExcelWriter(SHEET_PATH, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=WORKSHEET_NAME, index=False)
                print("  New file created successfully")
            
        print(f"✓ Successfully saved to {SHEET_PATH}")
        print(f"✓ Worksheet: {WORKSHEET_NAME}")
        print(f"✓ Total rows: {len(df)}")
        
        # Verificar que el archivo se actualizó
        if os.path.exists(SHEET_PATH):
            file_size = os.path.getsize(SHEET_PATH)
            print(f"✓ File size: {file_size} bytes")
            # Leer de nuevo para verificar
            verify_df = pd.read_excel(SHEET_PATH, sheet_name=WORKSHEET_NAME)
            print(f"✓ Verification: {len(verify_df)} rows in saved file")
        
    except Exception as e:
        print(f"✗ Error saving Excel file: {e}")
        # Backup como CSV
        try:
            backup_path = SHEET_PATH.replace('.xlsx', '_backup.csv')
            df.to_csv(backup_path, index=False)
            print(f"Backup saved to {backup_path}")
        except Exception as e2:
            print(f"✗ Critical error: {e2}")

if __name__ == "__main__":
    print("Please enter the HTML content (from marketplace-buy-items-list):")
    html_content = input()

    try:
        print("Starting script execution...")
        data = extract_html_data(html_content)
        if data:
            update_excel_sheet(data)
            print("\n✓ Operation completed successfully!")
        else:
            print("\n✗ No data extracted from HTML content")
    except Exception as e:
        print(f"\n✗ General error: {e}")