import argparse
import os
import lxml.etree as ET
import openpyxl
from datetime import datetime

def parse_args():
    parser = argparse.ArgumentParser(description="Convert an edited Excel sheet back to YML XML format.")
    parser.add_argument("--input", required=True, help="Path to input Excel file (.xlsx)")
    parser.add_argument("--output", required=True, help="Path to output XML file (.xml)")
    return parser.parse_args()

def parse_other_params(params_str):
    params = []
    if not params_str:
        return params
        
    # Split by semicolon followed by space
    parts = params_str.split("; ")
    for part in parts:
        if ":" in part:
            name, val = part.split(":", 1)
            params.append((name.strip(), val.strip()))
    return params

def main():
    args = parse_args()
    
    print(f"Loading Excel file: {args.input} ...")
    if not os.path.exists(args.input):
        print(f"Error: Excel file not found at {args.input}")
        return
        
    wb = openpyxl.load_workbook(args.input, data_only=True)
    ws = wb.active
    
    # Read headers to map columns
    headers = [cell.value for cell in ws[1]]
    header_map = {name: index for index, name in enumerate(headers)}
    
    required_headers = ["Offer ID", "Назва UA", "Ціна", "ID Поточної Категорії"]
    for req in required_headers:
        if req not in header_map:
            print(f"Error: Missing required column '{req}' in Excel file.")
            return

    # Helper to get cell value by column name
    def get_val(row, col_name):
        idx = header_map.get(col_name)
        if idx is not None:
            val = row[idx].value
            return str(val).strip() if val is not None else ""
        return ""

    # Build XML tree
    yml_catalog = ET.Element("yml_catalog", date=datetime.now().strftime("%Y-%m-%d %H:%M"))
    shop = ET.SubElement(yml_catalog, "shop")
    
    ET.SubElement(shop, "name").text = "AVI KASTA"
    ET.SubElement(shop, "url").text = "https://avi.in.ua"
    
    currencies = ET.SubElement(shop, "currencies")
    ET.SubElement(currencies, "currency", id="UAH", rate="1")
    
    categories = ET.SubElement(shop, "categories")
    offers = ET.SubElement(shop, "offers")
    
    unique_categories = {}
    
    row_count = 0
    # Iterate through rows starting from row 2
    for row in list(ws.iter_rows())[1:]:
        offer_id = get_val(row, "Offer ID")
        if not offer_id:
            continue # skip empty rows
            
        name_ua = get_val(row, "Назва UA")
        desc = get_val(row, "Опис")
        price = get_val(row, "Ціна")
        price_old = get_val(row, "Стара ціна")
        qty = get_val(row, "Кількість")
        available = get_val(row, "Наявність") or "true"
        vendor = get_val(row, "Бренд")
        
        # Category Logic: Use new category if mapped, else current category
        cat_id = get_val(row, "ID НОВОЇ Категорії")
        cat_name = get_val(row, "Назва НОВОЇ Категорії")
        
        if not cat_id:
            cat_id = get_val(row, "ID Поточної Категорії")
            cat_name = get_val(row, "Назва Поточної Категорії")
            
        if cat_id:
            unique_categories[cat_id] = cat_name or f"Категорія {cat_id}"
            
        pictures_str = get_val(row, "Посилання на фото")
        color = get_val(row, "Колір")
        size = get_val(row, "Розмір")
        other_params_str = get_val(row, "Інші Характеристики")
        
        # Build offer tag
        offer_el = ET.SubElement(offers, "offer", id=offer_id, available=available)
        
        ET.SubElement(offer_el, "price").text = price
        if price_old:
            ET.SubElement(offer_el, "price_old").text = price_old
            
        if qty:
            ET.SubElement(offer_el, "stock_quantity").text = qty
            
        ET.SubElement(offer_el, "currencyId").text = "UAH"
        ET.SubElement(offer_el, "categoryId").text = cat_id
        
        # Pictures
        if pictures_str:
            for pic_url in pictures_str.split(","):
                pic_url = pic_url.strip()
                if pic_url:
                    ET.SubElement(offer_el, "picture").text = pic_url
                    
        if vendor:
            ET.SubElement(offer_el, "vendor").text = vendor
            
        ET.SubElement(offer_el, "name_ua").text = name_ua
        
        # Description using CDATA
        if desc:
            ET.SubElement(offer_el, "description_ua").text = ET.CDATA(desc)
            
        # Write specific params
        if color:
            ET.SubElement(offer_el, "param", name="Колір").text = color
        if size:
            ET.SubElement(offer_el, "param", name="Розмір").text = size
            
        # Parse and write other params
        other_params = parse_other_params(other_params_str)
        for p_name, p_val in other_params:
            ET.SubElement(offer_el, "param", name=p_name).text = p_val
            
        row_count += 1
        if row_count % 1000 == 0:
            print(f"Constructed {row_count} XML offers...")
            
    # Write Categories block
    for cid, cname in unique_categories.items():
        ET.SubElement(categories, "category", id=cid).text = cname
        
    # Save file
    print(f"Saving final XML to: {args.output} ...")
    tree = ET.ElementTree(yml_catalog)
    
    with open(args.output, "wb") as f:
        tree.write(f, xml_declaration=True, encoding="UTF-8", pretty_print=True)
        
    print(f"Done! Recreated XML with {row_count} offers and {len(unique_categories)} categories.")

if __name__ == "__main__":
    main()
