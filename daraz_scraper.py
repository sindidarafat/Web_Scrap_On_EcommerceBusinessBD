import requests
import json
import re
import time
import os
from concurrent.futures import ThreadPoolExecutor, as_completed
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Common request headers mimicking desktop browser
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'application/json, text/plain, */*',
    'Accept-Language': 'en-US,en;q=0.9',
    'X-Requested-With': 'XMLHttpRequest'
}

# 80+ Diverse Search Queries spanning all major departments on Daraz
QUERIES = [
    # Sports & Fitness
    'tennis ball', 'cricket ball', 'cricket bat', 'football', 'basketball',
    'badminton racket', 'shuttlecock', 'gym gloves', 'yoga mat', 'dumbbell',
    'skipping rope', 'swimming goggles', 'bicycle accessories', 'sports jersey',
    
    # Kitchen & Dining
    'dinner plate', 'glass cup', 'coffee mug', 'water bottle', 'frying pan',
    'cooking pot', 'kitchen knife', 'blender', 'electric kettle', 'lunch box',
    'spoon fork set', 'spice jar', 'food storage container',
    
    # Home & Decor
    'bedsheet', 'pillow', 'blanket', 'window curtain', 'floor mat',
    'bath towel', 'wall clock', 'led bulb', 'table lamp', 'shoe rack',
    'clothes hanger', 'cushion cover',
    
    # Fashion & Apparel
    'men t-shirt', 'formal shirt', 'polo shirt', 'jeans pant', 'panjabi',
    'cotton saree', 'women kurti', 'salwar kameez', 'hoodie', 'jacket',
    'sneakers', 'sandals', 'leather belt', 'leather wallet', 'sunglasses', 'wrist watch',
    
    # Electronics & Gadgets
    'wireless earbuds', 'bluetooth headphone', 'power bank', 'smart watch',
    'phone charger', 'type c cable', 'mobile back cover', 'screen protector',
    'usb pendrive', 'wireless mouse', 'gaming keyboard', 'laptop bag', 'wifi router', 'tripod stand',
    
    # Beauty & Personal Care
    'face wash', 'body lotion', 'sunscreen cream', 'shampoo', 'hair oil',
    'perfume', 'body spray', 'beard trimmer', 'hair dryer',
    
    # Health & Wellness
    'electric massager', 'blood pressure monitor', 'digital weighing scale', 'first aid kit',
    
    # Stationery & Office
    'notebook diary', 'gel pen set', 'desk organizer', 'scientific calculator',
    
    # Baby & Toys
    'baby diaper', 'baby wipes', 'feeding bottle', 'toy car', 'building blocks', 'puzzle board',
    
    # Tools & Hardware
    'screwdriver set', 'cordless drill', 'measuring tape', 'motorcycle helmet', 'bike cover', 'car cleaning kit'
]

# Stopwords & fluff words to clean product descriptions into concise tags
STOPWORDS = {
    'the', 'and', 'or', 'for', 'with', 'from', 'in', 'on', 'at', 'to', 'of', 'a', 'an', 'by',
    'best', 'quality', 'hot', 'sale', 'new', 'pcs', 'piece', 'set', 'pack', 'combo', 'original',
    'premium', 'imported', 'free', 'shipping', 'offer', 'brand', 'unisex', 'fashion', 'stylish',
    'high', 'low', 'price', 'delivery', 'bd', 'men', 'women', 'kids', 'boy', 'girl', 'adult',
    'top', 'super', 'portable', 'mini', 'latest', 'good', 'color', 'size', 'standard', 'heavy',
    'duty', 'collection', 'exclusive', 'special', 'item', 'product', '100%', 'grade', 'pro',
    'durable', 'pure', 'multi', 'smart', 'perfect', 'enthusiasts', 'style', 'buy', 'shop', 'online'
}

# Regex to detect phone numbers in text
PHONE_REGEX = re.compile(r'(?:\+?880\s?|0)(?:1[3-9]\d{2}[-\s]?\d{6}|[2-9]\d{6,8})\b')

# Regex to detect external website URLs (excluding Daraz/Lazada/Alibaba/CDN/Social Media)
URL_REGEX = re.compile(
    r'\b(?:https?://)?(?:www\.)?([a-zA-Z0-9][-a-zA-Z0-9]{1,62}\.(?:com|org|net|com\.bd|bd|shop|store|co|io|biz|info))\b',
    re.IGNORECASE
)
IGNORED_DOMAINS = {'daraz', 'lazada', 'alicdn', 'slatic', 'facebook', 'instagram', 'youtube', 'tiktok', 'google', 'twitter', 'apple'}


def fetch_catalog_page(query, page, retries=2):
    """Fetch one page of search results for a given query from Daraz catalog API."""
    url = f'https://www.daraz.com.bd/catalog/?q={requests.utils.quote(query)}&page={page}&ajax=true'
    for attempt in range(retries):
        try:
            r = requests.get(url, headers=HEADERS, timeout=12)
            if r.status_code == 200:
                data = r.json()
                items = data.get('mods', {}).get('listItems', [])
                return query, page, items, None
            elif r.status_code == 429:
                time.sleep(2)
        except Exception as e:
            if attempt == retries - 1:
                return query, page, [], str(e)
            time.sleep(1)
    return query, page, [], 'Max retries reached'


def extract_concise_product_types(titles, query_tags):
    """
    Synthesize a very short, clean list of product types (e.g., 'Tennis Ball, Sports Accessories')
    as specifically requested: 'keep it very short like tennis ball, accessories, plate, glass etc'
    """
    types_found = []
    
    # 1. First add the clean query tags associated with this seller
    for q in query_tags:
        cleaned_q = q.title()
        if cleaned_q not in types_found:
            types_found.append(cleaned_q)
            
    # 2. Extract 2-word noun pairs from product titles
    for title in titles:
        # remove punctuation
        clean = re.sub(r'[\(\)\[\]\{\}\-_/\\|,:;!+*&%#@?~`"\'<>]', ' ', title)
        words = [w.strip() for w in clean.split() if w.strip()]
        meaningful = [w.title() for w in words if w.lower() not in STOPWORDS and not w.isdigit() and len(w) > 2]
        
        if len(meaningful) >= 2:
            candidate = f"{meaningful[0]} {meaningful[1]}"
            if candidate not in types_found and not any(candidate in t or t in candidate for t in types_found):
                types_found.append(candidate)
        elif len(meaningful) == 1:
            candidate = meaningful[0]
            if candidate not in types_found and not any(candidate in t or t in candidate for t in types_found):
                types_found.append(candidate)
                
        if len(types_found) >= 4:
            break
            
    # Take at most 3-4 clean terms
    result = ', '.join(types_found[:4])
    return result if result else 'General Merchandise'


def scan_for_phone_and_website(seller_name, items):
    """Scan product titles, descriptions, and seller metadata for any listed phone or website."""
    phone = None
    website = None
    
    for it in items:
        text_blobs = [it.get('name', '')]
        desc = it.get('description', [])
        if isinstance(desc, list):
            text_blobs.extend(desc)
        elif isinstance(desc, str):
            text_blobs.append(desc)
            
        full_text = ' '.join(text_blobs)
        
        # Check for phone
        if not phone:
            m_phone = PHONE_REGEX.search(full_text)
            if m_phone:
                phone = m_phone.group(0).strip()
                
        # Check for external website
        if not website:
            m_url = URL_REGEX.search(full_text)
            if m_url:
                dom = m_url.group(1).lower()
                if not any(ign in dom for ign in IGNORED_DOMAINS):
                    website = m_url.group(0).strip()
                    if not website.startswith('http'):
                        website = 'https://' + website

    # Construct Daraz store URL
    clean_slug = seller_name.strip().replace(' ', '-')
    clean_slug = re.sub(r'[^a-zA-Z0-9\-]', '', clean_slug)
    daraz_shop_url = f"https://www.daraz.com.bd/shop/{clean_slug}"
    
    # If external website not found, Daraz shop URL serves as the store's primary web address
    web_field = website if website else daraz_shop_url
    phone_field = phone if phone else "Not publicly listed"
    
    return phone_field, web_field, daraz_shop_url


def main():
    print(f"Starting Daraz Seller Scraper across {len(QUERIES)} diverse categories...")
    tasks = []
    for q in QUERIES:
        for page in [1, 2]:  # pages 1 & 2 gives up to 80 items per query
            tasks.append((q, page))
            
    print(f"Total catalog requests queued: {len(tasks)}")
    
    raw_items_count = 0
    seller_items_map = {}  # seller_name -> list of items
    seller_queries_map = {} # seller_name -> set of queries
    
    t0 = time.time()
    with ThreadPoolExecutor(max_workers=8) as executor:
        future_to_task = {executor.submit(fetch_catalog_page, q, p): (q, p) for q, p in tasks}
        
        completed = 0
        for future in as_completed(future_to_task):
            completed += 1
            q, p = future_to_task[future]
            try:
                query, page, items, err = future.result()
                if err:
                    pass
                else:
                    raw_items_count += len(items)
                    for it in items:
                        seller = it.get('sellerName')
                        if not seller:
                            continue
                        seller = seller.strip()
                        if seller not in seller_items_map:
                            seller_items_map[seller] = []
                            seller_queries_map[seller] = set()
                        seller_items_map[seller].append(it)
                        seller_queries_map[seller].add(query)
            except Exception as exc:
                print(f"Task ({q}, {p}) generated an exception: {exc}")
                
            if completed % 25 == 0 or completed == len(tasks):
                print(f"[{completed}/{len(tasks)}] Scraped {raw_items_count} products | Found {len(seller_items_map)} unique sellers...")

    duration = round(time.time() - t0, 2)
    print(f"\nScraping complete in {duration}s! Extracted {raw_items_count} listings across {len(seller_items_map)} unique sellers.")
    
    # Process seller records
    records = []
    for seller_name, items in seller_items_map.items():
        # Representative titles
        titles = [it.get('name', '') for it in items if it.get('name')]
        query_tags = list(seller_queries_map.get(seller_name, []))
        
        # Product type (very short, e.g. "Tennis Ball, Accessories, Plate")
        product_type = extract_concise_product_types(titles, query_tags)
        
        # Phone and Website
        phone_val, website_val, daraz_shop_url = scan_for_phone_and_website(seller_name, items)
        
        # Seller ID and Shop ID
        seller_id = items[0].get('sellerId', '')
        shop_id = ''
        for it in items:
            ct = it.get('clickTrace', '')
            for part in ct.split(';'):
                if part.startswith('shop_id:'):
                    shop_id = part.split(':', 1)[1]
                    break
            if shop_id:
                break
                
        # Location
        locations = [it.get('location') for it in items if it.get('location')]
        loc = locations[0] if locations else 'Bangladesh'
        
        # Sample products
        sample_prods = ' | '.join(titles[:2])
        
        records.append({
            'Seller Name': seller_name,
            'Shop Name': seller_name,
            'Phone': phone_val,
            'Website': website_val,
            'Product Type': product_type,
            'Location': loc,
            'Store URL': daraz_shop_url,
            'Products in Catalog': len(items),
            'Sample Products': sample_prods,
            'Seller ID': seller_id,
            'Shop ID': shop_id
        })
        
    # Sort by number of products found (descending), then alphabetically
    records.sort(key=lambda x: (-x['Products in Catalog'], x['Seller Name']))
    
    df = pd.DataFrame(records)
    
    # Save CSV
    csv_path = 'c:/Users/User/Downloads/Daraz seller/Daraz_Sellers_List.csv'
    df.to_csv(csv_path, index=False, encoding='utf-8-sig')
    print(f"Saved CSV: {csv_path} with {len(df)} records.")
    
    # Save formatted Excel spreadsheet
    xlsx_path = 'c:/Users/User/Downloads/Daraz seller/Daraz_Sellers_List.xlsx'
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Daraz Sellers"
    
    # Styles
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    zebra_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    
    # Write headers
    headers = list(df.columns)
    ws.append(headers)
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border
    ws.row_dimensions[1].height = 28
    
    # Write rows
    for r_idx, row in enumerate(df.itertuples(index=False), start=2):
        ws.append(list(row))
        ws.row_dimensions[r_idx].height = 20
        use_zebra = (r_idx % 2 == 0)
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=r_idx, column=col_idx)
            cell.font = data_font
            cell.border = thin_border
            if use_zebra:
                cell.fill = zebra_fill
            if headers[col_idx - 1] in ['Phone', 'Location', 'Products in Catalog', 'Seller ID', 'Shop ID']:
                cell.alignment = center_align
            else:
                cell.alignment = left_align
                
    # Auto-adjust column widths
    for col_idx, col_name in enumerate(headers, start=1):
        max_len = len(col_name)
        for row in range(2, min(len(df) + 2, 500)):
            val = ws.cell(row=row, column=col_idx).value
            if val:
                max_len = max(max_len, len(str(val)))
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)
        
    # Freeze top row
    ws.freeze_panes = 'A2'
    
    wb.save(xlsx_path)
    print(f"Saved Excel: {xlsx_path} with formatting, freeze panes, and custom widths.")
    
    # Print sample top 10 sellers
    print("\n--- SAMPLE 10 SCRAPED SELLERS ---")
    for r in records[:10]:
        print(f"Seller: {r['Seller Name']} | Shop: {r['Shop Name']} | Types: {r['Product Type']} | Website: {r['Website']}")

if __name__ == '__main__':
    main()
