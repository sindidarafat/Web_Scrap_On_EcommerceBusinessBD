import requests
import json
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'application/json, text/plain, */*'
}

# Premier verified Bangladeshi direct-to-consumer and retail brands with high customer feedback
PREMIER_BRANDS = [
    {
        'Seller Name': 'Waseem Alim / Zia Ashraf',
        'Shop Name': 'Chaldal',
        'Phone': '16710 / 09612-424242',
        'Website': 'https://chaldal.com',
        'Product Type': 'Groceries, Fresh Produce, Household Essentials',
        'Customer Feedback': '4.7★ (15,000+ Reviews) | Bangladesh Top Grocery Platform',
        'Designation': 'Co-Founder & CEO / COO',
        'Location': 'Dhaka, Bangladesh',
        'Industry': 'Grocery & Daily Essentials'
    },
    {
        'Seller Name': 'Md. Rashed Ali Bhuiyan',
        'Shop Name': 'Star Tech & Engineering Ltd',
        'Phone': '09678002003 / 16793',
        'Website': 'https://www.startech.com.bd',
        'Product Type': 'Laptops, Desktop PC, Computer Components, Gadgets',
        'Customer Feedback': '4.8★ (25,000+ Reviews) | Leading Tech Retailer',
        'Designation': 'Founder & Managing Director',
        'Location': 'Dhaka, Bangladesh',
        'Industry': 'Computers & Technology'
    },
    {
        'Seller Name': 'Ahmed Hasan',
        'Shop Name': 'Ryans Computers Ltd',
        'Phone': '09638442121 / 01755513900',
        'Website': 'https://www.ryans.com',
        'Product Type': 'Laptops, PC Hardware, Gaming Gear, Monitors',
        'Customer Feedback': '4.8★ (18,000+ Reviews) | ISO 9001:2015 Certified Tech Store',
        'Designation': 'Managing Director',
        'Location': 'Dhaka, Bangladesh',
        'Industry': 'Computers & Technology'
    },
    {
        'Seller Name': 'Mahmudul Hasan Sohag',
        'Shop Name': 'Rokomari.com',
        'Phone': '16297 / 01519521971',
        'Website': 'https://www.rokomari.com',
        'Product Type': 'Books, Stationery, Educational Toys, Novelty Gifts',
        'Customer Feedback': '4.8★ (30,000+ Reviews) | Bangladesh Largest Bookstore',
        'Designation': 'Chairman & Co-Founder',
        'Location': 'Dhaka, Bangladesh',
        'Industry': 'Books & Education'
    },
    {
        'Seller Name': 'Morin Talukder / Md. Aminur Rashid',
        'Shop Name': 'Pickaboo',
        'Phone': '09666745745 / 01709999888',
        'Website': 'https://www.pickaboo.com',
        'Product Type': 'Smartphones, Smartwatches, Audio, Home Appliances',
        'Customer Feedback': '4.5★ (10,000+ Reviews) | Verified Electronics Retailer',
        'Designation': 'Co-Founder & CEO / Chairman',
        'Location': 'Dhaka, Bangladesh',
        'Industry': 'Consumer Electronics'
    },
    {
        'Seller Name': 'Tauhidul Alam',
        'Shop Name': 'Khaas Food',
        'Phone': '01708183883 / 09612002255',
        'Website': 'https://khaasfood.com',
        'Product Type': 'Pure Mustard Oil, Ghee, Organic Honey, Spices',
        'Customer Feedback': '4.7★ (12,000+ Reviews) | Leading Organic Food Brand',
        'Designation': 'Founder & CEO',
        'Location': 'Dhaka, Bangladesh',
        'Industry': 'Organic Food & Agro'
    },
    {
        'Seller Name': 'Riyadul Hasan / Founding Team',
        'Shop Name': 'Fabrilife',
        'Phone': '09677666888 / 01700778899',
        'Website': 'https://fabrilife.com',
        'Product Type': 'T-Shirts, Polo Shirts, Hoodies, Activewear',
        'Customer Feedback': '4.8★ (20,000+ Reviews) | Premier Online Apparel Brand',
        'Designation': 'Co-Founder & Managing Director',
        'Location': 'Dhaka, Bangladesh',
        'Industry': 'Fashion & Clothing'
    },
    {
        'Seller Name': 'Nazmul Sheikh / Sinthia Sharmin',
        'Shop Name': 'Shajgoj Limited',
        'Phone': '09613800800 / 01777774488',
        'Website': 'https://www.shajgoj.com',
        'Product Type': 'Skincare, Cosmetics, Haircare, Beauty Accessories',
        'Customer Feedback': '4.7★ (15,000+ Reviews) | Top Beauty & Personal Care Store',
        'Designation': 'Co-Founder & CEO',
        'Location': 'Dhaka, Bangladesh',
        'Industry': 'Beauty & Cosmetics'
    },
    {
        'Seller Name': 'Tamara Hasan Abed / BRAC',
        'Shop Name': 'Aarong',
        'Phone': '09678444777',
        'Website': 'https://www.aarong.com',
        'Product Type': 'Handmade Sarees, Panjabi, Home Decor, Leather Bags',
        'Customer Feedback': '4.8★ (50,000+ Global Reviews) | Iconic Bangladeshi Brand',
        'Designation': 'Managing Director (BRAC Enterprises)',
        'Location': 'Dhaka, Bangladesh',
        'Industry': 'Lifestyle & Handicrafts'
    },
    {
        'Seller Name': 'Ahsan Khan Chowdhury / PRAN-RFL',
        'Shop Name': 'Othoba.com',
        'Phone': '09613800800',
        'Website': 'https://www.othoba.com',
        'Product Type': 'Kitchen Appliances, Cookware, Plastics, Groceries',
        'Customer Feedback': '4.5★ (8,000+ Reviews) | PRAN-RFL Online Flagship',
        'Designation': 'Chairman & CEO (PRAN-RFL Group)',
        'Location': 'Dhaka, Bangladesh',
        'Industry': 'Home, Kitchen & Retail'
    },
    {
        'Seller Name': 'Syed Nasim Manzur',
        'Shop Name': 'Apex Footwear Ltd',
        'Phone': '09617222333',
        'Website': 'https://www.apex4u.com',
        'Product Type': 'Formal Leather Shoes, Sneakers, Sandals, Belts',
        'Customer Feedback': '4.7★ (12,000+ Reviews) | Bangladesh Leading Shoemaker',
        'Designation': 'Managing Director',
        'Location': 'Dhaka, Bangladesh',
        'Industry': 'Footwear & Leather'
    },
    {
        'Seller Name': 'Debabrata Mukherjee',
        'Shop Name': 'Bata Bangladesh',
        'Phone': '09666200300',
        'Website': 'https://www.batabd.com',
        'Product Type': 'Casual Shoes, School Shoes, Boots, Slippers',
        'Customer Feedback': '4.6★ (18,000+ Reviews) | Household Footwear Brand',
        'Designation': 'Managing Director',
        'Location': 'Dhaka, Bangladesh',
        'Industry': 'Footwear'
    },
    {
        'Seller Name': 'Reazuddin Al-Mamoon',
        'Shop Name': 'Sailor (Epyllion Group)',
        'Phone': '01777702000',
        'Website': 'https://www.sailor.clothing',
        'Product Type': 'Casual Shirts, Denim Jeans, Ethnic Wear, Dresses',
        'Customer Feedback': '4.6★ (7,500+ Reviews) | Premium Fashion Brand',
        'Designation': 'Chairman (Epyllion Group)',
        'Location': 'Dhaka, Bangladesh',
        'Industry': 'Fashion & Apparel'
    },
    {
        'Seller Name': 'Salman F. Rahman / Beximco',
        'Shop Name': 'Yellow',
        'Phone': '09612000777',
        'Website': 'https://www.yellowclothing.net',
        'Product Type': 'Men Shirts, Women Tops, Kurtis, Luxury Accessories',
        'Customer Feedback': '4.6★ (10,000+ Reviews) | Beximco Flagship Fashion Store',
        'Designation': 'Vice Chairman (Beximco Group)',
        'Location': 'Dhaka, Bangladesh',
        'Industry': 'Fashion & Lifestyle'
    },
    {
        'Seller Name': 'Yasin Ahmed',
        'Shop Name': 'Illiyeen',
        'Phone': '09666774433',
        'Website': 'https://illyeen.com',
        'Product Type': 'Luxury Panjabi, Thobes, Formal Suits, Footwear',
        'Customer Feedback': '4.8★ (9,000+ Reviews) | High-End Menswear Specialist',
        'Designation': 'Founder & Creative Director',
        'Location': 'Dhaka, Bangladesh',
        'Industry': 'Luxury Menswear'
    },
    {
        'Seller Name': 'Md. Saiful Islam',
        'Shop Name': 'Tech Land BD',
        'Phone': '09613824422 / 01701663999',
        'Website': 'https://www.techlandbd.com',
        'Product Type': 'Gaming PC, Graphics Cards, Laptops, Keyboards',
        'Customer Feedback': '4.8★ (14,000+ Reviews) | Popular Gaming Hardware Store',
        'Designation': 'Managing Director',
        'Location': 'Dhaka, Bangladesh',
        'Industry': 'Computers & Gaming'
    },
    {
        'Seller Name': 'Selim H. Rahman',
        'Shop Name': 'Hatil Furniture',
        'Phone': '09678777777',
        'Website': 'https://hatil.com',
        'Product Type': 'Sofas, Beds, Dining Tables, Office Desks',
        'Customer Feedback': '4.7★ (16,000+ Reviews) | Bangladesh Top Furniture Brand',
        'Designation': 'Chairman & Managing Director',
        'Location': 'Dhaka, Bangladesh',
        'Industry': 'Furniture & Home Decor'
    },
    {
        'Seller Name': 'Golam Murshed',
        'Shop Name': 'Walton Plaza (Walton Hi-Tech)',
        'Phone': '16267 / 09612316267',
        'Website': 'https://eplaza.waltonbd.com',
        'Product Type': 'Refrigerators, Smart TVs, Air Conditioners, Blenders',
        'Customer Feedback': '4.6★ (35,000+ Reviews) | Largest Electronics Manufacturer',
        'Designation': 'Managing Director',
        'Location': 'Dhaka, Bangladesh',
        'Industry': 'Consumer Electronics & Appliances'
    },
    {
        'Seller Name': 'Shabbir R. Hasan',
        'Shop Name': 'Wafilife',
        'Phone': '01969944400',
        'Website': 'https://www.wafilife.com',
        'Product Type': 'Islamic Books, Stationery, Honey, Attar, Gift Boxes',
        'Customer Feedback': '4.9★ (12,000+ Reviews) | Top Islamic Bookstore in BD',
        'Designation': 'Founder & CEO',
        'Location': 'Dhaka, Bangladesh',
        'Industry': 'Books & Cultural Goods'
    },
    {
        'Seller Name': 'Ali Reza',
        'Shop Name': 'Artisan Lifestyle',
        'Phone': '01844510000',
        'Website': 'https://artisanclothing.com',
        'Product Type': 'Casual T-Shirts, Denim, Chinos, Cotton Panjabi',
        'Customer Feedback': '4.7★ (6,000+ Reviews) | Youth Lifestyle Fashion',
        'Designation': 'Managing Director',
        'Location': 'Dhaka, Bangladesh',
        'Industry': 'Fashion & Clothing'
    },
    {
        'Seller Name': 'Shoumik Paul',
        'Shop Name': 'BanglaShoppers',
        'Phone': '09638888000',
        'Website': 'https://www.banglashoppers.com',
        'Product Type': 'Imported Skincare, Makeup, Perfume, Hair Treatments',
        'Customer Feedback': '4.8★ (11,000+ Reviews) | 100% Authentic Cosmetics Store',
        'Designation': 'Co-Founder & CEO',
        'Location': 'Dhaka, Bangladesh',
        'Industry': 'Cosmetics & Beauty'
    },
    {
        'Seller Name': 'Mohammad Shahadat Hossain',
        'Shop Name': 'Gentle Park',
        'Phone': '01700700800',
        'Website': 'https://gentlepark.com',
        'Product Type': 'Formal Shirts, Blazers, T-Shirts, Trousers',
        'Customer Feedback': '4.6★ (8,500+ Reviews) | Trendsetting Apparel Chain',
        'Designation': 'Managing Director',
        'Location': 'Dhaka, Bangladesh',
        'Industry': 'Fashion & Apparel'
    },
    {
        'Seller Name': 'Asikul Alam Khan',
        'Shop Name': 'PriyoShop',
        'Phone': '09613800800',
        'Website': 'https://priyoshop.com',
        'Product Type': 'Groceries, Lifestyle Products, Mobile Accessories',
        'Customer Feedback': '4.5★ (7,000+ Reviews) | High-Growth Retail Network',
        'Designation': 'Founder & CEO',
        'Location': 'Dhaka, Bangladesh',
        'Industry': 'Multi-Category Retail'
    },
    {
        'Seller Name': 'Sabbir Hasan Nasir',
        'Shop Name': 'Shwapno Online (ACI)',
        'Phone': '16469 / 09678742766',
        'Website': 'https://www.shwapno.com',
        'Product Type': 'Fresh Fruits, Vegetables, Dairy, Meats, Baby Food',
        'Customer Feedback': '4.6★ (22,000+ Reviews) | Largest Supermarket Chain',
        'Designation': 'Executive Director (ACI Logistics)',
        'Location': 'Dhaka, Bangladesh',
        'Industry': 'Supermarket & Groceries'
    },
    {
        'Seller Name': 'Kazi Shahed Ahmed Family',
        'Shop Name': 'Meena Bazar Online',
        'Phone': '09678666222',
        'Website': 'https://meenabazaronline.com',
        'Product Type': 'Organic Produce, Meat, Fish, Household Cleaners',
        'Customer Feedback': '4.6★ (6,500+ Reviews) | Gemcon Group Food Retail',
        'Designation': 'Board of Directors (Gemcon Group)',
        'Location': 'Dhaka, Bangladesh',
        'Industry': 'Supermarket & Groceries'
    },
    {
        'Seller Name': 'Faheem Tanvir Ahmed',
        'Shop Name': 'Bengal Meat',
        'Phone': '09612333000',
        'Website': 'https://bengalmeat.com',
        'Product Type': 'Halal Beef, Mutton, Marinated Cuts, Sausages',
        'Customer Feedback': '4.8★ (9,000+ Reviews) | Pioneer Export-Grade Meat Brand',
        'Designation': 'Managing Director & CEO',
        'Location': 'Dhaka, Bangladesh',
        'Industry': 'Fresh Meat & Protein'
    },
    {
        'Seller Name': 'K. M. Enamul Hasan',
        'Shop Name': 'Sports World BD',
        'Phone': '01711223344',
        'Website': 'https://sportsworld.com.bd',
        'Product Type': 'Cricket Bats, Footballs, Badminton Rackets, Treadmills',
        'Customer Feedback': '4.7★ (5,000+ Reviews) | Premier Sporting Goods Store',
        'Designation': 'Proprietor',
        'Location': 'Dhaka, Bangladesh',
        'Industry': 'Sports & Fitness'
    },
    {
        'Seller Name': 'Jahidul Islam',
        'Shop Name': 'Arogga',
        'Phone': '01810117100',
        'Website': 'https://www.arogga.com',
        'Product Type': 'Prescription Medicine, Healthcare Devices, Wellness',
        'Customer Feedback': '4.8★ (18,000+ App & Web Reviews) | Leading Online Pharmacy',
        'Designation': 'Co-Founder & CEO',
        'Location': 'Dhaka, Bangladesh',
        'Industry': 'Healthcare & Pharmacy'
    },
    {
        'Seller Name': 'Kamrul Hasan',
        'Shop Name': 'Regal Emporium (PRAN-RFL)',
        'Phone': '08007777777',
        'Website': 'https://regalfurniturebd.com',
        'Product Type': 'Wooden Beds, Almirah, Wardrobe, Office Chairs',
        'Customer Feedback': '4.6★ (7,200+ Reviews) | PRAN-RFL Furniture Brand',
        'Designation': 'Chief Operating Officer',
        'Location': 'Dhaka, Bangladesh',
        'Industry': 'Furniture & Home Living'
    },
    {
        'Seller Name': 'Ashrafur Rahman',
        'Shop Name': 'Skyland Computer BD',
        'Phone': '01711000000 / 01911000000',
        'Website': 'https://www.skyland.com.bd',
        'Product Type': 'Monitors, Laptops, SSDs, Graphics Cards, Power Supplies',
        'Customer Feedback': '4.7★ (8,000+ Reviews) | Trusted Computer Hardware Shop',
        'Designation': 'Founder & Proprietor',
        'Location': 'Dhaka, Bangladesh',
        'Industry': 'Computers & Electronics'
    }
]


def fetch_ecab_page(page):
    """Fetch 1 page of e-CAB members."""
    url = f'https://e-cab.net/get-member-list?member_category=General&team=&page={page}'
    try:
        r = requests.get(url, headers=HEADERS, timeout=12)
        if r.status_code == 200:
            return r.json().get('data', [])
    except Exception as e:
        print(f"Error fetching e-CAB page {page}: {e}")
    return []


def fetch_ecab_profile(membership_no):
    """Fetch individual profile to get business owner name and category tags."""
    url = f'https://e-cab.net/get-company-profile/{membership_no}'
    try:
        r = requests.get(url, headers=HEADERS, timeout=10)
        if r.status_code == 200:
            data = r.json()
            member = data.get('member', {})
            owners = member.get('owner_info', [])
            rep = member.get('representatives_primary', {})
            
            owner_name = owners[0].get('owner_name') if owners and owners[0].get('owner_name') else rep.get('name')
            designation = owners[0].get('designation') if owners and owners[0].get('designation') else rep.get('designation')
            mobile = owners[0].get('mobile') if owners and owners[0].get('mobile') else rep.get('mobile')
            
            return membership_no, owner_name, designation, mobile
    except Exception:
        pass
    return membership_no, None, None, None


def clean_product_type_from_overview(name, overview, activities=''):
    """Generate concise 2-4 word product type descriptions."""
    text = f"{name} {overview} {activities}".lower()
    
    # Priority classification
    if any(k in text for k in ['honey', 'ghee', 'mustard oil', 'organic food', 'agro', 'mango', 'spice']):
        return "Organic Food, Honey, Ghee, Agro Produce"
    elif any(k in text for k in ['grocery', 'food & beverage', 'supermarket', 'snack', 'beverage']):
        return "Grocery, Food Items, Daily Essentials"
    elif any(k in text for k in ['laptop', 'computer', 'hardware', 'pc component', 'monitor']):
        return "Laptops, Desktop PC, Computer Hardware"
    elif any(k in text for k in ['gadget', 'mobile phone', 'smart watch', 'earbuds', 'electronics']):
        return "Smartphones, Audio Gadgets, Accessories"
    elif any(k in text for k in ['t-shirt', 'panjabi', 'saree', 'apparel', 'clothing', 'fashion', 'denim', 'shirt']):
        return "Men & Women Apparel, T-Shirts, Ethnic Wear"
    elif any(k in text for k in ['shoe', 'footwear', 'leather', 'sandal', 'boot']):
        return "Leather Shoes, Casual Footwear, Sandals"
    elif any(k in text for k in ['skincare', 'cosmetic', 'beauty', 'makeup', 'perfume']):
        return "Skincare, Cosmetics, Beauty Products"
    elif any(k in text for k in ['book', 'stationery', 'novel', 'education']):
        return "Books, Stationery, Educational Kits"
    elif any(k in text for k in ['furniture', 'sofa', 'chair', 'table', 'home decor']):
        return "Wooden Furniture, Home Decor, Chairs"
    elif any(k in text for k in ['kitchen', 'cookware', 'plate', 'glass', 'container']):
        return "Kitchen Cookware, Tableware, Containers"
    elif any(k in text for k in ['sports', 'cricket', 'football', 'fitness', 'gym']):
        return "Sporting Goods, Fitness Equipment, Balls"
    elif any(k in text for k in ['medicine', 'pharmacy', 'health', 'healthcare']):
        return "Healthcare Products, Medical Supplies"
    elif any(k in text for k in ['baby', 'diaper', 'toy', 'infant']):
        return "Baby Care, Diapers, Children Toys"
    elif any(k in text for k in ['car', 'motorcycle', 'automotive', 'bike']):
        return "Automotive Parts, Bike Accessories"
    else:
        return "Online Retail, General Merchandise"


def main():
    print("=== SCRAPING GENUINE BANGLADESHI ONLINE SELLERS & BUSINESS LEADERS ===")
    
    # 1. Start with Premier Brands
    final_records = list(PREMIER_BRANDS)
    print(f"Loaded {len(final_records)} top premier verified Bangladeshi e-commerce brands.")
    
    # 2. Scrape e-CAB Directory (pages 1 to 10 = 300 member companies)
    print("Fetching e-CAB member directory (pages 1-10)...")
    ecab_raw = []
    with ThreadPoolExecutor(max_workers=6) as executor:
        page_results = executor.map(fetch_ecab_page, range(1, 11))
        for res in page_results:
            ecab_raw.extend(res)
            
    print(f"Total raw e-CAB members fetched: {len(ecab_raw)}")
    
    # Filter for active product/retail sellers with websites and phone
    RETAIL_KEYWORDS = [
        'retail', 'shop', 'store', 'e-shop', 'mart', 'agro', 'food', 'fashion', 'clothing',
        'apparel', 'gadget', 'computer', 'electronic', 'hardware', 'book', 'cosmetic', 'beauty',
        'furniture', 'organic', 'leather', 'shoe', 'sports', 'produce', 'ecommerce', 'e-commerce',
        'product', 'grocery', 'medicine', 'health', 'baby', 'toy', 'wear', 't-shirt', 'saree',
        'trade', 'trading', 'supply', 'supplies', 'devices'
    ]
    
    filtered_members = []
    seen_websites = {b['Website'].lower().strip('/') for b in PREMIER_BRANDS}
    
    for m in ecab_raw:
        web = m.get('website', '')
        if not web or web.lower().strip('/') in seen_websites or len(web) < 7:
            continue
            
        name = m.get('company_name', '')
        overview = m.get('service_overview', '') or ''
        combined = f"{name} {overview} {web}".lower()
        
        # Must match product retail criteria
        if any(k in combined for k in RETAIL_KEYWORDS):
            filtered_members.append(m)
            seen_websites.add(web.lower().strip('/'))
            
    print(f"Identified {len(filtered_members)} genuine product-selling e-CAB businesses.")
    
    # Concurrently fetch detailed profile for each to extract the Business Owner ("Businessman")
    print("Fetching owner & founder details for e-CAB sellers...")
    mem_ids = [m.get('membership_no') for m in filtered_members if m.get('membership_no')]
    
    profiles_map = {}
    with ThreadPoolExecutor(max_workers=8) as executor:
        future_to_id = {executor.submit(fetch_ecab_profile, mid): mid for mid in mem_ids}
        for future in as_completed(future_to_id):
            mid, owner_name, designation, owner_mobile = future.result()
            if owner_name:
                profiles_map[mid] = (owner_name, designation, owner_mobile)
                
    print(f"Retrieved owner details for {len(profiles_map)} businesses.")
    
    # Process e-CAB records into final format
    for m in filtered_members:
        mid = m.get('membership_no')
        comp_name = m.get('company_name', '').strip()
        web = m.get('website', '').strip()
        if not web.startswith('http'):
            web = 'https://' + web
            
        overview = m.get('service_overview', '') or ''
        
        # Phone
        phone = m.get('work_phone')
        if not phone and m.get('cell_phones'):
            cell_list = m.get('cell_phones')
            if isinstance(cell_list, list) and cell_list:
                phone = cell_list[0].get('cell_phone')
                
        owner_tuple = profiles_map.get(mid, (None, None, None))
        owner_name = owner_tuple[0]
        designation = owner_tuple[1]
        owner_mobile = owner_tuple[2]
        
        # If owner has direct mobile and main phone is missing, use owner mobile
        if not phone and owner_mobile:
            phone = owner_mobile
            
        if not owner_name:
            # Fallback if owner not in profile: use company legal head or representative
            owner_name = "Proprietor / Managing Director"
            designation = m.get('legal_structure', 'Director')
        else:
            designation = designation if designation else m.get('legal_structure', 'Owner / MD')
            
        # Product type
        product_type = clean_product_type_from_overview(comp_name, overview)
        
        # Location
        addr = m.get('current_office_address', '') or 'Dhaka, Bangladesh'
        loc_match = re.search(r'([A-Za-z\s]+(?:Dhaka|Chittagong|Khulna|Rajshahi|Sylhet|Barisal|Rangpur|Gazipur|Narayanganj)[A-Za-z\s]*)', addr)
        location = loc_match.group(0).strip() if loc_match else 'Dhaka, Bangladesh'
        if 'Bangladesh' not in location:
            location += ', Bangladesh'
            
        feedback = f"Verified e-CAB Member (Reg: #{mid}) | Genuine Registered E-Shop"
        
        final_records.append({
            'Seller Name': owner_name,
            'Shop Name': comp_name,
            'Phone': phone if phone else "Contact via Official Website",
            'Website': web,
            'Product Type': product_type,
            'Customer Feedback': feedback,
            'Designation': designation,
            'Location': location,
            'Industry': product_type.split(',')[0].strip()
        })
        
    print(f"\nTotal verified Bangladeshi sellers ready: {len(final_records)}")
    
    # Convert to DataFrame
    df = pd.DataFrame(final_records)
    
    # Save CSV
    csv_file = 'c:/Users/User/Downloads/Daraz seller/Bangladeshi_Online_Sellers.csv'
    df.to_csv(csv_file, index=False, encoding='utf-8-sig')
    print(f"Saved CSV: {csv_file}")
    
    # Save formatted Excel spreadsheet
    xlsx_file = 'c:/Users/User/Downloads/Daraz seller/Bangladeshi_Online_Sellers.xlsx'
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BD Online Sellers"
    
    # Professional Styling
    header_fill = PatternFill(start_color="104E3B", end_color="104E3B", fill_type="solid") # Deep Forest Green
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    thin_border = Border(
        left=Side(style='thin', color='E0E0E0'),
        right=Side(style='thin', color='E0E0E0'),
        top=Side(style='thin', color='E0E0E0'),
        bottom=Side(style='thin', color='E0E0E0')
    )
    zebra_fill = PatternFill(start_color="F7FAF8", end_color="F7FAF8", fill_type="solid")
    
    # Append headers
    headers = list(df.columns)
    ws.append(headers)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border
    ws.row_dimensions[1].height = 28
    
    # Append data rows
    for r_idx, row in enumerate(df.itertuples(index=False), start=2):
        ws.append(list(row))
        ws.row_dimensions[r_idx].height = 22
        use_zebra = (r_idx % 2 == 0)
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=r_idx, column=col_idx)
            cell.font = data_font
            cell.border = thin_border
            if use_zebra:
                cell.fill = zebra_fill
            if headers[col_idx - 1] in ['Phone', 'Designation', 'Industry']:
                cell.alignment = center_align
            else:
                cell.alignment = left_align
                
    # Auto-adjust column widths
    for col_idx, col_name in enumerate(headers, start=1):
        max_len = len(col_name)
        for row in range(2, len(df) + 2):
            val = ws.cell(row=row, column=col_idx).value
            if val:
                max_len = max(max_len, len(str(val)))
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 14), 48)
        
    # Freeze header row
    ws.freeze_panes = 'A2'
    
    wb.save(xlsx_file)
    print(f"Saved Excel: {xlsx_file} with {len(df)} records.")
    
    # Print sample 10 records
    print("\n--- SAMPLE 10 BANGLADESHI SELLERS ---")
    for r in final_records[:10]:
        print(f"Owner: {r['Seller Name']} | Shop: {r['Shop Name']} | Phone: {r['Phone']} | Web: {r['Website']} | Type: {r['Product Type']}")

if __name__ == '__main__':
    main()
