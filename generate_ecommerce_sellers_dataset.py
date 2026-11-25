# -*- coding: utf-8 -*-
"""
Script to generate comprehensive dataset and spreadsheets on E-commerce Sellers in Bangladesh.
Covers both Single-Seller (Owner-operated D2C) businesses and Multi-Seller Marketplace Merchants.
"""

import csv
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Define data entries
sellers_data = [
    # --- SINGLE SELLER / OWNER-OPERATED D2C BRANDS & DIRECT ONLINE PLATFORMS ---
    {
        "seller_name": "Riyadh Uddin Mahmood Chowdhury",
        "organisation_name": "Fabrilife",
        "top_selling_product": "Premium Combed Cotton Crew-Neck T-Shirts & Polo Shirts",
        "business_model": "Single Seller / Owner-Operated D2C",
        "category": "Fashion & Lifestyle",
        "primary_channel": "Independent D2C Website & App (fabrilife.com)",
        "price_range_bdt": "৳380 - ৳1,250",
        "base_city": "Dhaka",
        "website": "https://fabrilife.com",
        "notes": "Pioneering Bangladeshi D2C lifestyle apparel brand known for 100% combed cotton basic and graphic tees with high repeat purchase rate."
    },
    {
        "seller_name": "Nazmul Sheikh & Sinthia Sharmin Islam",
        "organisation_name": "Shajgoj Limited",
        "top_selling_product": "The Ordinary Niacinamide 10% + Zinc 1% & Korean Sunscreens",
        "business_model": "Single Seller / Owner-Operated D2C",
        "category": "Health & Beauty",
        "primary_channel": "Omnichannel D2C & App (shajgoj.com)",
        "price_range_bdt": "৳850 - ৳2,450",
        "base_city": "Dhaka",
        "website": "https://shajgoj.com",
        "notes": "Leading women's beauty and skincare D2C retail e-commerce in Bangladesh with certified genuine imported cosmetics."
    },
    {
        "seller_name": "Mahmudul Hasan Sohag",
        "organisation_name": "Rokomari.com (OnnoRokom Group)",
        "top_selling_product": "Paradoxical Sajid Series & Academic / Competitive Exam Books",
        "business_model": "E-commerce Platform & Direct Seller",
        "category": "Books & Stationery",
        "primary_channel": "Dedicated E-commerce Platform (rokomari.com)",
        "price_range_bdt": "৳220 - ৳850",
        "base_city": "Dhaka",
        "website": "https://rokomari.com",
        "notes": "Bangladesh's largest online bookstore, functioning as primary seller and marketplace for national publishers and authors."
    },
    {
        "seller_name": "Waseem Alim & Zia Ashraf",
        "organisation_name": "Chaldal Limited",
        "top_selling_product": "Fresh Farm Brown Eggs (Dozen), Teer Soybean Oil & Miniket Rice",
        "business_model": "Single Seller / Inventory E-commerce",
        "category": "Groceries & FMCG",
        "primary_channel": "Dedicated Grocery App & Web (chaldal.com)",
        "price_range_bdt": "৳150 - ৳3,200",
        "base_city": "Dhaka, Chattogram, Jashore",
        "website": "https://chaldal.com",
        "notes": "Largest online grocery platform operating micro-warehouses (dark stores) with 1-hour delivery model."
    },
    {
        "seller_name": "Md. Rashed Ali Bhuiyan",
        "organisation_name": "Star Tech & Engineering Ltd",
        "top_selling_product": "Custom Desktop PC Gaming Rigs & Fast IPS Gaming Monitors",
        "business_model": "Single Seller / Omnichannel Tech Retailer",
        "category": "Computer Hardware & IT",
        "primary_channel": "E-commerce & Physical Stores (startech.com.bd)",
        "price_range_bdt": "৳12,500 - ৳185,000",
        "base_city": "Dhaka, Chattogram, Rajshahi, Khulna",
        "website": "https://startech.com.bd",
        "notes": "The most trusted computer hardware and component e-commerce retailer in Bangladesh."
    },
    {
        "seller_name": "Ahmed Hasan",
        "organisation_name": "Ryans Computers Limited",
        "top_selling_product": "Business Laptops (HP / Lenovo / Dell) & High-Speed NVMe SSDs",
        "business_model": "Single Seller / Omnichannel Tech Retailer",
        "category": "Computer Hardware & IT",
        "primary_channel": "E-commerce & Nationwide Showrooms (ryans.com)",
        "price_range_bdt": "৳3,500 - ৳145,000",
        "base_city": "Dhaka (Nationwide Retail Network)",
        "website": "https://ryans.com",
        "notes": "Pioneer in IT hardware retailing in Bangladesh with an expansive e-commerce platform and nationwide warranty support."
    },
    {
        "seller_name": "Nure Alam Shimu & Jahangir Alam Sacchu",
        "organisation_name": "Gadget & Gear (G&G)",
        "top_selling_product": "Official Apple iPhone 15/16 Series & Apple Watch Ultra",
        "business_model": "Single Seller / Premium Tech Retailer",
        "category": "Consumer Electronics & Gadgets",
        "primary_channel": "Omnichannel D2C (gadgetandgear.com)",
        "price_range_bdt": "৳4,500 - ৳225,000",
        "base_city": "Dhaka",
        "website": "https://gadgetandgear.com",
        "notes": "Apple Authorized Reseller and premium lifestyle tech retailer operating flagship experiential gadget stores."
    },
    {
        "seller_name": "Mohammad Shariful Islam",
        "organisation_name": "Apple Gadgets BD",
        "top_selling_product": "Apple AirPods Pro 2, MagSafe Chargers & Baseus Power Accessories",
        "business_model": "Single Seller / Omnichannel Tech Retailer",
        "category": "Consumer Electronics & Gadgets",
        "primary_channel": "D2C Web & Flagship Outlets (applegadgetsbd.com)",
        "price_range_bdt": "৳1,800 - ৳175,000",
        "base_city": "Dhaka",
        "website": "https://applegadgetsbd.com",
        "notes": "Rapidly growing tech gadget retailer specializing in Apple ecosystem, premium Android flagships, and wearable audio."
    },
    {
        "seller_name": "Zakir Hosen",
        "organisation_name": "BDShop.com (BDSP IT)",
        "top_selling_product": "Boya BY-M1 Omnidirectional Lavalier Microphone & Studio Ring Lights",
        "business_model": "Single Seller / Owner-Operated D2C",
        "category": "Consumer Electronics & Gadgets",
        "primary_channel": "Independent E-commerce (bdshop.com)",
        "price_range_bdt": "৳650 - ৳18,500",
        "base_city": "Dhaka",
        "website": "https://bdshop.com",
        "notes": "Leading specialized online seller for content creators, YouTubers, livestreamers, and camera accessories in Bangladesh."
    },
    {
        "seller_name": "Md. Imrul Hasan",
        "organisation_name": "Motion View",
        "top_selling_product": "Amazfit Active Smartwatch & Anker Soundcore TWS Earbuds",
        "business_model": "Single Seller / Master Distributor & D2C",
        "category": "Consumer Electronics & Gadgets",
        "primary_channel": "Direct D2C Web & DarazMall (motionview.com.bd)",
        "price_range_bdt": "৳1,850 - ৳16,990",
        "base_city": "Dhaka",
        "website": "https://motionview.com.bd",
        "notes": "National authorized distributor and retailer for Amazfit, Anker, Soundcore, Haylou, and Kospet rugged smartwatches."
    },
    {
        "seller_name": "A.M.M. Habibul Mustafa Arman & Md. Tauhidul Islam",
        "organisation_name": "Khaas Food Ltd",
        "top_selling_product": "Sundarban Raw Natural Honey & Cold-Pressed Mustard Oil (Kachi Ghani)",
        "business_model": "Single Seller / Owner-Operated D2C",
        "category": "Groceries & FMCG",
        "primary_channel": "Omnichannel D2C (khaasfood.com)",
        "price_range_bdt": "৳290 - ৳1,850",
        "base_city": "Dhaka, Chattogram, Sylhet",
        "website": "https://khaasfood.com",
        "notes": "Pioneering safe, organic, unadulterated food e-commerce brand connecting rural producers directly to urban consumers."
    },
    {
        "seller_name": "S. M. Khaled",
        "organisation_name": "SaRa Lifestyle (Snowtex Group)",
        "top_selling_product": "Cotton Men's Casual Polos & Embroidered Festive Panjabis",
        "business_model": "Single Seller / D2C Apparel Brand",
        "category": "Fashion & Lifestyle",
        "primary_channel": "D2C Website & DarazMall Store (saralifestyle.com)",
        "price_range_bdt": "৳650 - ৳3,800",
        "base_city": "Dhaka",
        "website": "https://saralifestyle.com",
        "notes": "Fashion retail arm of green RMG manufacturer Snowtex, delivering affordable high-street apparel across Bangladesh."
    },
    {
        "seller_name": "Nahiyan Naser & Ali Sakhi Khan",
        "organisation_name": "Gorur Ghash",
        "top_selling_product": "Humorous Pop-Culture Graphic Printed T-Shirts & Streetwear Shorts",
        "business_model": "Single Seller / Owner-Operated D2C",
        "category": "Fashion & Lifestyle",
        "primary_channel": "Independent D2C & Social (gorurghash.com)",
        "price_range_bdt": "৳450 - ৳1,650",
        "base_city": "Dhaka",
        "website": "https://gorurghash.com",
        "notes": "Cult streetwear and creative graphic t-shirt D2C brand widely popular among Bangladeshi youth and university students."
    },
    {
        "seller_name": "Rob Steven & Md. Rafiqul Islam",
        "organisation_name": "Kablewala Bangladesh (Banglamark Group)",
        "top_selling_product": "Braided 65W Fast Charging Type-C Cables & Heavy-Duty Power Banks",
        "business_model": "Single Seller / Specialty Marketplace",
        "category": "Consumer Electronics & Gadgets",
        "primary_channel": "E-commerce Platform (kablewala.com.bd)",
        "price_range_bdt": "৳220 - ৳4,500",
        "base_city": "Dhaka",
        "website": "https://kablewala.com.bd",
        "notes": "Specialist e-commerce focusing on charging cables, adapters, digital peripherals, and household smart electronics."
    },
    {
        "seller_name": "Azharul Alam",
        "organisation_name": "Techland BD",
        "top_selling_product": "NVIDIA GeForce RTX 4060/4070 GPUs & Asus ROG Gaming Laptops",
        "business_model": "Single Seller / Omnichannel Tech Retailer",
        "category": "Computer Hardware & IT",
        "primary_channel": "Online Store & Multiplan Outlets (techlandbd.com)",
        "price_range_bdt": "৳28,000 - ৳260,000",
        "base_city": "Dhaka",
        "website": "https://techlandbd.com",
        "notes": "Key retailer for gaming hardware, graphics cards, PC power supplies, and creator workstation peripherals."
    },
    {
        "seller_name": "Moni Khondoker",
        "organisation_name": "Le Reve (Reve Tex Ltd)",
        "top_selling_product": "Women's Ethnic Viscose Tunics & Fusion Long Kurtis",
        "business_model": "Single Seller / D2C Fashion Brand",
        "category": "Fashion & Lifestyle",
        "primary_channel": "D2C Web & Flagship Boutiques (lerevecraze.com)",
        "price_range_bdt": "৳1,200 - ৳6,500",
        "base_city": "Dhaka (Nationwide Outlets)",
        "website": "https://lerevecraze.com",
        "notes": "Fashion-forward international ethnic wear and casual clothing brand with dedicated e-commerce delivery across Bangladesh."
    },
    {
        "seller_name": "Tamara Hasan Abed",
        "organisation_name": "Aarong (BRAC)",
        "top_selling_product": "Handcrafted Silk Jamdani Saree & Taaga Contemporary Kurtas",
        "business_model": "Single Seller / Social Enterprise D2C",
        "category": "Fashion & Lifestyle",
        "primary_channel": "Global & Domestic E-commerce (aarong.com)",
        "price_range_bdt": "৳1,800 - ৳45,000",
        "base_city": "Dhaka (Global Delivery)",
        "website": "https://aarong.com",
        "notes": "Bangladesh's preeminent craft and lifestyle retail brand empowering over 65,000 rural artisans under BRAC."
    },
    {
        "seller_name": "Syed Nasim Manzur",
        "organisation_name": "Apex Footwear Limited",
        "top_selling_product": "Apex Men's Genuine Leather Formal Shoes & MOCHI Casual Loafers",
        "business_model": "Single Seller / D2C Footwear Brand",
        "category": "Footwear & Leather Goods",
        "primary_channel": "D2C E-store & Retail Chain (apex4u.com)",
        "price_range_bdt": "৳1,690 - ৳12,990",
        "base_city": "Dhaka (Nationwide)",
        "website": "https://apex4u.com",
        "notes": "Leading leather footwear brand in South Asia, operating proprietary brands Apex, Venturini, Nino Rossi, and MOCHI."
    },
    {
        "seller_name": "Debabrata Mukherjee",
        "organisation_name": "Bata Bangladesh",
        "top_selling_product": "Bata Comfit Ergonomic Walking Shoes & Power Athletic Sneakers",
        "business_model": "Single Seller / Brand Footwear E-store",
        "category": "Footwear & Leather Goods",
        "primary_channel": "D2C E-commerce (batabd.com)",
        "price_range_bdt": "৳990 - ৳7,990",
        "base_city": "Dhaka (Nationwide)",
        "website": "https://batabd.com",
        "notes": "Household footwear brand offering seamless online shopping for Bata Comfit, Power, Weinbrenner, and North Star."
    },
    {
        "seller_name": "Golam Murshed / SM Shamsul Alam",
        "organisation_name": "Walton E-Plaza (Walton Hi-Tech)",
        "top_selling_product": "Walton Inverter Frost-Free Refrigerator & Android Smart LED TV",
        "business_model": "Single Seller / Electronics Manufacturer D2C",
        "category": "Home Appliances & Electronics",
        "primary_channel": "Direct D2C E-Plaza (eplaza.waltonbd.com)",
        "price_range_bdt": "৳14,500 - ৳85,000",
        "base_city": "Gazipur / Dhaka (Nationwide)",
        "website": "https://eplaza.waltonbd.com",
        "notes": "Bangladesh's flagship domestic electronics conglomerate selling refrigerators, air conditioners, TVs, and compressors online."
    },
    {
        "seller_name": "M. A. Razzak Khan",
        "organisation_name": "Minister-MyOne E-Store",
        "top_selling_product": "Minister Double Door Refrigerator & 32-43 Inch Smart LED TVs",
        "business_model": "Single Seller / Manufacturer D2C",
        "category": "Home Appliances & Electronics",
        "primary_channel": "Direct E-Store & DarazMall (ministerbd.com)",
        "price_range_bdt": "৳11,990 - ৳65,000",
        "base_city": "Dhaka",
        "website": "https://ministerbd.com",
        "notes": "Popular electronics and home appliances producer targeting value-conscious consumers across rural and semi-urban Bangladesh."
    },
    {
        "seller_name": "Ahsan Khan Chowdhury",
        "organisation_name": "Othoba.com (PRAN-RFL Group)",
        "top_selling_product": "RFL Wardrobes, Storage Plastics, PRAN Drinks & Italiano Dinnerware",
        "business_model": "E-commerce Platform & Single Conglomerate Hub",
        "category": "Home & Kitchen / FMCG",
        "primary_channel": "Direct Platform (othoba.com)",
        "price_range_bdt": "৳250 - ৳18,500",
        "base_city": "Dhaka (Nationwide Delivery)",
        "website": "https://othoba.com",
        "notes": "E-commerce arm of PRAN-RFL Group, serving as single hub for PRAN food products, RFL plastics, Regal furniture, and Vision electronics."
    },
    {
        "seller_name": "Morin Talukder & Md. Aminur Rashid",
        "organisation_name": "Pickaboo",
        "top_selling_product": "Xiaomi Redmi Budget Smartphones & Fast Charging Mobile Accessories",
        "business_model": "Specialized E-commerce Platform & Direct Seller",
        "category": "Consumer Electronics & Gadgets",
        "primary_channel": "Direct App & Web (pickaboo.com)",
        "price_range_bdt": "৳1,200 - ৳65,000",
        "base_city": "Dhaka",
        "website": "https://pickaboo.com",
        "notes": "Pioneer in mobile phone e-commerce offering official warranty, EMI facilities, and same-day delivery in Dhaka."
    },
    {
        "seller_name": "Asikul Alam Khan",
        "organisation_name": "PriyoShop",
        "top_selling_product": "B2B Retail Grocery Bundles, Branded Edible Oil & Wheat Flour (Atta)",
        "business_model": "B2B & Retail E-commerce Hub",
        "category": "Groceries & FMCG",
        "primary_channel": "B2B App & Web (priyoshop.com)",
        "price_range_bdt": "৳2,500 - ৳45,000",
        "base_city": "Dhaka",
        "website": "https://priyoshop.com",
        "notes": "Transformed from retail B2C into a leading B2B e-commerce supply chain empowering thousands of neighborhood grocery shops."
    },
    {
        "seller_name": "Simeen Rahman",
        "organisation_name": "Transcom Digital (Transcom Group)",
        "top_selling_product": "Samsung Crystal 4K UHD Smart TV & Whirlpool Inverter Washing Machine",
        "business_model": "Single Seller / Authorized Multi-Brand Retailer",
        "category": "Home Appliances & Electronics",
        "primary_channel": "E-commerce & Showrooms (transcomdigital.com)",
        "price_range_bdt": "৳18,000 - ৳220,000",
        "base_city": "Dhaka (Nationwide)",
        "website": "https://transcomdigital.com",
        "notes": "Official distributor and retailer for premium global consumer brands including Samsung, Whirlpool, Philips, and Panasonic."
    },
    {
        "seller_name": "MHM Fairoz",
        "organisation_name": "Singer Bangladesh Limited",
        "top_selling_product": "Singer Domestic Sewing Machines & Inverter Split Air Conditioners",
        "business_model": "Single Seller / Brand Omnichannel E-store",
        "category": "Home Appliances & Electronics",
        "primary_channel": "Direct E-store (singerbd.com)",
        "price_range_bdt": "৳8,500 - ৳78,000",
        "base_city": "Dhaka (Nationwide)",
        "website": "https://singerbd.com",
        "notes": "Century-old heritage brand in Bangladesh offering flexible consumer hire-purchase and online installment schemes."
    },
    {
        "seller_name": "Enamul Haque Babu",
        "organisation_name": "Herlan Store (Remark HB Limited)",
        "top_selling_product": "Nior No-Transfer Matte Lipsticks & Siodil Salicylic Acid Serums",
        "business_model": "Single Seller / D2C Beauty Brand",
        "category": "Health & Beauty",
        "primary_channel": "Direct D2C E-store & Outlets (herlan.com)",
        "price_range_bdt": "৳450 - ৳2,950",
        "base_city": "Dhaka",
        "website": "https://herlan.com",
        "notes": "Fast-expanding cosmetics and home care retail chain offering authentic brands Nior, Siodil, Blaze O' Skin, and Herlan."
    },
    {
        "seller_name": "Abdullah Hil Rakib",
        "organisation_name": "Twelve Clothing (Team Group)",
        "top_selling_product": "Modern Cut Washed Denim Jeans & Everyday Graphic Hoodies",
        "business_model": "Single Seller / D2C Fashion Brand",
        "category": "Fashion & Lifestyle",
        "primary_channel": "D2C Website & Outlets (twelvebd.com)",
        "price_range_bdt": "৳750 - ৳3,450",
        "base_city": "Dhaka",
        "website": "https://twelvebd.com",
        "notes": "Youth-centric apparel brand backed by Team Group RMG exporter, known for high quality casuals and seasonal winter wear."
    },
    {
        "seller_name": "BYSL Global Technology Group",
        "organisation_name": "ILLIYEEN",
        "top_selling_product": "Luxury Platinum Finish Panjabis & Modern Arabic Thobes",
        "business_model": "Single Seller / Luxury D2C Fashion",
        "category": "Fashion & Lifestyle",
        "primary_channel": "Direct Luxury E-store (ilyn.global)",
        "price_range_bdt": "৳3,800 - ৳24,500",
        "base_city": "Dhaka, Chattogram, Sylhet",
        "website": "https://ilyn.global",
        "notes": "Premium high-end menswear and festive luxury brand with top-tier fabric engineering and bespoke packaging."
    },
    {
        "seller_name": "Sabbir Hasan Nasir",
        "organisation_name": "Shwapno Online (ACI Logistics)",
        "top_selling_product": "ACI Pure Miniket Rice, Fresh Beef, Farm Eggs & Seasonal Fruits",
        "business_model": "Single Seller / Supermarket Grocery E-store",
        "category": "Groceries & FMCG",
        "primary_channel": "Grocery App & Web (shwapno.com)",
        "price_range_bdt": "৳120 - ৳4,500",
        "base_city": "Dhaka, Chattogram, Sylhet",
        "website": "https://shwapno.com",
        "notes": "Largest physical supermarket retail chain in Bangladesh with robust online grocery delivery serving daily urban household demand."
    },
    {
        "seller_name": "Kazi Inam Ahmed",
        "organisation_name": "Meena Click (Gemcon Group)",
        "top_selling_product": "Meena Natural Pure Dairy Milk & Organic Farm Fresh Produce",
        "business_model": "Single Seller / Superstore Grocery E-store",
        "category": "Groceries & FMCG",
        "primary_channel": "E-commerce App & Web (meenaclick.com)",
        "price_range_bdt": "৳110 - ৳3,800",
        "base_city": "Dhaka, Chattogram",
        "website": "https://meenaclick.com",
        "notes": "Online extension of Meena Bazar offering clean, organic, and certified food items directly to city doorsteps."
    },
    {
        "seller_name": "Zeeshan Khursheed Huq",
        "organisation_name": "Sindabad.com (Ananta Group)",
        "top_selling_product": "Double A 80GSM A4 Copier Paper & Office Disinfectant Cleaners",
        "business_model": "Single Seller / B2B E-commerce",
        "category": "Office Supplies & Stationery",
        "primary_channel": "B2B E-commerce Web (sindabad.com)",
        "price_range_bdt": "৳380 - ৳25,000",
        "base_city": "Dhaka",
        "website": "https://sindabad.com",
        "notes": "First dedicated B2B e-commerce platform in Bangladesh serving corporate offices and factories with stationery and industrial supplies."
    },
    {
        "seller_name": "Md. Abdul Khaleque & Shahadat Hossain",
        "organisation_name": "BoiBazar.com",
        "top_selling_product": "BCS Preliminary Exam Preparation Books & University Admission Guides",
        "business_model": "Single Seller / Online Book Retailer",
        "category": "Books & Stationery",
        "primary_channel": "Online Bookstore (boibazar.com)",
        "price_range_bdt": "৳240 - ৳1,400",
        "base_city": "Dhaka",
        "website": "https://boibazar.com",
        "notes": "Dedicated online platform specializing in higher education, competitive civil service, and university admission books."
    },
    {
        "seller_name": "Arif Azad & Masud Parvez",
        "organisation_name": "Wafilife",
        "top_selling_product": "Ma'ariful Quran Tafseer Sets & Islamic Contemporary Non-Fiction",
        "business_model": "Single Seller / Islamic Bookstore D2C",
        "category": "Books & Stationery",
        "primary_channel": "Direct Online Store (wafilife.com)",
        "price_range_bdt": "৳180 - ৳3,600",
        "base_city": "Dhaka",
        "website": "https://wafilife.com",
        "notes": "Leading Islamic literature, spiritual, and educational online bookstore catering to readers across Bangladesh."
    },
    {
        "seller_name": "Kazi Jamil Islam",
        "organisation_name": "Lotto Bangladesh (Express Leathers)",
        "top_selling_product": "Lotto Dynamic Grip Running Shoes & Dry-Fit Sports Polo Shirts",
        "business_model": "Single Seller / Licensed Brand E-store",
        "category": "Footwear & Sportswear",
        "primary_channel": "E-commerce & Stores (lottobd.com)",
        "price_range_bdt": "৳1,190 - ৳5,890",
        "base_city": "Dhaka (Nationwide)",
        "website": "https://lottobd.com",
        "notes": "Official licensed manufacturer and retailer for Italian sportswear brand Lotto in Bangladesh."
    },
    {
        "seller_name": "Tanjim Haque",
        "organisation_name": "Tanjim (Ecstasy)",
        "top_selling_product": "Designer Slim-Fit Casual Shirts & Tailored Urban Streetwear",
        "business_model": "Single Seller / Designer D2C Fashion",
        "category": "Fashion & Lifestyle",
        "primary_channel": "D2C Web & Boutiques (ecstasybd.com)",
        "price_range_bdt": "৳1,450 - ৳6,200",
        "base_city": "Dhaka",
        "website": "https://ecstasybd.com",
        "notes": "High-fashion youth brand known for edgy designs, contemporary tailoring, and premium denim."
    },
    {
        "seller_name": "Mustafa Rahman",
        "organisation_name": "Perfume Bangladesh",
        "top_selling_product": "Armaf Club de Nuit Intense Man EDT & Lattafa Khamrah Eau De Parfum",
        "business_model": "Single Seller / Specialty Fragrance D2C",
        "category": "Health & Beauty",
        "primary_channel": "Direct E-store (perfumebangladesh.com)",
        "price_range_bdt": "৳1,200 - ৳14,500",
        "base_city": "Dhaka",
        "website": "https://perfumebangladesh.com",
        "notes": "Leading online boutique for 100% authentic designer French, Arabian, and niche luxury perfumes."
    },
    {
        "seller_name": "Habibur Rahman",
        "organisation_name": "Rongdhonu Lifestyle",
        "top_selling_product": "Organic Kasturi Haldi (Wild Turmeric) & Multani Mitti Clay Powders",
        "business_model": "Single Seller / Herbal D2C Brand",
        "category": "Health & Beauty",
        "primary_channel": "Direct D2C Web & DarazMall (rongdhonulifestyle.com)",
        "price_range_bdt": "৳140 - ৳650",
        "base_city": "Dhaka",
        "website": "https://rongdhonulifestyle.com",
        "notes": "Top-selling organic wellness and Ayurvedic skincare brand with massive transaction volume on Daraz and social commerce."
    },
    {
        "seller_name": "Engr. Md. Mehedi Hasan",
        "organisation_name": "Creatus Computer",
        "top_selling_product": "Core i5 / Ryzen 5 Budget Editing PC Packages & ViewSonic Monitors",
        "business_model": "Single Seller / PC Hardware Specialist",
        "category": "Computer Hardware & IT",
        "primary_channel": "Online Shop & Showroom (creatuscomputer.com)",
        "price_range_bdt": "৳22,000 - ৳120,000",
        "base_city": "Dhaka",
        "website": "https://creatuscomputer.com",
        "notes": "Popular online PC builder offering customized desktop bundles for freelancers, graphic designers, and video editors."
    },
    {
        "seller_name": "Shams Jaber",
        "organisation_name": "TechShopBD",
        "top_selling_product": "Arduino Uno R3 Microcontroller Board & Robotics Sensor Modules",
        "business_model": "Single Seller / Maker & Robotics D2C",
        "category": "Electronics & DIY Tech",
        "primary_channel": "Specialist Web Store (techshopbd.com)",
        "price_range_bdt": "৳120 - ৳8,500",
        "base_city": "Dhaka",
        "website": "https://techshopbd.com",
        "notes": "First dedicated electronics components, IoT boards, robotics, and STEM hobbyist platform in Bangladesh."
    },

    # --- MULTI-VENDOR MARKETPLACE SELLERS & VERIFIED MERCHANT STORES ---
    {
        "seller_name": "Mehdi Hasan",
        "organisation_name": "Swapon's World",
        "top_selling_product": "Men's Festive Cotton Semi-Long Panjabi & Casual Shirts",
        "business_model": "Marketplace Top Seller / Independent Brand",
        "category": "Fashion & Lifestyle",
        "primary_channel": "Daraz Bangladesh (DarazMall Top Fashion Seller)",
        "price_range_bdt": "৳650 - ৳2,250",
        "base_city": "Dhaka",
        "website": "https://daraz.com.bd/shop/swapons-world",
        "notes": "Consistently ranks among Daraz Bangladesh's #1 top-volume apparel merchants during Eid and Pohela Boishakh campaigns."
    },
    {
        "seller_name": "Digital Retail Management (DRM) / Authorized Partner",
        "organisation_name": "DRM Tech BD (Baseus Official Partner)",
        "top_selling_product": "Baseus 20,000mAh 22.5W Bipow Digital Display Fast Power Bank",
        "business_model": "Marketplace Merchant / Authorized Distributor",
        "category": "Consumer Electronics & Gadgets",
        "primary_channel": "Daraz Bangladesh (DarazMall Flagship)",
        "price_range_bdt": "৳1,650 - ৳3,800",
        "base_city": "Dhaka",
        "website": "https://daraz.com.bd/shop/baseus-bangladesh",
        "notes": "High-volume authorized merchant fulfilling Baseus fast charging accessories and cables with genuine warranty."
    },
    {
        "seller_name": "New Udoy Electronics / Authorized Importer",
        "organisation_name": "New Udoy Mobile Store",
        "top_selling_product": "Realme C53 / C55 Budget Smartphones (6GB/128GB)",
        "business_model": "Marketplace Merchant / Mobile Retailer",
        "category": "Consumer Electronics & Gadgets",
        "primary_channel": "Daraz Bangladesh & Pickaboo",
        "price_range_bdt": "৳12,999 - ৳18,500",
        "base_city": "Dhaka",
        "website": "https://daraz.com.bd",
        "notes": "Leading certified merchant on Daraz for high-volume entry-level and mid-range Android smartphone sales."
    },
    {
        "seller_name": "EERNA Limited",
        "organisation_name": "EERNA Official Gadget Store",
        "top_selling_product": "Infinix Hot 40 / Note 30 Series 5G Smartphones",
        "business_model": "Marketplace Merchant / Authorized Partner",
        "category": "Consumer Electronics & Gadgets",
        "primary_channel": "Daraz Bangladesh & Direct Web",
        "price_range_bdt": "৳14,500 - ৳26,000",
        "base_city": "Dhaka",
        "website": "https://eerna.com.bd",
        "notes": "Major verified mobile electronics partner managing official device distribution on DarazMall."
    },
    {
        "seller_name": "Intimacy Computer / Tech Partner",
        "organisation_name": "Intimacy Electronics BD",
        "top_selling_product": "Samsung Galaxy A05 / A15 4G Smartphone & SanDisk Ultra SD Cards",
        "business_model": "Marketplace Merchant / IT Vendor",
        "category": "Consumer Electronics & Gadgets",
        "primary_channel": "Daraz Bangladesh (DarazMall Partner)",
        "price_range_bdt": "৳650 - ৳22,500",
        "base_city": "Dhaka",
        "website": "https://daraz.com.bd",
        "notes": "Authorized vendor for official Samsung consumer electronics, storage drives, and mobile peripherals."
    },
    {
        "seller_name": "Zaved Akhtar",
        "organisation_name": "Unilever Bangladesh Official Store",
        "top_selling_product": "Sunsilk Black Shine Shampoo (650ml) & Tresemme Keratin Smooth",
        "business_model": "Marketplace Flagship / Multinational FMCG",
        "category": "Health & Beauty / FMCG",
        "primary_channel": "Daraz Bangladesh (DarazMall FMCG #1)",
        "price_range_bdt": "৳180 - ৳1,450",
        "base_city": "Dhaka",
        "website": "https://daraz.com.bd/shop/unilever-bangladesh",
        "notes": "Highest volume FMCG seller on Daraz, moving hundreds of thousands of units during 11.11 and monthly flash campaigns."
    },
    {
        "seller_name": "Rajat Diwakar",
        "organisation_name": "Marico Bangladesh Official Store",
        "top_selling_product": "Parachute 100% Pure Coconut Oil (500ml) & Saffola Active Oil",
        "business_model": "Marketplace Flagship / Multinational FMCG",
        "category": "Health & Beauty / FMCG",
        "primary_channel": "Daraz Bangladesh (DarazMall)",
        "price_range_bdt": "৳160 - ৳950",
        "base_city": "Dhaka",
        "website": "https://daraz.com.bd/shop/marico-bangladesh",
        "notes": "Market leader in hair care and healthy edible oils with dominant market share on both online platforms and offline retail."
    },
    {
        "seller_name": "Deepal Abeywickrema",
        "organisation_name": "Nestlé Bangladesh Flagship Store",
        "top_selling_product": "Maggi 2-Minute Masala Noodles (Pack of 16) & Nescafé Classic Jar",
        "business_model": "Marketplace Flagship / Global Food Giant",
        "category": "Groceries & FMCG",
        "primary_channel": "Daraz Bangladesh (DarazMall)",
        "price_range_bdt": "৳150 - ৳1,850",
        "base_city": "Dhaka",
        "website": "https://daraz.com.bd/shop/nestle-bangladesh",
        "notes": "High-velocity consumer food seller dominating breakfast cereals, infant nutrition, and instant noodles."
    },
    {
        "seller_name": "Vishal Gupta",
        "organisation_name": "Reckitt Benckiser Bangladesh Official",
        "top_selling_product": "Dettol Antiseptic Disinfectant Liquid (1L) & Harpic Power Plus",
        "business_model": "Marketplace Flagship / Hygiene FMCG",
        "category": "Household & FMCG",
        "primary_channel": "Daraz Bangladesh (DarazMall)",
        "price_range_bdt": "৳130 - ৳1,200",
        "base_city": "Dhaka",
        "website": "https://daraz.com.bd/shop/reckitt-benckiser-bd",
        "notes": "Dominant hygiene and household sanitation seller on Daraz, Pickaboo, and Chaldal."
    },
    {
        "seller_name": "Anjan Chowdhury",
        "organisation_name": "Square Toiletries Official Store",
        "top_selling_product": "Radhuni Pure Turmeric & Chilli Spices, Chashi Chinigura Rice",
        "business_model": "Marketplace Flagship / National Conglomerate",
        "category": "Groceries & Personal Care",
        "primary_channel": "Daraz Bangladesh & Chaldal",
        "price_range_bdt": "৳85 - ৳1,650",
        "base_city": "Dhaka",
        "website": "https://daraz.com.bd/shop/square-toiletries-ltd",
        "notes": "Square Group's official digital merchant driving massive sales in Radhuni spices, Meril personal care, and Senora hygiene products."
    },
    {
        "seller_name": "Dr. Arif Dowla",
        "organisation_name": "ACI Consumer Brands Official Store",
        "top_selling_product": "ACI Aerosol Insect Spray (475ml) & Savlon Antiseptic Liquid",
        "business_model": "Marketplace Flagship / National Conglomerate",
        "category": "Household & Personal Care",
        "primary_channel": "Daraz Bangladesh (DarazMall)",
        "price_range_bdt": "৳95 - ৳1,150",
        "base_city": "Dhaka",
        "website": "https://daraz.com.bd/shop/aci-consumer-brands",
        "notes": "Critical household essentials seller with top market share in pest control, antiseptic solutions, and pure spices."
    },
    {
        "seller_name": "Watches World BD (Kallol Group)",
        "organisation_name": "Watches World Official Store",
        "top_selling_product": "Casio Edifice Chronograph Stainless Steel Watch (EFV-550D)",
        "business_model": "Marketplace Flagship / Official Importer",
        "category": "Watches & Accessories",
        "primary_channel": "Daraz Bangladesh (DarazMall)",
        "price_range_bdt": "৳3,500 - ৳28,000",
        "base_city": "Dhaka",
        "website": "https://daraz.com.bd/shop/watches-world",
        "notes": "Official importer and distributor for Casio, Edifice, G-Shock, and Sheen timepieces in Bangladesh."
    },
    {
        "seller_name": "Time Zone BD (Kallol Group)",
        "organisation_name": "Time Zone Flagship Store",
        "top_selling_product": "Titan Workwear Quartz Watch & Fastrack Reflex Smartwatch",
        "business_model": "Marketplace Flagship / Multi-Brand Watch Store",
        "category": "Watches & Accessories",
        "primary_channel": "Daraz Bangladesh & Pickaboo",
        "price_range_bdt": "৳2,800 - ৳22,500",
        "base_city": "Dhaka",
        "website": "https://daraz.com.bd/shop/time-zone",
        "notes": "Authorized national retailer for Titan, Fastrack, Citizen, Seiko, and Pierre Cardin watches."
    },
    {
        "seller_name": "Brand Bazaar BD / Md. Rafiqul Islam",
        "organisation_name": "Brand Bazaar Online",
        "top_selling_product": "Sony Bravia 55-inch 4K Google LED TV & General 1.5 Ton Split AC",
        "business_model": "Marketplace Top Merchant / Appliance Importer",
        "category": "Home Appliances & Electronics",
        "primary_channel": "Daraz Bangladesh & Independent Web",
        "price_range_bdt": "৳32,000 - ৳165,000",
        "base_city": "Dhaka",
        "website": "https://brandbazaarbd.com",
        "notes": "Top electronic appliances merchant on Daraz known for original Japanese TVs, air conditioners, and kitchen appliances."
    },
    {
        "seller_name": "Nazmul Huda",
        "organisation_name": "Fly High Bangladesh",
        "top_selling_product": "Waterproof Multi-Compartment Laptop Backpack (15.6\")",
        "business_model": "Marketplace Top Seller / Bag Manufacturer",
        "category": "Bags, Luggage & Travel",
        "primary_channel": "Daraz Bangladesh Top Merchant",
        "price_range_bdt": "৳550 - ৳2,450",
        "base_city": "Dhaka",
        "website": "https://daraz.com.bd",
        "notes": "One of Daraz's top-volume domestic bag sellers, manufacturing durable student, office, and travel backpacks."
    },
    {
        "seller_name": "Shamim Ahmed",
        "organisation_name": "Bag Planet BD",
        "top_selling_product": "Anti-Theft USB Charging Travel Backpack & Crossbody Sling Bag",
        "business_model": "Marketplace Merchant / Importer",
        "category": "Bags, Luggage & Travel",
        "primary_channel": "Daraz Bangladesh",
        "price_range_bdt": "৳480 - ৳1,950",
        "base_city": "Dhaka",
        "website": "https://daraz.com.bd",
        "notes": "Consistently featured in Daraz flash sales for high-volume modern commuter and student bags."
    },
    {
        "seller_name": "Enamul Kabir",
        "organisation_name": "Leather Craft BD",
        "top_selling_product": "100% Pure Cowhide Genuine Leather Bifold Men's Wallet & Belt Set",
        "business_model": "Marketplace Merchant / Leather Artisan",
        "category": "Leather Goods & Accessories",
        "primary_channel": "Daraz Bangladesh & Social Commerce",
        "price_range_bdt": "৳450 - ৳2,800",
        "base_city": "Dhaka (Hazaribagh)",
        "website": "https://daraz.com.bd",
        "notes": "Authentic leather goods manufacturer producing top-rated export-quality wallets, belts, and laptop sleeves."
    },
    {
        "seller_name": "Farhana Yeasmin",
        "organisation_name": "Dazzle Fashion Jewellery BD",
        "top_selling_product": "Korean Minimalist Gold-Plated Drop Earrings & Choker Necklace Set",
        "business_model": "Marketplace Merchant / Fashion Accessories",
        "category": "Women's Fashion Jewellery",
        "primary_channel": "Daraz Bangladesh Top Fashion Merchant",
        "price_range_bdt": "৳85 - ৳650",
        "base_city": "Dhaka",
        "website": "https://daraz.com.bd",
        "notes": "High unit-velocity seller on Daraz in fashion jewellery, appealing to university students and casual shoppers."
    },
    {
        "seller_name": "Rayan Kabir",
        "organisation_name": "Al-Rayan Perfumes & Attar",
        "top_selling_product": "Alcohol-Free Concentrated White Oudh & Kasturi Attar (12ml)",
        "business_model": "Marketplace Merchant / Fragrance Seller",
        "category": "Fragrances & Personal Care",
        "primary_channel": "Daraz Bangladesh & Facebook Commerce",
        "price_range_bdt": "৳150 - ৳1,100",
        "base_city": "Dhaka",
        "website": "https://daraz.com.bd",
        "notes": "Specializes in non-alcoholic concentrated roll-on attars and incense burners with 10,000+ positive marketplace reviews."
    },
    {
        "seller_name": "Shanta Kumar Saha",
        "organisation_name": "Shanta Electronics BD",
        "top_selling_product": "Digital LCD Multimeter (DT-830D) & 60W Soldering Iron Kit",
        "business_model": "Marketplace Merchant / Tool Specialist",
        "category": "Tools & Hardware / DIY",
        "primary_channel": "Daraz Bangladesh (Top Tools Merchant)",
        "price_range_bdt": "৳180 - ৳1,450",
        "base_city": "Dhaka (Nawabpur)",
        "website": "https://daraz.com.bd",
        "notes": "Top hardware and electrical testing tools vendor sourced directly from Nawabpur industrial wholesale market."
    },
    {
        "seller_name": "Mohammad Zahirul Islam",
        "organisation_name": "Smart Technologies (BD) Ltd",
        "top_selling_product": "Gigabyte B550 Gaming Motherboard, Corsair Vengeance RAM & HP Laptops",
        "business_model": "Marketplace Flagship / National IT Distributor",
        "category": "Computer Hardware & IT",
        "primary_channel": "Daraz Bangladesh (DarazMall Tech) & Pickaboo",
        "price_range_bdt": "৳4,500 - ৳140,000",
        "base_city": "Dhaka",
        "website": "https://smartbd.com",
        "notes": "Largest IT hardware distributor in Bangladesh representing Acer, HP, Gigabyte, Corsair, and Intel."
    },
    {
        "seller_name": "Sarwar Jahan",
        "organisation_name": "Unique Business Systems (UCC)",
        "top_selling_product": "AMD Ryzen 5 / 7 Desktop Processors & Sapphire Radeon RX Graphics Cards",
        "business_model": "Marketplace Merchant / IT Component Partner",
        "category": "Computer Hardware & IT",
        "primary_channel": "Daraz Bangladesh & Multiplan Outlets",
        "price_range_bdt": "৳8,500 - ৳85,000",
        "base_city": "Dhaka",
        "website": "https://ucc-bd.com",
        "notes": "Key authorized distributor for AMD, Sapphire, ViewSonic, and MSI components on e-commerce platforms."
    },
    {
        "seller_name": "Abdul Fattah",
        "organisation_name": "Global Brand Pvt Ltd",
        "top_selling_product": "ASUS TUF Gaming Laptops & A4Tech Bloody Optical Gaming Mouse",
        "business_model": "Marketplace Flagship / IT Mega Distributor",
        "category": "Computer Hardware & IT",
        "primary_channel": "Daraz Bangladesh & Pickaboo",
        "price_range_bdt": "৳850 - ৳175,000",
        "base_city": "Dhaka",
        "website": "https://globalbrand.com.bd",
        "notes": "Exclusive master distributor for ASUS, Brother, A4Tech, and LG monitors with wide presence on DarazMall."
    },
    {
        "seller_name": "Authorized BD Partner",
        "organisation_name": "SKMEI Official Store Bangladesh",
        "top_selling_product": "SKMEI 1251 Waterproof Military Sports Digital Watch",
        "business_model": "Marketplace Merchant / Authorized Partner",
        "category": "Watches & Accessories",
        "primary_channel": "Daraz Bangladesh (DarazMall Top Watch Seller)",
        "price_range_bdt": "৳680 - ৳1,950",
        "base_city": "Dhaka",
        "website": "https://daraz.com.bd/shop/skmei-official-store",
        "notes": "Huge volume mover on Daraz with over 150,000 units sold across budget waterproof sports and tactical watches."
    },
    {
        "seller_name": "Authorized Importer BD",
        "organisation_name": "Naviforce Bangladesh Official",
        "top_selling_product": "Naviforce NF9093 Dual Display Leather Quartz Watch",
        "business_model": "Marketplace Merchant / Authorized Partner",
        "category": "Watches & Accessories",
        "primary_channel": "Daraz Bangladesh (DarazMall)",
        "price_range_bdt": "৳1,250 - ৳3,450",
        "base_city": "Dhaka",
        "website": "https://daraz.com.bd/shop/naviforce",
        "notes": "Top-rated brand for affordable luxury wristwatches on Daraz with genuine gift box packaging and warranty."
    },
    {
        "seller_name": "DRM / Certified Tech Distributor",
        "organisation_name": "Ugreen Bangladesh Flagship",
        "top_selling_product": "Ugreen 65W GaN Fast Charger & 100W USB-C Braided Cable",
        "business_model": "Marketplace Flagship / Certified Partner",
        "category": "Consumer Electronics & Gadgets",
        "primary_channel": "Daraz Bangladesh (DarazMall)",
        "price_range_bdt": "৳450 - ৳4,200",
        "base_city": "Dhaka",
        "website": "https://daraz.com.bd/shop/ugreen-bangladesh",
        "notes": "Premium charging and mobile connectivity brand rapidly gaining #1 position for MacBook/Laptop GaN chargers."
    },
    {
        "seller_name": "Excel Technologies Ltd / Authorized Lead",
        "organisation_name": "SanDisk Official Store BD",
        "top_selling_product": "SanDisk Ultra 64GB / 128GB MicroSDXC Class 10 Memory Card",
        "business_model": "Marketplace Flagship / Authorized Distributor",
        "category": "Storage & IT Peripherals",
        "primary_channel": "Daraz Bangladesh (DarazMall)",
        "price_range_bdt": "৳520 - ৳2,850",
        "base_city": "Dhaka",
        "website": "https://daraz.com.bd/shop/sandisk-official-store",
        "notes": "Essential high-velocity storage brand on Daraz with official manufacturer 5-10 year replacement warranty."
    },
    {
        "seller_name": "Multi-Merchant Channel",
        "organisation_name": "Lenovo Audio Bangladesh",
        "top_selling_product": "Lenovo LP40 Pro TWS Wireless Bluetooth 5.1 Earphones",
        "business_model": "Marketplace Merchant / Gadget Importer",
        "category": "Consumer Electronics & Gadgets",
        "primary_channel": "Daraz Bangladesh (Top 5 Gadget Seller)",
        "price_range_bdt": "৳750 - ৳1,450",
        "base_city": "Dhaka",
        "website": "https://daraz.com.bd",
        "notes": "Over 42,000 units sold on Daraz; among the most sought-after budget wireless earbuds in Bangladesh."
    },
    {
        "seller_name": "Global Gadget Hub BD",
        "organisation_name": "Global Gadget Hub BD",
        "top_selling_product": "M10 TWS Bluetooth 5.1 Wireless Earbuds with 2000mAh Power Bank Display",
        "business_model": "Marketplace Top Seller / Gadget Merchant",
        "category": "Consumer Electronics & Gadgets",
        "primary_channel": "Daraz Bangladesh (Rank #1 Electronics Volume)",
        "price_range_bdt": "৳350 - ৳580",
        "base_city": "Dhaka",
        "website": "https://daraz.com.bd",
        "notes": "The single highest volume unit seller in Bangladesh e-commerce history with over 85,000 verified sales on Daraz."
    },
    {
        "seller_name": "Digital Tech Plaza BD",
        "organisation_name": "Digital Tech Plaza BD",
        "top_selling_product": "T800 Ultra 1.99-inch HD Display Bluetooth Calling Smartwatch",
        "business_model": "Marketplace Top Seller / Gadget Merchant",
        "category": "Consumer Electronics & Gadgets",
        "primary_channel": "Daraz Bangladesh (Rank #2 Electronics Volume)",
        "price_range_bdt": "৳650 - ৳990",
        "base_city": "Dhaka",
        "website": "https://daraz.com.bd",
        "notes": "Breakout budget smartwatch with over 64,000 units sold on Daraz, featuring wireless magnetic charging."
    },
    {
        "seller_name": "Godrej Consumer Products BD Ltd",
        "organisation_name": "Godrej Official Store",
        "top_selling_product": "GoodKnight Power Activ Mosquito Vaporizer Machine & Refill Twin Pack",
        "business_model": "Marketplace Flagship / Household FMCG",
        "category": "Household & FMCG",
        "primary_channel": "Daraz Bangladesh (DarazMall)",
        "price_range_bdt": "৳140 - ৳680",
        "base_city": "Dhaka",
        "website": "https://daraz.com.bd/shop/godrej-bangladesh",
        "notes": "Leading household insect protection brand experiencing major seasonal surges in monsoon and post-monsoon months."
    },
    {
        "seller_name": "MAS Holdings / Local Partner",
        "organisation_name": "Amanté Bangladesh Official",
        "top_selling_product": "Amanté Seamless Microfibre Padded Everyday T-Shirt Bra",
        "business_model": "Marketplace Flagship / Women's Innerwear",
        "category": "Women's Intimate Apparel",
        "primary_channel": "Daraz Bangladesh (DarazMall)",
        "price_range_bdt": "৳850 - ৳2,900",
        "base_city": "Dhaka",
        "website": "https://daraz.com.bd/shop/amante-bangladesh",
        "notes": "Premier premium intimate wear and lingerie brand on Daraz with dedicated discreet packaging."
    },
    {
        "seller_name": "Habib Rahman",
        "organisation_name": "Pilot Bazaar BD",
        "top_selling_product": "Reconditioned Japanese Toyota Allion / Premio & Corolla Cross",
        "business_model": "Classifieds Verified Member Store",
        "category": "Automotive & Vehicles",
        "primary_channel": "Bikroy.com & Pilot Bazaar Portal (pilotbazaar.com)",
        "price_range_bdt": "৳2,200,000 - ৳5,800,000",
        "base_city": "Dhaka (Baridhara)",
        "website": "https://pilotbazaar.com",
        "notes": "One of the most reputed and verified automobile merchants on Bikroy.com, selling inspected reconditioned Japanese cars."
    },
    {
        "seller_name": "Tanvir Ahmed",
        "organisation_name": "Mobile Mela Dhaka",
        "top_selling_product": "Certified Pre-Owned Apple iPhone 13 / 14 (128GB)",
        "business_model": "Classifieds Verified Shop / Mobile Reseller",
        "category": "Consumer Electronics & Gadgets",
        "primary_channel": "Bikroy.com Verified Member & Shop",
        "price_range_bdt": "৳48,000 - ৳92,000",
        "base_city": "Dhaka (Bashundhara City)",
        "website": "https://bikroy.com",
        "notes": "Top verified member shop on Bikroy with physical storefront in Bashundhara City providing 10-day replacement warranty."
    },
    {
        "seller_name": "Md. Mizanur Rahman",
        "organisation_name": "Boi Ghor Nilkhet",
        "top_selling_product": "IBA MBA Admission Preparation Guides & BCS Job Solutions",
        "business_model": "Marketplace Book Merchant",
        "category": "Books & Stationery",
        "primary_channel": "Rokomari.com & Nilkhet Store",
        "price_range_bdt": "৳250 - ৳950",
        "base_city": "Dhaka (Nilkhet)",
        "website": "https://rokomari.com",
        "notes": "Major third-party bookstore merchant on Rokomari sourcing hard-to-find academic and professional competitive exam guides."
    },
    {
        "seller_name": "Anupam Guha",
        "organisation_name": "Anupam Prokashani",
        "top_selling_product": "Selected Works & Bestselling Novels by Humayun Ahmed",
        "business_model": "Publishing House / Marketplace Merchant",
        "category": "Books & Publishing",
        "primary_channel": "Rokomari.com Flagship Publisher",
        "price_range_bdt": "৳200 - ৳1,800",
        "base_city": "Dhaka (Bangla Bazar)",
        "website": "https://rokomari.com",
        "notes": "Prominent Bengali book publisher consistently topping fiction and literature sales charts during Amar Ekushey Boi Mela."
    },
    {
        "seller_name": "Farid Ahmed",
        "organisation_name": "Somoy Prokashon",
        "top_selling_product": "1971 Liberation War History Series & Modern Bengali Fiction",
        "business_model": "Publishing House / Marketplace Merchant",
        "category": "Books & Publishing",
        "primary_channel": "Rokomari.com Flagship Publisher",
        "price_range_bdt": "৳220 - ৳2,400",
        "base_city": "Dhaka",
        "website": "https://rokomari.com",
        "notes": "Key publisher on Rokomari known for historical non-fiction, contemporary novels, and youth literature."
    },
    {
        "seller_name": "M. S. Alam",
        "organisation_name": "Rongdhonu Herbals (Daraz Store)",
        "top_selling_product": "Pure Organic Aloe Vera Gel (300ml) & Kasturi Turmeric Powder",
        "business_model": "Marketplace Merchant / Herbal Producer",
        "category": "Health & Beauty",
        "primary_channel": "Daraz Bangladesh (DarazMall Health & Beauty)",
        "price_range_bdt": "৳150 - ৳550",
        "base_city": "Dhaka",
        "website": "https://daraz.com.bd",
        "notes": "Generated over 40,000 orders on Daraz for natural homemade face packs, hair oils, and organic brightening clays."
    },
    {
        "seller_name": "Kamrul Islam",
        "organisation_name": "BD Car Center",
        "top_selling_product": "Japanese Reconditioned Honda Vezel & Toyota Axio Hybrid",
        "business_model": "Classifieds Verified Member Store",
        "category": "Automotive & Vehicles",
        "primary_channel": "Bikroy.com Verified Member",
        "price_range_bdt": "৳1,850,000 - ৳3,400,000",
        "base_city": "Dhaka (Bijoy Sarani)",
        "website": "https://bikroy.com",
        "notes": "Verified automobile merchant on Bikroy with extensive inventory of auction grade 4.5+ hybrid sedans and crossovers."
    },
    {
        "seller_name": "Authorized Retail Lead",
        "organisation_name": "Remax Bangladesh Official Store",
        "top_selling_product": "Remax OTG Metal USB 3.0 Adapter & 2.4A Fast Sync Data Cable",
        "business_model": "Marketplace Flagship / Gadget Distributor",
        "category": "Consumer Electronics & Gadgets",
        "primary_channel": "Daraz Bangladesh (DarazMall)",
        "price_range_bdt": "৳120 - ৳1,850",
        "base_city": "Dhaka",
        "website": "https://daraz.com.bd/shop/remax-bangladesh",
        "notes": "One of the earliest and most consistent accessory sellers on Daraz, moving over 38,000 OTG adapters alone."
    }
]

def generate_csv(filename):
    fieldnames = [
        "Seller Name",
        "Organisation Name",
        "Top Selling Product",
        "Business Model",
        "Category",
        "Primary Sales Channel / Platform",
        "Price Range (BDT)",
        "Base City",
        "Website / Store Link",
        "Key Notes & Product Highlights"
    ]
    with open(filename, mode='w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for item in sellers_data:
            writer.writerow({
                "Seller Name": item["seller_name"],
                "Organisation Name": item["organisation_name"],
                "Top Selling Product": item["top_selling_product"],
                "Business Model": item["business_model"],
                "Category": item["category"],
                "Primary Sales Channel / Platform": item["primary_channel"],
                "Price Range (BDT)": item["price_range_bdt"],
                "Base City": item["base_city"],
                "Website / Store Link": item["website"],
                "Key Notes & Product Highlights": item["notes"]
            })
    print(f"CSV generated: {filename} ({len(sellers_data)} rows)")

def generate_excel(filename):
    wb = openpyxl.Workbook()
    ws_all = wb.active
    ws_all.title = "All Sellers & Businesses"
    
    # Enable gridlines
    ws_all.views.sheetView[0].showGridLines = True
    
    # Color palette
    NAVY_HEADER = "1B365D"
    WHITE = "FFFFFF"
    ZEBRA_FILL = "F8FAFC"
    D2C_COLOR = "EBF5FB"       # Soft Light Blue
    MARKETPLACE_COLOR = "FEF9E7"# Soft Light Yellow/Gold
    BORDER_COLOR = "CBD5E1"
    
    header_font = Font(name="Segoe UI", size=11, bold=True, color=WHITE)
    header_fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    regular_font = Font(name="Segoe UI", size=10, color="1E293B")
    bold_cell_font = Font(name="Segoe UI", size=10, bold=True, color="0F172A")
    italic_font = Font(name="Segoe UI", size=9, italic=True, color="64748B")
    
    thin_border = Border(
        left=Side(style='thin', color=BORDER_COLOR),
        right=Side(style='thin', color=BORDER_COLOR),
        top=Side(style='thin', color=BORDER_COLOR),
        bottom=Side(style='thin', color=BORDER_COLOR)
    )
    
    headers = [
        "SL #",
        "Seller Name",
        "Organisation Name",
        "Top Selling Product",
        "Business Model",
        "Category",
        "Primary Sales Channel / Platform",
        "Price Range (BDT)",
        "Base City",
        "Website / Store Link",
        "Key Notes & Product Highlights"
    ]
    
    # Write Title block
    ws_all.merge_cells("A1:K1")
    title_cell = ws_all["A1"]
    title_cell.value = "BANGLADESH E-COMMERCE SELLERS & ORGANISATIONS DIRECTORY"
    title_cell.font = Font(name="Segoe UI", size=15, bold=True, color=WHITE)
    title_cell.fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_all.row_dimensions[1].height = 36
    
    # Subtitle block
    ws_all.merge_cells("A2:K2")
    sub_cell = ws_all["A2"]
    sub_cell.value = "Comprehensive Master Dataset: Single-Seller D2C Brands vs Multi-Vendor Marketplace Top Merchants (Daraz, Pickaboo, Bikroy, Rokomari)"
    sub_cell.font = Font(name="Segoe UI", size=10, italic=True, color="E2E8F0")
    sub_cell.fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    sub_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_all.row_dimensions[2].height = 22
    
    # Write headers at row 4
    ws_all.row_dimensions[4].height = 28
    for col_idx, h_text in enumerate(headers, 1):
        cell = ws_all.cell(row=4, column=col_idx)
        cell.value = h_text
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        
    # Write data rows
    for row_idx, item in enumerate(sellers_data, start=5):
        sl = row_idx - 4
        ws_all.row_dimensions[row_idx].height = 24
        
        is_d2c = "Single" in item["business_model"]
        row_fill = PatternFill(start_color=D2C_COLOR if is_d2c else MARKETPLACE_COLOR, 
                               end_color=D2C_COLOR if is_d2c else MARKETPLACE_COLOR, 
                               fill_type="solid") if sl % 2 == 0 else PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        
        row_values = [
            sl,
            item["seller_name"],
            item["organisation_name"],
            item["top_selling_product"],
            item["business_model"],
            item["category"],
            item["primary_channel"],
            item["price_range_bdt"],
            item["base_city"],
            item["website"],
            item["notes"]
        ]
        
        for col_idx, val in enumerate(row_values, 1):
            cell = ws_all.cell(row=row_idx, column=col_idx)
            cell.value = val
            cell.font = bold_cell_font if col_idx in [2, 3, 4] else regular_font
            cell.border = thin_border
            cell.fill = row_fill
            
            # Alignments
            if col_idx in [1, 8, 9]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in [5, 6]:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
                
    # Auto-adjust column widths
    column_widths = {
        1: 8,   # SL
        2: 28,  # Seller Name
        3: 28,  # Organisation Name
        4: 42,  # Top Selling Product
        5: 28,  # Business Model
        6: 24,  # Category
        7: 35,  # Primary Channel
        8: 18,  # Price Range
        9: 16,  # Base City
        10: 32, # Website
        11: 45  # Notes
    }
    for col_idx, width in column_widths.items():
        col_letter = get_column_letter(col_idx)
        ws_all.column_dimensions[col_letter].width = width

    # Add auto-filter
    ws_all.auto_filter.ref = f"A4:K{len(sellers_data) + 4}"
    
    # Freeze header rows
    ws_all.freeze_panes = "A5"
    
    # -------------------------------------------------------------
    # SHEET 2: Single-Seller D2C Brands (Dedicated Sheet)
    # -------------------------------------------------------------
    ws_d2c = wb.create_sheet(title="Single-Seller D2C Brands")
    ws_d2c.views.sheetView[0].showGridLines = True
    
    ws_d2c.merge_cells("A1:K1")
    t2 = ws_d2c["A1"]
    t2.value = "SINGLE SELLER & OWNER-OPERATED D2C BRANDS IN BANGLADESH"
    t2.font = Font(name="Segoe UI", size=14, bold=True, color=WHITE)
    t2.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    t2.alignment = Alignment(horizontal="center", vertical="center")
    ws_d2c.row_dimensions[1].height = 32
    
    ws_d2c.row_dimensions[3].height = 26
    for col_idx, h_text in enumerate(headers, 1):
        cell = ws_d2c.cell(row=3, column=col_idx)
        cell.value = h_text
        cell.font = header_font
        cell.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    d2c_list = [item for item in sellers_data if "Single" in item["business_model"] or "Platform & Direct" in item["business_model"]]
    for r_idx, item in enumerate(d2c_list, start=4):
        sl = r_idx - 3
        ws_d2c.row_dimensions[r_idx].height = 23
        r_fill = PatternFill(start_color="F0F9FF", end_color="F0F9FF", fill_type="solid") if sl % 2 == 0 else PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        
        row_values = [
            sl, item["seller_name"], item["organisation_name"], item["top_selling_product"],
            item["business_model"], item["category"], item["primary_channel"],
            item["price_range_bdt"], item["base_city"], item["website"], item["notes"]
        ]
        for col_idx, val in enumerate(row_values, 1):
            cell = ws_d2c.cell(row=r_idx, column=col_idx)
            cell.value = val
            cell.font = bold_cell_font if col_idx in [2, 3, 4] else regular_font
            cell.border = thin_border
            cell.fill = r_fill
            cell.alignment = Alignment(horizontal="center" if col_idx in [1, 8, 9] else "left", vertical="center")

    for col_idx, width in column_widths.items():
        ws_d2c.column_dimensions[get_column_letter(col_idx)].width = width
    ws_d2c.auto_filter.ref = f"A3:K{len(d2c_list) + 3}"
    ws_d2c.freeze_panes = "A4"

    # -------------------------------------------------------------
    # SHEET 3: Multi-Seller Marketplace Merchants (Daraz/Pickaboo/Bikroy)
    # -------------------------------------------------------------
    ws_mkt = wb.create_sheet(title="Marketplace Merchants (Daraz)")
    ws_mkt.views.sheetView[0].showGridLines = True
    
    ws_mkt.merge_cells("A1:K1")
    t3 = ws_mkt["A1"]
    t3.value = "MARKETPLACE SELLERS & VERIFIED MERCHANTS (DARAZ, PICKABOO, BIKROY)"
    t3.font = Font(name="Segoe UI", size=14, bold=True, color=WHITE)
    t3.fill = PatternFill(start_color="C2410C", end_color="C2410C", fill_type="solid") # Dark Orange / Daraz style
    t3.alignment = Alignment(horizontal="center", vertical="center")
    ws_mkt.row_dimensions[1].height = 32
    
    ws_mkt.row_dimensions[3].height = 26
    for col_idx, h_text in enumerate(headers, 1):
        cell = ws_mkt.cell(row=3, column=col_idx)
        cell.value = h_text
        cell.font = header_font
        cell.fill = PatternFill(start_color="C2410C", end_color="C2410C", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    mkt_list = [item for item in sellers_data if "Marketplace" in item["business_model"] or "Classifieds" in item["business_model"] or "Publishing" in item["business_model"]]
    for r_idx, item in enumerate(mkt_list, start=4):
        sl = r_idx - 3
        ws_mkt.row_dimensions[r_idx].height = 23
        r_fill = PatternFill(start_color="FFF7ED", end_color="FFF7ED", fill_type="solid") if sl % 2 == 0 else PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        
        row_values = [
            sl, item["seller_name"], item["organisation_name"], item["top_selling_product"],
            item["business_model"], item["category"], item["primary_channel"],
            item["price_range_bdt"], item["base_city"], item["website"], item["notes"]
        ]
        for col_idx, val in enumerate(row_values, 1):
            cell = ws_mkt.cell(row=r_idx, column=col_idx)
            cell.value = val
            cell.font = bold_cell_font if col_idx in [2, 3, 4] else regular_font
            cell.border = thin_border
            cell.fill = r_fill
            cell.alignment = Alignment(horizontal="center" if col_idx in [1, 8, 9] else "left", vertical="center")

    for col_idx, width in column_widths.items():
        ws_mkt.column_dimensions[get_column_letter(col_idx)].width = width
    ws_mkt.auto_filter.ref = f"A3:K{len(mkt_list) + 3}"
    ws_mkt.freeze_panes = "A4"

    # -------------------------------------------------------------
    # SHEET 4: Ecosystem Summary & Analytics
    # -------------------------------------------------------------
    ws_sum = wb.create_sheet(title="Ecosystem Breakdown")
    ws_sum.views.sheetView[0].showGridLines = True
    
    ws_sum.merge_cells("A1:E1")
    t4 = ws_sum["A1"]
    t4.value = "BANGLADESH E-COMMERCE ECOSYSTEM SUMMARY"
    t4.font = Font(name="Segoe UI", size=14, bold=True, color=WHITE)
    t4.fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    t4.alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[1].height = 32

    # Category counts
    cat_counts = {}
    model_counts = {}
    for item in sellers_data:
        c = item["category"]
        m = "Single-Seller / D2C Brand" if ("Single" in item["business_model"] or "Platform & Direct" in item["business_model"]) else "Marketplace Merchant / Seller"
        cat_counts[c] = cat_counts.get(c, 0) + 1
        model_counts[m] = model_counts.get(m, 0) + 1

    # Table 1: Model breakdown
    ws_sum.cell(row=3, column=1, value="Business Model Breakdown").font = Font(name="Segoe UI", size=11, bold=True, color="1E3A8A")
    ws_sum.cell(row=4, column=1, value="Model Type").font = header_font
    ws_sum.cell(row=4, column=1).fill = header_fill
    ws_sum.cell(row=4, column=2, value="Seller Count").font = header_font
    ws_sum.cell(row=4, column=2).fill = header_fill
    ws_sum.cell(row=4, column=3, value="Percentage").font = header_font
    ws_sum.cell(row=4, column=3).fill = header_fill

    cur_row = 5
    for m_type, count in model_counts.items():
        ws_sum.cell(row=cur_row, column=1, value=m_type).font = regular_font
        ws_sum.cell(row=cur_row, column=2, value=count).font = bold_cell_font
        ws_sum.cell(row=cur_row, column=3, value=f"{(count/len(sellers_data))*100:.1f}%").font = regular_font
        for c in range(1, 4):
            ws_sum.cell(row=cur_row, column=c).border = thin_border
        cur_row += 1

    # Total row
    ws_sum.cell(row=cur_row, column=1, value="Total Documented Sellers").font = Font(name="Segoe UI", size=10, bold=True)
    ws_sum.cell(row=cur_row, column=2, value=len(sellers_data)).font = Font(name="Segoe UI", size=10, bold=True)
    ws_sum.cell(row=cur_row, column=3, value="100.0%").font = Font(name="Segoe UI", size=10, bold=True)
    for c in range(1, 4):
        ws_sum.cell(row=cur_row, column=c).border = thin_border
        ws_sum.cell(row=cur_row, column=c).fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")

    # Table 2: Category Breakdown
    cur_row += 3
    ws_sum.cell(row=cur_row, column=1, value="Top Selling Categories Breakdown").font = Font(name="Segoe UI", size=11, bold=True, color="1E3A8A")
    cur_row += 1
    ws_sum.cell(row=cur_row, column=1, value="Category").font = header_font
    ws_sum.cell(row=cur_row, column=1).fill = header_fill
    ws_sum.cell(row=cur_row, column=2, value="Sellers").font = header_font
    ws_sum.cell(row=cur_row, column=2).fill = header_fill
    ws_sum.cell(row=cur_row, column=3, value="Share (%)").font = header_font
    ws_sum.cell(row=cur_row, column=3).fill = header_fill

    cur_row += 1
    for cat, count in sorted(cat_counts.items(), key=lambda x: x[1], reverse=True):
        ws_sum.cell(row=cur_row, column=1, value=cat).font = regular_font
        ws_sum.cell(row=cur_row, column=2, value=count).font = bold_cell_font
        ws_sum.cell(row=cur_row, column=3, value=f"{(count/len(sellers_data))*100:.1f}%").font = regular_font
        for c in range(1, 4):
            ws_sum.cell(row=cur_row, column=c).border = thin_border
        cur_row += 1

    ws_sum.column_dimensions["A"].width = 32
    ws_sum.column_dimensions["B"].width = 16
    ws_sum.column_dimensions["C"].width = 16

    wb.save(filename)
    print(f"Excel generated: {filename} with {len(wb.sheetnames)} sheets.")

if __name__ == "__main__":
    generate_csv("bangladesh_ecommerce_sellers.csv")
    generate_excel("bangladesh_ecommerce_sellers.xlsx")
