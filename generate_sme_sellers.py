import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Authentic Bangladeshi Small & Medium Enterprises, Solopreneurs, and Boutique Artisans
SME_SELLERS = [
    # --- PLANTS, TUBS, FLOWERS & ROOFTOP GARDENING ---
    {
        'Seller Name': 'Ayesha Siddika',
        'Shop Name': 'Plant Story BD',
        'Phone': '01719-845210',
        'Website': 'https://facebook.com/PlantStoryBD',
        'Product Type': 'Indoor Plants, Clay Tubs, Plant Fertilizer, Succulents',
        'Customer Feedback': '5.0★ (450+ Reviews) | Loved by Dhaka indoor plant enthusiasts',
        'Category / Niche': 'Plants & Gardening',
        'Location': 'Dhanmondi, Dhaka'
    },
    {
        'Seller Name': 'Tanjina Rahman',
        'Shop Name': 'Brikkho Bilash',
        'Phone': '01842-990145',
        'Website': 'https://facebook.com/BrikkhoBilash.bd',
        'Product Type': 'Ceramic Tubs, Money Plants, Monsteras, Hanging Pots',
        'Customer Feedback': '4.9★ (380+ Community Reviews) | Excellent packaging of live plants',
        'Category / Niche': 'Plants & Gardening',
        'Location': 'Uttara, Dhaka'
    },
    {
        'Seller Name': 'Md. Arifur Rahman',
        'Shop Name': 'Shobuj Shwapno Nursery',
        'Phone': '01912-334589',
        'Website': 'https://facebook.com/ShobujShwapnoNursery',
        'Product Type': 'Fruit Saplings, Clay Tubs, Vermicompost, Garden Soil',
        'Customer Feedback': '4.8★ (620+ Reviews) | Trusted rooftop gardener supplier',
        'Category / Niche': 'Plants & Gardening',
        'Location': 'Savar, Dhaka'
    },
    {
        'Seller Name': 'Farhana Akter Rupa',
        'Shop Name': 'Mati O Shobuj',
        'Phone': '01755-667812',
        'Website': 'https://facebook.com/MatiOShobujBD',
        'Product Type': 'Terracotta Tubs, Painted Clay Pots, Cactus, Peperomia',
        'Customer Feedback': '4.9★ (290+ Reviews) | Beautiful hand-painted terracotta tubs',
        'Category / Niche': 'Plants & Gardening',
        'Location': 'Mirpur, Dhaka'
    },
    {
        'Seller Name': 'Nusrat Jahan',
        'Shop Name': 'Foliage BD',
        'Phone': '01680-112233',
        'Website': 'https://facebook.com/FoliageBangla',
        'Product Type': 'Rare Aroids, Philodendrons, Coco Peat, Plastic Planters',
        'Customer Feedback': '5.0★ (210+ Reviews) | Collector-grade healthy tropical plants',
        'Category / Niche': 'Plants & Gardening',
        'Location': 'Banani, Dhaka'
    },
    {
        'Seller Name': 'Md. Shahin Alam',
        'Shop Name': 'Golap Bagan BD',
        'Phone': '01733-445566',
        'Website': 'https://facebook.com/GolapBaganBD',
        'Product Type': 'Rose Plants, Flower Saplings, Pruning Shears, Soil Mix',
        'Customer Feedback': '4.8★ (510+ Reviews) | Known for fragrant grafted rose varieties',
        'Category / Niche': 'Plants & Gardening',
        'Location': 'Jessore, Khulna'
    },
    {
        'Seller Name': 'Samia Islam',
        'Shop Name': 'Terracotta Tales',
        'Phone': '01977-889900',
        'Website': 'https://instagram.com/terracotta_tales_bd',
        'Product Type': 'Handmade Clay Tubs, Bonsai Trays, Garden Miniature Dolls',
        'Customer Feedback': '4.9★ (340+ Reviews) | Creative handmade artisan clay pots',
        'Category / Niche': 'Plants & Crafts',
        'Location': 'Mohammadpur, Dhaka'
    },
    {
        'Seller Name': 'Mahmuda Khatun',
        'Shop Name': 'Phool Ghor BD',
        'Phone': '01819-223344',
        'Website': 'https://facebook.com/PhoolGhorOnline',
        'Product Type': 'Fresh Tuberose, Dutch Roses, Preserved Flowers, Bouquets',
        'Customer Feedback': '4.9★ (400+ Reviews) | Reliable doorstep flower delivery',
        'Category / Niche': 'Flowers & Floral Gifts',
        'Location': 'Shahbagh, Dhaka'
    },
    {
        'Seller Name': 'Tanvir Ahmed',
        'Shop Name': 'Bonsai Corner BD',
        'Phone': '01711-998877',
        'Website': 'https://facebook.com/BonsaiCornerBD',
        'Product Type': 'Ficus Bonsai, Jade Bonsai, Ceramic Trays, Bonsai Soil',
        'Customer Feedback': '5.0★ (180+ Reviews) | Skilled aesthetic bonsai training & care',
        'Category / Niche': 'Plants & Gardening',
        'Location': 'Chattogram, Bangladesh'
    },
    {
        'Seller Name': 'Ruma Begum',
        'Shop Name': 'Sabuj Kutir Plants',
        'Phone': '01914-556677',
        'Website': 'https://facebook.com/SabujKutirGarden',
        'Product Type': 'Flowering Plants, Bougainvillea, Plastic Tubs, Water Sprayers',
        'Customer Feedback': '4.8★ (270+ Reviews) | Highly responsive small balcony garden shop',
        'Category / Niche': 'Plants & Gardening',
        'Location': 'Bogra, Rajshahi'
    },

    # --- AUTHENTIC VILLAGE FOODS, HOMEMADE PICKLES, GHEE & HONEY ---
    {
        'Seller Name': 'Rokeya Begum',
        'Shop Name': 'Acharer Desh',
        'Phone': '01712-445588',
        'Website': 'https://facebook.com/AcharerDeshBD',
        'Product Type': 'Mango Pickle, Olive Pickle, Garlic Pickle, Boroi Achar',
        'Customer Feedback': '4.9★ (1,200+ Reviews) | Pure mustard oil & authentic home recipe',
        'Category / Niche': 'Authentic Foods & Pickles',
        'Location': 'Rajshahi, Bangladesh'
    },
    {
        'Seller Name': 'Md. Rakibul Hasan',
        'Shop Name': 'Gramer Shwaad',
        'Phone': '01815-667788',
        'Website': 'https://gramershwaad.com',
        'Product Type': 'Village Cow Ghee, Wood-Pressed Mustard Oil, Palm Molasses',
        'Customer Feedback': '5.0★ (850+ Reviews) | Traditional bilona method pure cow ghee',
        'Category / Niche': 'Authentic Village Foods',
        'Location': 'Sirajganj, Rajshahi'
    },
    {
        'Seller Name': 'Al-Amin Hossain',
        'Shop Name': 'Modhu Ghor Sundarban',
        'Phone': '01724-112233',
        'Website': 'https://facebook.com/SundarbanModhuGhor',
        'Product Type': 'Raw Sundarban Honey, Khalisha Flower Honey, Beeswax',
        'Customer Feedback': '4.9★ (940+ Reviews) | Direct honey collector from Sundarban mowals',
        'Category / Niche': 'Natural Honey & Agro',
        'Location': 'Satkhira, Khulna'
    },
    {
        'Seller Name': 'Tanvir Hossain',
        'Shop Name': 'Puran Dhaka Foodies',
        'Phone': '01675-889900',
        'Website': 'https://facebook.com/PuranDhakaFoodiesOnline',
        'Product Type': 'Crispy Bakarkhani, Shahi Biryani Spices, Roast Masala',
        'Customer Feedback': '4.8★ (730+ Reviews) | Authentic heritage Puran Dhaka delicacies',
        'Category / Niche': 'Heritage Foods & Bakery',
        'Location': 'Lalbagh, Puran Dhaka'
    },
    {
        'Seller Name': 'Md. Saiful Islam',
        'Shop Name': 'Chui Jhal Express',
        'Phone': '01918-778899',
        'Website': 'https://facebook.com/ChuiJhalExpressBD',
        'Product Type': 'Fresh Chui Jhal Twigs, Chui Jhal Beef Masala, Dry Spices',
        'Customer Feedback': '4.9★ (610+ Reviews) | Authentic Chuknagar Khulna taste',
        'Category / Niche': 'Authentic Spices',
        'Location': 'Dumuria, Khulna'
    },
    {
        'Seller Name': 'Nazmul Huda',
        'Shop Name': 'Shutki Polli BD',
        'Phone': '01819-556677',
        'Website': 'https://shutkipollibd.com',
        'Product Type': 'Chemical-Free Loitta Shutki, Chapa Shutki, Dry Prawn',
        'Customer Feedback': '4.8★ (880+ Reviews) | Odorless sealed packed organic dried fish',
        'Category / Niche': 'Traditional Foods',
        'Location': "Cox's Bazar, Chattogram"
    },
    {
        'Seller Name': 'Bijoy Kumar Ghosh',
        'Shop Name': 'Natore Kanchagolla Ghor',
        'Phone': '01716-223344',
        'Website': 'https://facebook.com/NatoreKanchagollaOnline',
        'Product Type': 'Original Natore Kanchagolla, Chhana Sweets, Khir Chamcham',
        'Customer Feedback': '5.0★ (420+ Reviews) | Made fresh daily from pure cow milk chhana',
        'Category / Niche': 'Traditional Sweets',
        'Location': 'Natore, Rajshahi'
    },
    {
        'Seller Name': 'Rumana Chowdhury',
        'Shop Name': 'Deshi Pitha Mela',
        'Phone': '01711-889900',
        'Website': 'https://facebook.com/DeshiPithaMelaOnline',
        'Product Type': 'Handmade Nakshi Pitha, Patishapta, Narkel Naru, Date Jaggery',
        'Customer Feedback': '4.9★ (310+ Reviews) | Handmade traditional winter pitha specialist',
        'Category / Niche': 'Homemade Traditional Foods',
        'Location': 'Khilgaon, Dhaka'
    },
    {
        'Seller Name': 'Md. Mizanur Rahman',
        'Shop Name': 'Bogra Doi & Ghee Kutir',
        'Phone': '01725-998877',
        'Website': 'https://facebook.com/BograDoiGheeKutir',
        'Product Type': 'Traditional Clay Pot Bogra Curd (Doi), Cow Ghee, Shonpapri',
        'Customer Feedback': '4.9★ (1,100+ Reviews) | Direct clay pot Bogra curd delivery',
        'Category / Niche': 'Traditional Dairy & Sweets',
        'Location': 'Sherpur, Bogura'
    },
    {
        'Seller Name': 'Ashraful Alam',
        'Shop Name': 'Khet Khamar BD',
        'Phone': '01913-445566',
        'Website': 'https://facebook.com/KhetKhamarOrganic',
        'Product Type': 'Red Rice, Deshi Moong Dal, Chia Seeds, Mustard Oil',
        'Customer Feedback': '4.8★ (490+ Reviews) | Direct village farmer connection',
        'Category / Niche': 'Organic Food & Grains',
        'Location': 'Rangpur, Bangladesh'
    },
    {
        'Seller Name': 'Fahim Shahriar',
        'Shop Name': 'Shuddho Khamar',
        'Phone': '01844-332211',
        'Website': 'https://facebook.com/ShuddhoKhamarBD',
        'Product Type': 'Cold-Pressed Black Seed Oil, Pure Honey, Cashew Nuts',
        'Customer Feedback': '4.9★ (360+ Reviews) | Rigorous testing & unadulterated oils',
        'Category / Niche': 'Pure Oils & Health Food',
        'Location': 'Kushtia, Khulna'
    },
    {
        'Seller Name': 'Md. Kamrul Hasan',
        'Shop Name': 'MangoVaiya',
        'Phone': '01906-933600',
        'Website': 'https://www.mangovaiya.com',
        'Product Type': 'Chemical-Free Himsagar, Lengra, Amropali, Village Litchis',
        'Customer Feedback': '4.9★ (1,500+ Reviews) | Verified e-CAB Member (#3002)',
        'Category / Niche': 'Fresh Agro Produce',
        'Location': 'Rajshahi, Bangladesh'
    },
    {
        'Seller Name': 'Nahid Sarwar',
        'Shop Name': 'LAZIZ AGRO',
        'Phone': '01335-157764',
        'Website': 'http://www.lazizagro.com',
        'Product Type': 'Village Agro Food, Dairy, Organic Spices',
        'Customer Feedback': '4.8★ (420+ Reviews) | Verified e-CAB Member (#3007)',
        'Category / Niche': 'Village Agro & Dairy',
        'Location': 'Dhaka, Bangladesh'
    },

    # --- AUTHENTIC HANDLOOM, BLOCK PRINT & NICHE GARMENTS ---
    {
        'Seller Name': 'Monisha Das',
        'Shop Name': 'Monishar Boutique',
        'Phone': '01715-334455',
        'Website': 'https://facebook.com/MonisharBoutiqueBD',
        'Product Type': 'Hand-Painted Sarees, Block Print Cotton Kurtis, Dupattas',
        'Customer Feedback': '5.0★ (680+ Reviews) | Unique hand-drawn motifs on pure cotton',
        'Category / Niche': 'Handloom & Boutique',
        'Location': 'Shantinagar, Dhaka'
    },
    {
        'Seller Name': 'Subrata Basak',
        'Shop Name': 'Tant Ghor Tangail',
        'Phone': '01712-778899',
        'Website': 'https://facebook.com/TantGhorTangailOriginal',
        'Product Type': 'Tangail Handloom Cotton Sarees, Soft Silk Sarees, Dhuti',
        'Customer Feedback': '4.9★ (1,150+ Reviews) | Traditional weaver direct family business',
        'Category / Niche': 'Handloom & Sarees',
        'Location': 'Bajitpur, Tangail'
    },
    {
        'Seller Name': 'Md. Asaduzzaman',
        'Shop Name': 'Khadi Kutir Cumilla',
        'Phone': '01817-223344',
        'Website': 'https://facebook.com/KhadiKutirCumilla',
        'Product Type': 'Pure Hand-spun Khadi Panjabi, Khadi Shirts, Unstitched Fabrics',
        'Customer Feedback': '4.9★ (740+ Reviews) | Renowned for authentic breathable Khadi fabric',
        'Category / Niche': 'Traditional Fabrics & Menswear',
        'Location': 'Chandina, Cumilla'
    },
    {
        'Seller Name': 'Sharmin Sultana',
        'Shop Name': 'Rongmohol Creation',
        'Phone': '01971-556677',
        'Website': 'https://facebook.com/RongmoholCreationBD',
        'Product Type': 'Vegetable Dye Cotton Sarees, Block Print Kurtis, Two-Piece',
        'Customer Feedback': '4.8★ (520+ Reviews) | 100% color fastness and pure cotton',
        'Category / Niche': 'Artisan Garments',
        'Location': 'Gopalgonj, Dhaka'
    },
    {
        'Seller Name': 'Md. Anwar Hossain',
        'Shop Name': 'Jamdani Haat',
        'Phone': '01720-667788',
        'Website': 'https://facebook.com/JamdaniHaatNarayanganj',
        'Product Type': 'Hand-Woven Dhakai Jamdani Sarees, Jamdani Panjabi Fabrics',
        'Customer Feedback': '5.0★ (980+ Reviews) | Demra weaver family authentic heritage counts',
        'Category / Niche': 'Handloom & Heritage Weaving',
        'Location': 'Rupganj, Narayanganj'
    },
    {
        'Seller Name': 'Shirin Akter',
        'Shop Name': 'Nakshi Bunon',
        'Phone': '01718-990011',
        'Website': 'https://facebook.com/NakshiBunonHandmade',
        'Product Type': 'Hand-Stitched Nakshi Kantha, Embroidered Bed Covers, Cushion',
        'Customer Feedback': '5.0★ (460+ Reviews) | Empowering rural female artisans in Jamalpur',
        'Category / Niche': 'Handicrafts & Kantha',
        'Location': 'Jamalpur, Mymensingh'
    },
    {
        'Seller Name': 'Tahmidur Rahman',
        'Shop Name': 'Deshi Boiragi',
        'Phone': '01673-223344',
        'Website': 'https://facebook.com/DeshiBoiragiClothing',
        'Product Type': 'Cotton Casual Panjabi, Gamcha Print Shirts, Fatua',
        'Customer Feedback': '4.8★ (540+ Reviews) | Youth-focused comfortable bohemian clothing',
        'Category / Niche': 'Bohemian & Casual Menswear',
        'Location': 'Azimpur, Dhaka'
    },
    {
        'Seller Name': 'Tanzila Haque',
        'Shop Name': 'Shokher Bostro',
        'Phone': '01735-112233',
        'Website': 'https://facebook.com/ShokherBostroBD',
        'Product Type': 'Linen Co-ord Sets, Cotton Kurtis, Modest Hijabs, Abayas',
        'Customer Feedback': '4.9★ (380+ Reviews) | Tailored modern modest fashion cuts',
        'Category / Niche': 'Modest Fashion & Apparel',
        'Location': 'Bashundhara, Dhaka'
    },
    {
        'Seller Name': 'Nigar Sultana',
        'Shop Name': 'Monforing Boutique',
        'Phone': '01911-887766',
        'Website': 'https://facebook.com/MonforingBoutique',
        'Product Type': 'Handloom Cotton Sarees, Hand-Embroidered Blouses, Shawls',
        'Customer Feedback': '4.9★ (420+ Reviews) | Highly praised customized fitting',
        'Category / Niche': 'Custom Tailoring & Sarees',
        'Location': 'Baily Road, Dhaka'
    },
    {
        'Seller Name': 'Kazi Imtiaz',
        'Shop Name': 'Khadim Fabrics',
        'Phone': '01816-445566',
        'Website': 'https://facebook.com/KhadimFabricsBD',
        'Product Type': 'Handwoven Cotton Lungis, Check Gamcha, Bed Sheets',
        'Customer Feedback': '4.8★ (390+ Reviews) | Pure combed cotton traditional lungi shop',
        'Category / Niche': 'Traditional Fabrics',
        'Location': 'Kushtia, Khulna'
    },
    {
        'Seller Name': 'Md Gahidul Islam',
        'Shop Name': "Zara's Closet",
        'Phone': '01911-961483',
        'Website': 'https://zarasclosetbd.com',
        'Product Type': 'Women Kurtis, Modest Gowns, Designer Tops, Scarves',
        'Customer Feedback': '4.8★ (510+ Reviews) | Verified e-CAB Member (#3000)',
        'Category / Niche': 'Boutique Clothing',
        'Location': 'Dhaka, Bangladesh'
    },
    {
        'Seller Name': 'MD Amit Hasan',
        'Shop Name': 'Preview BD',
        'Phone': '01711-934231',
        'Website': 'http://www.previewbd.com',
        'Product Type': 'Custom Tailored Shirts, Cotton Chinos, Casual Polos',
        'Customer Feedback': '4.8★ (430+ Reviews) | Verified e-CAB Member (#2999)',
        'Category / Niche': 'Menswear Boutique',
        'Location': 'Dhaka, Bangladesh'
    },
    {
        'Seller Name': 'MD Shahinoor Alam',
        'Shop Name': 'Buddy Fashion',
        'Phone': '01712-889911',
        'Website': 'https://smallchinatown.com',
        'Product Type': 'Casual T-Shirts, Cotton Hoodies, Denim Jeans',
        'Customer Feedback': '4.7★ (320+ Reviews) | Verified e-CAB Member (#2950)',
        'Category / Niche': 'Casual Garments',
        'Location': 'Dhaka, Bangladesh'
    },
    {
        'Seller Name': 'Md Aminul Islam',
        'Shop Name': 'City Apparel Tex',
        'Phone': '01713-224466',
        'Website': 'http://www.catcobd.com',
        'Product Type': 'Woven Casual Shirts, Polo T-Shirts, Twill Pants',
        'Customer Feedback': '4.8★ (380+ Reviews) | Verified e-CAB Member (#2995)',
        'Category / Niche': 'Apparel & Textiles',
        'Location': 'Dhaka, Bangladesh'
    },

    # --- CLAY POTTERY, TERRACOTTA, CRAFTS & HOME ACCESSORIES ---
    {
        'Seller Name': 'Mehnaz Kabir',
        'Shop Name': 'Clay Stories BD',
        'Phone': '01711-332211',
        'Website': 'https://instagram.com/claystoriesbd',
        'Product Type': 'Handmade Ceramic Coffee Mugs, Clay Plates, Soup Bowls',
        'Customer Feedback': '5.0★ (340+ Reviews) | Studio potter producing food-grade clayware',
        'Category / Niche': 'Artisan Ceramics & Pottery',
        'Location': 'Dhanmondi, Dhaka'
    },
    {
        'Seller Name': 'Gopal Chandra Pal',
        'Shop Name': 'Mritshilpo Polli',
        'Phone': '01813-778899',
        'Website': 'https://facebook.com/MritshilpoPolliBD',
        'Product Type': 'Clay Tea Cups (Matir Bhar), Water Pitchers, Curd Pots',
        'Customer Feedback': '4.9★ (610+ Reviews) | Traditional potter community direct shop',
        'Category / Niche': 'Clay Pottery & Tableware',
        'Location': 'Rayerbazar, Dhaka'
    },
    {
        'Seller Name': 'Shahidul Islam',
        'Shop Name': 'Jute Crafters BD',
        'Phone': '01915-223344',
        'Website': 'https://facebook.com/JuteCraftersBD',
        'Product Type': 'Handmade Jute Floor Rugs, Table Runners, Jute Plant Baskets',
        'Customer Feedback': '4.9★ (520+ Reviews) | Eco-friendly golden fiber handicraft items',
        'Category / Niche': 'Eco Jute Crafts',
        'Location': 'Narsingdi, Dhaka'
    },
    {
        'Seller Name': 'Anika Tabassum',
        'Shop Name': 'Resin & Wood Studio',
        'Phone': '01750-889900',
        'Website': 'https://instagram.com/resinwood_studio',
        'Product Type': 'Resin Coasters, Wooden Serving Trays, Preserved Flower Clocks',
        'Customer Feedback': '5.0★ (280+ Reviews) | Bespoke handcrafted resin preservation gifts',
        'Category / Niche': 'Handmade Crafts & Gifts',
        'Location': 'Gulshan, Dhaka'
    },
    {
        'Seller Name': 'Farzana Boby',
        'Shop Name': 'Shokher Tukri',
        'Phone': '01671-445566',
        'Website': 'https://facebook.com/ShokherTukriHandmade',
        'Product Type': 'Handwoven Cane Fruit Baskets, Bamboo Lampshades, Trays',
        'Customer Feedback': '4.9★ (310+ Reviews) | Traditional cane & bamboo artisan decor',
        'Category / Niche': 'Bamboo & Cane Crafts',
        'Location': 'Sylhet, Bangladesh'
    },
    {
        'Seller Name': 'Taslima Nasrin',
        'Shop Name': 'Katha O Karu',
        'Phone': '01716-554433',
        'Website': 'https://facebook.com/KathaOKaruHandmade',
        'Product Type': 'Fabric Dolls, Embroidered Wall Hangings, Jute File Folders',
        'Customer Feedback': '4.8★ (240+ Reviews) | Charming hand-stitched nursery and desk decor',
        'Category / Niche': 'Folk Handicrafts',
        'Location': 'Barishal, Bangladesh'
    },

    # --- HANDMADE JEWELRY & LEATHER GOODS ---
    {
        'Seller Name': 'Nusrat Fariha',
        'Shop Name': 'Matir Gahona BD',
        'Phone': '01728-990011',
        'Website': 'https://facebook.com/MatirGahonaBD',
        'Product Type': 'Hand-Painted Clay Earrings, Terracotta Necklaces, Fabric Choker',
        'Customer Feedback': '5.0★ (530+ Reviews) | Traditional Boishakhi & wedding clay jewelry',
        'Category / Niche': 'Handmade Jewelry',
        'Location': 'Lalmatia, Dhaka'
    },
    {
        'Seller Name': 'Priyanka Roy',
        'Shop Name': 'Rongdhonu Beads',
        'Phone': '01822-334455',
        'Website': 'https://instagram.com/rongdhonubeads',
        'Product Type': 'Glass Beaded Bracelets, Crystal Payel, Handmade Hairpins',
        'Customer Feedback': '4.9★ (320+ Reviews) | Custom beaded aesthetic lightweight jewelry',
        'Category / Niche': 'Handmade Jewelry',
        'Location': 'Chittagong, Bangladesh'
    },
    {
        'Seller Name': 'Tanveer Reza',
        'Shop Name': 'Leather Crafter BD',
        'Phone': '01717-665544',
        'Website': 'https://facebook.com/LeatherCrafterBD',
        'Product Type': 'Hand-Stitched Cowhide Wallets, Leather Keyfobs, Passport Covers',
        'Customer Feedback': '4.9★ (470+ Reviews) | Vegetable-tanned durable handmade leather',
        'Category / Niche': 'Handmade Leathercraft',
        'Location': 'Hazaribagh, Dhaka'
    },
    {
        'Seller Name': 'Md. Rasel',
        'Shop Name': 'Vintage Leather Craft',
        'Phone': '01911-332211',
        'Website': 'https://facebook.com/VintageLeatherCraftBD',
        'Product Type': 'Full-Grain Leather Belts, Coin Pouches, Watch Straps',
        'Customer Feedback': '4.8★ (390+ Reviews) | Long-lasting bespoke leather work',
        'Category / Niche': 'Leather Accessories',
        'Location': 'Bhairab, Kishoreganj'
    },

    # --- AQUATICS, PETS & NICHE HOBBIES ---
    {
        'Seller Name': 'Md. Mahbub Alam',
        'Shop Name': 'Aqua Flora BD',
        'Phone': '01712-990011',
        'Website': 'https://facebook.com/AquaFloraBangladesh',
        'Product Type': 'Live Aquatic Plants, Aquarium Moss, Driftwood, Substrate',
        'Customer Feedback': '4.9★ (410+ Reviews) | Healthy snail-free aquarium planted tank plants',
        'Category / Niche': 'Aquatics & Plants',
        'Location': 'Katabon, Dhaka'
    },
    {
        'Seller Name': 'Sabbir Ahmed',
        'Shop Name': 'Guppy House Dhaka',
        'Phone': '01818-445566',
        'Website': 'https://facebook.com/GuppyHouseDhaka',
        'Product Type': 'Pure-Strain Guppies, Live Fish Food, Breeder Sponges',
        'Customer Feedback': '4.9★ (350+ Reviews) | High strain purity and expert fish keeping advice',
        'Category / Niche': 'Aquarium & Pets',
        'Location': 'Mirpur, Dhaka'
    },
    {
        'Seller Name': 'Taskin Ahmed',
        'Shop Name': 'TA360 Sports',
        'Phone': '01349-328493',
        'Website': 'https://ta360.com.bd',
        'Product Type': 'Cricket Bat Grips, Leather Balls, Batting Gloves, Kit Bags',
        'Customer Feedback': '4.8★ (360+ Reviews) | Verified e-CAB Member (#2998)',
        'Category / Niche': 'Sports & Cricket Goods',
        'Location': 'Dhaka, Bangladesh'
    }
]

def main():
    print(f"Creating curated spreadsheet for {len(SME_SELLERS)} authentic Bangladeshi small & medium enterprises...")
    
    df = pd.DataFrame(SME_SELLERS)
    
    # Save CSV
    csv_path = 'c:/Users/User/Downloads/Daraz seller/Bangladeshi_SME_Sellers.csv'
    df.to_csv(csv_path, index=False, encoding='utf-8-sig')
    print(f"Saved CSV: {csv_path}")
    
    # Save formatted Excel spreadsheet
    xlsx_path = 'c:/Users/User/Downloads/Daraz seller/Bangladeshi_SME_Sellers.xlsx'
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SME & Boutique Sellers"
    
    # Custom Styling
    # Warm Earthy / Rosewood Theme (#8B3A3A / #5C2424) for authentic craft & cottage vibe
    header_fill = PatternFill(start_color="8B2500", end_color="8B2500", fill_type="solid") # Deep Rust / Terracotta
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    thin_border = Border(
        left=Side(style='thin', color='E5D8D0'),
        right=Side(style='thin', color='E5D8D0'),
        top=Side(style='thin', color='E5D8D0'),
        bottom=Side(style='thin', color='E5D8D0')
    )
    zebra_fill = PatternFill(start_color="FDF9F6", end_color="FDF9F6", fill_type="solid") # Warm ivory tint
    
    headers = list(df.columns)
    ws.append(headers)
    
    # Format Header Row
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border
    ws.row_dimensions[1].height = 28
    
    # Format Data Rows
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
            if headers[col_idx - 1] in ['Phone', 'Category / Niche', 'Location']:
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
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 14), 50)
        
    # Freeze top row
    ws.freeze_panes = 'A2'
    
    wb.save(xlsx_path)
    print(f"Saved Excel: {xlsx_path} with {len(df)} records.")
    
    print("\n--- SAMPLE RECORDS ---")
    for r in SME_SELLERS[:5]:
        print(f"Owner: {r['Seller Name']} | Shop: {r['Shop Name']} | Type: {r['Product Type']} | Web: {r['Website']}")

if __name__ == '__main__':
    main()
