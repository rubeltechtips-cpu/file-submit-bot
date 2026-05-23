import logging
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Updater, CommandHandler, CallbackQueryHandler, MessageHandler, Filters
import json
import os
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import hashlib

# Bot Configuration
BOT_TOKEN = os.getenv "8723772018:AAHZqWXkTabPQ-PK35NCdne3OTmHzKOz9AI"
CHANNEL_USERNAME = "@quick_sell_bd"
CHANNEL_URL = "https://t.me/quick_sell_bd"
ADMIN_IDS = [8061006207]

# Google Sheets Configuration
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/109yQRUwOtbmmsfQY5gvSVfMaMcSsnZqelwQzI8lqe4Y/edit?usp=sharing"
SHEET_NAME = "Sheet1"

# Global mapping for short IDs
category_short_map = {}
vip_short_map = {}

def get_short_id(long_id, map_dict):
    short_id = hashlib.md5(long_id.encode()).hexdigest()[:8]
    map_dict[short_id] = long_id
    return short_id

def get_long_id(short_id, map_dict):
    return map_dict.get(short_id, short_id)

def get_emoji_for_category(name):
    emoji_map = {
        'facebook': '📘', 'instagram': '📸', 'youtube': '▶️', 'tiktok': '🎵',
        'twitter': '🐦', 'whatsapp': '💬', 'telegram': '✈️', 'linkedin': '🔗',
        'hotmail': '📧', 'outlook': '📧', 'gmail': '📧', 'cookies': '🍪',
        'friend': '👥', '2fa': '🔐'
    }
    name_lower = name.lower()
    for key, emoji in emoji_map.items():
        if key in name_lower:
            return emoji
    return '📁'

def setup_google_sheets():
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
        creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_url(GOOGLE_SHEET_URL)
        worksheet = sheet.worksheet(SHEET_NAME)
        headers = worksheet.row_values(1)
        if not headers:
            headers = ['Order ID', 'User ID', 'Name', 'Username', 'Order Time', 'Category', 
                      'Quantity', 'Rate (TK per piece)', 'Link', 'Payment Method', 'Payment Number', 'Status', 'User Type']
            worksheet.insert_row(headers, 1)
        return worksheet
    except Exception as e:
        logging.error(f"Google Sheets setup error: {e}")
        return None

google_sheet = setup_google_sheets()

DATA_FILE = "user_data.json"
CATEGORIES_FILE = "categories.json"
ORDERS_FILE = "orders.json"
VIP_USERS_FILE = "vip_users.json"

logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def save_data(data):
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

def load_categories():
    if os.path.exists(CATEGORIES_FILE):
        with open(CATEGORIES_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {
        "facebook": {"name": "Facebook", "price": 10, "emoji": "📘"},
        "instagram": {"name": "Instagram", "price": 15, "emoji": "📸"},
        "youtube": {"name": "YouTube", "price": 20, "emoji": "▶️"},
        "tiktok": {"name": "TikTok", "price": 25, "emoji": "🎵"}
    }

def save_categories(categories):
    with open(CATEGORIES_FILE, 'w', encoding='utf-8') as f:
        json.dump(categories, f, ensure_ascii=False, indent=4)

def load_orders():
    if os.path.exists(ORDERS_FILE):
        with open(ORDERS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def save_orders(orders):
    with open(ORDERS_FILE, 'w', encoding='utf-8') as f:
        json.dump(orders, f, ensure_ascii=False, indent=4)

def load_vip_users():
    if os.path.exists(VIP_USERS_FILE):
        with open(VIP_USERS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def save_vip_users(vip_users):
    with open(VIP_USERS_FILE, 'w', encoding='utf-8') as f:
        json.dump(vip_users, f, ensure_ascii=False, indent=4)

def is_vip_user(user_id):
    vip_users = load_vip_users()
    return str(user_id) in vip_users

def get_user_price(category_id, user_id):
    categories = load_categories()
    category = categories.get(category_id, {})
    default_price = category.get('price', 0)
    if is_vip_user(user_id):
        vip_users = load_vip_users()
        user_vip_data = vip_users.get(str(user_id), {})
        vip_rates = user_vip_data.get('custom_rates', {})
        if category_id in vip_rates:
            return vip_rates[category_id]
    return default_price

def save_to_google_sheets(order_data):
    try:
        if google_sheet:
            row = [order_data['order_id'], order_data['user_id'], order_data['user_name'], order_data['username'],
                   order_data['order_time'], order_data['category'], order_data['quantity'], order_data['rate'],
                   order_data['link'], order_data['payment_method'], order_data['payment_number'], 
                   order_data['status'], order_data.get('user_type', 'Normal')]
            google_sheet.append_row(row)
            return True
        return False
    except Exception as e:
        logging.error(f"Error saving to Google Sheets: {e}")
        return False

def get_user_info(update, context):
    user = update.effective_user
    return {
        "user_id": user.id,
        "name": user.full_name if user.full_name else "N/A",
        "username": f"@{user.username}" if user.username else "No username",
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

def is_member(bot, user_id):
    try:
        member = bot.get_chat_member(chat_id=CHANNEL_USERNAME, user_id=user_id)
        return member.status in ['member', 'administrator', 'creator']
    except:
        return False

def start(update, context):
    user_id = update.effective_user.id
    if user_id in ADMIN_IDS:
        user_data = load_data()
        if str(user_id) not in user_data:
            user_data[str(user_id)] = {"payment_method": "ADMIN", "payment_number": "ADMIN"}
            save_data(user_data)
        if not is_member(context.bot, user_id):
            keyboard = [[InlineKeyboardButton("Join Channel", url=CHANNEL_URL)],
                        [InlineKeyboardButton("Check Membership", callback_data="check_join")]]
            update.message.reply_text(f"Please join our channel: {CHANNEL_USERNAME}", reply_markup=InlineKeyboardMarkup(keyboard))
            return
        show_main_menu_from_msg(update, user_id)
        return
    
    user_data = load_data()
    if str(user_id) not in user_data:
        user_data[str(user_id)] = {"payment_method": None, "payment_number": None}
        save_data(user_data)
    if not is_member(context.bot, user_id):
        keyboard = [[InlineKeyboardButton("Join Channel", url=CHANNEL_URL)],
                    [InlineKeyboardButton("Check Membership", callback_data="check_join")]]
        update.message.reply_text(f"Please join our channel first: {CHANNEL_USERNAME}", reply_markup=InlineKeyboardMarkup(keyboard))
        return
    if user_data[str(user_id)]["payment_method"] is None:
        ask_payment_method(update, context)
    else:
        show_main_menu_from_msg(update, user_id)

def ask_payment_method(update, context):
    keyboard = [[InlineKeyboardButton("bKash", callback_data="method_bkash")],
                [InlineKeyboardButton("Nagad", callback_data="method_nagad")],
                [InlineKeyboardButton("Rocket", callback_data="method_rocket")],
                [InlineKeyboardButton("Binance", callback_data="method_binance")]]
    update.message.reply_text("Select your payment method:", reply_markup=InlineKeyboardMarkup(keyboard))

def ask_payment_number(update, context, method):
    context.user_data['temp_method'] = method
    query = update.callback_query
    query.message.reply_text(f"Enter your {method} number:")
    context.user_data['waiting_for_payment_number'] = True

def save_payment_number(update, context):
    user_id = update.effective_user.id
    number = update.message.text.strip()
    method = context.user_data.get('temp_method')
    if not method:
        update.message.reply_text("Please use /start")
        return
    user_data = load_data()
    user_data[str(user_id)]["payment_method"] = method
    user_data[str(user_id)]["payment_number"] = number
    save_data(user_data)
    context.user_data['waiting_for_payment_number'] = False
    context.user_data['temp_method'] = None
    update.message.reply_text(f"{method} number {number} saved successfully!")
    show_main_menu_from_msg(update, user_id)

def show_main_menu_from_msg(update, user_id):
    is_admin = user_id in ADMIN_IDS
    keyboard = [[InlineKeyboardButton("File Submit", callback_data="submit")],
                [InlineKeyboardButton("Admin Support", callback_data="support")],
                [InlineKeyboardButton("Settings", callback_data="setting")]]
    if is_admin:
        keyboard.append([InlineKeyboardButton("Admin Panel", callback_data="admin_panel")])
    try:
        if hasattr(update, 'callback_query') and update.callback_query:
            update.callback_query.edit_message_text("Main Menu", reply_markup=InlineKeyboardMarkup(keyboard))
        else:
            update.message.reply_text("Main Menu", reply_markup=InlineKeyboardMarkup(keyboard))
    except:
        update.message.reply_text("Main Menu", reply_markup=InlineKeyboardMarkup(keyboard))

def show_categories_menu(update, user_id):
    categories = load_categories()
    keyboard = []
    for cat_id, cat_info in categories.items():
        price = get_user_price(cat_id, user_id) if is_vip_user(user_id) else cat_info['price']
        price_str = f"{price:.2f}".rstrip('0').rstrip('.') if isinstance(price, float) else str(price)
        button_text = f"{cat_info['emoji']} {cat_info['name']} - {price_str} TK"
        if is_vip_user(user_id):
            button_text += " (VIP)"
        short_id = get_short_id(cat_id, category_short_map)
        keyboard.append([InlineKeyboardButton(button_text, callback_data=f"s_{short_id}")])
    keyboard.append([InlineKeyboardButton("Back", callback_data="back_main")])
    try:
        if hasattr(update, 'callback_query') and update.callback_query:
            update.callback_query.edit_message_text("Select Service:", reply_markup=InlineKeyboardMarkup(keyboard))
        else:
            update.message.reply_text("Select Service:", reply_markup=InlineKeyboardMarkup(keyboard))
    except:
        update.message.reply_text("Select Service:", reply_markup=InlineKeyboardMarkup(keyboard))

def ask_quantity(update, context, short_id):
    cat_id = get_long_id(short_id, category_short_map)
    categories = load_categories()
    category = categories.get(cat_id)
    user_id = update.effective_user.id
    if not category:
        show_categories_menu(update, user_id)
        return
    context.user_data['selected_category'] = cat_id
    price = get_user_price(cat_id, user_id)
    price_str = f"{price:.2f}".rstrip('0').rstrip('.') if isinstance(price, float) else str(price)
    user_type = "VIP" if is_vip_user(user_id) else "Normal"
    keyboard = [[InlineKeyboardButton("Back", callback_data="submit")]]
    try:
        update.callback_query.edit_message_text(
            f"Selected: {category['emoji']} {category['name']}\nUser Type: {user_type}\nPrice: {price_str} TK per piece\n\nEnter quantity:",
            reply_markup=InlineKeyboardMarkup(keyboard))
    except:
        update.callback_query.message.reply_text(
            f"Selected: {category['emoji']} {category['name']}\nUser Type: {user_type}\nPrice: {price_str} TK per piece\n\nEnter quantity:",
            reply_markup=InlineKeyboardMarkup(keyboard))
    context.user_data['awaiting_quantity'] = True

def process_quantity(update, context, user_id, quantity_str):
    try:
        quantity = int(quantity_str)
        if quantity <= 0:
            update.message.reply_text("Please enter a valid number!")
            return False
        cat_id = context.user_data.get('selected_category')
        if not cat_id:
            update.message.reply_text("Please use /start")
            return False
        categories = load_categories()
        category = categories.get(cat_id)
        if not category:
            update.message.reply_text("Category not found!")
            return False
        price = get_user_price(cat_id, user_id)
        context.user_data['order_quantity'] = quantity
        context.user_data['order_total_price'] = quantity * price
        context.user_data['awaiting_quantity'] = False
        update.message.reply_text("Please provide your Excel file link:\n\n(e.g., https://drive.google.com/file/d/xxxxx)")
        context.user_data['awaiting_link'] = True
        return True
    except ValueError:
        update.message.reply_text("Please enter a valid number! (e.g., 245)")
        return False

def save_order_with_link(update, context, user_id, link):
    cat_id = context.user_data.get('selected_category')
    quantity = context.user_data.get('order_quantity')
    if not all([cat_id, quantity]):
        update.message.reply_text("Please use /start")
        return
    categories = load_categories()
    category = categories.get(cat_id)
    user_data = load_data()
    user_info = user_data.get(str(user_id), {})
    user_info_full = get_user_info(update, context)
    if not category:
        update.message.reply_text("Category not found!")
        return
    price = get_user_price(cat_id, user_id)
    user_type = "VIP" if is_vip_user(user_id) else "Normal"
    orders = load_orders()
    order_id = str(len(orders) + 1)
    price_str = f"{price:.2f}".rstrip('0').rstrip('.') if isinstance(price, float) else str(price)
    order_data = {"order_id": order_id, "user_id": user_id, "user_name": user_info_full["name"],
                  "username": user_info_full["username"], "order_time": user_info_full["timestamp"],
                  "category": category['name'], "quantity": quantity, "rate": price, "link": link,
                  "payment_method": user_info.get('payment_method'), "payment_number": user_info.get('payment_number'),
                  "status": "pending", "user_type": user_type}
    orders[order_id] = order_data
    save_orders(orders)
    save_to_google_sheets(order_data)
    keyboard = [[InlineKeyboardButton("Home", callback_data="back_main")]]
    update.message.reply_text(f"✅ Order submitted!\n\n📁 {category['emoji']} {category['name']}\n👤 Type: {user_type}\n🔢 Qty: {quantity}\n💰 Rate: {price_str} TK", reply_markup=InlineKeyboardMarkup(keyboard))
    admin_text = f"📦 New Order!\n\n🆔 User: {user_id}\n👤 Name: {user_info_full['name']}\n🔖 Username: {user_info_full['username']}\n⭐ Type: {user_type}\n⏰ Time: {user_info_full['timestamp']}\n📁 Category: {category['emoji']} {category['name']}\n🔢 Qty: {quantity}\n💰 Rate: {price_str} TK\n🔗 Link: {link}\n💳 Payment: {user_info.get('payment_method')}\n📞 Number: {user_info.get('payment_number')}"
    for admin_id in ADMIN_IDS:
        try:
            context.bot.send_message(chat_id=admin_id, text=admin_text)
        except:
            pass
    context.user_data.clear()

def show_support_menu(update, user_id):
    keyboard = [[InlineKeyboardButton("Back", callback_data="back_main")]]
    try:
        if hasattr(update, 'callback_query') and update.callback_query:
            update.callback_query.edit_message_text("🆘 Admin Support\n\nContact: @Rubel_QSB", reply_markup=InlineKeyboardMarkup(keyboard))
        else:
            update.message.reply_text("🆘 Admin Support\n\nContact: @Rubel_QSB", reply_markup=InlineKeyboardMarkup(keyboard))
    except:
        update.message.reply_text("🆘 Admin Support\n\nContact: @Rubel_QSB", reply_markup=InlineKeyboardMarkup(keyboard))

def show_setting_menu(update, user_id):
    keyboard = [[InlineKeyboardButton("Change Payment", callback_data="change_payment")],
                [InlineKeyboardButton("Back", callback_data="back_main")]]
    user_data = load_data()
    pm = user_data.get(str(user_id), {})
    vip_status = "✅ VIP Member" if is_vip_user(user_id) else "Normal User"
    try:
        if hasattr(update, 'callback_query') and update.callback_query:
            update.callback_query.edit_message_text(f"⚙️ Settings\n\n{vip_status}\n\n💳 Payment: {pm.get('payment_method', 'N/A')}\n📞 Number: {pm.get('payment_number', 'N/A')}", reply_markup=InlineKeyboardMarkup(keyboard))
        else:
            update.message.reply_text(f"⚙️ Settings\n\n{vip_status}\n\n💳 Payment: {pm.get('payment_method', 'N/A')}\n📞 Number: {pm.get('payment_number', 'N/A')}", reply_markup=InlineKeyboardMarkup(keyboard))
    except:
        update.message.reply_text(f"⚙️ Settings\n\n{vip_status}\n\n💳 Payment: {pm.get('payment_method', 'N/A')}\n📞 Number: {pm.get('payment_number', 'N/A')}", reply_markup=InlineKeyboardMarkup(keyboard))

def show_admin_panel(update, user_id):
    keyboard = [[InlineKeyboardButton("📋 Order List", callback_data="admin_orders")],
                [InlineKeyboardButton("📁 Manage Categories", callback_data="manage_categories")],
                [InlineKeyboardButton("👑 VIP Users", callback_data="admin_vip")],
                [InlineKeyboardButton("📢 Broadcast", callback_data="admin_broadcast")],
                [InlineKeyboardButton("🔙 Back", callback_data="back_main")]]
    try:
        update.callback_query.edit_message_text("👑 Admin Panel", reply_markup=InlineKeyboardMarkup(keyboard))
    except:
        update.callback_query.message.reply_text("👑 Admin Panel", reply_markup=InlineKeyboardMarkup(keyboard))

def show_admin_orders(update, user_id):
    orders = load_orders()
    keyboard = [[InlineKeyboardButton("Back", callback_data="admin_panel")]]
    if not orders:
        try:
            update.callback_query.edit_message_text("No orders found.", reply_markup=InlineKeyboardMarkup(keyboard))
        except:
            update.callback_query.message.reply_text("No orders found.", reply_markup=InlineKeyboardMarkup(keyboard))
        return
    order_text = "📋 Order List:\n\n"
    for oid, order in list(orders.items())[-10:]:
        rate_str = f"{order['rate']:.2f}".rstrip('0').rstrip('.') if isinstance(order['rate'], float) else str(order['rate'])
        order_text += f"━━━━━━━━━━━━━━━━━━━━\n🆔 #{oid}\n👤 User: {order['user_id']}\n📛 Name: {order.get('user_name', 'N/A')}\n⭐ Type: {order.get('user_type', 'Normal')}\n📁 Cat: {order['category']}\n🔢 Qty: {order['quantity']}\n💰 Rate: {rate_str} TK\n✅ Status: {order['status']}\n"
    try:
        update.callback_query.edit_message_text(order_text[:4000], reply_markup=InlineKeyboardMarkup(keyboard))
    except:
        update.callback_query.message.reply_text(order_text[:4000], reply_markup=InlineKeyboardMarkup(keyboard))

def show_admin_categories(update, user_id):
    categories = load_categories()
    keyboard = []
    for cat_id, cat_info in categories.items():
        price_str = f"{cat_info['price']:.2f}".rstrip('0').rstrip('.') if isinstance(cat_info['price'], float) else str(cat_info['price'])
        short_id = get_short_id(cat_id, category_short_map)
        keyboard.append([InlineKeyboardButton(f"✏️ {cat_info['emoji']} {cat_info['name']} - {price_str} TK", callback_data=f"edit_{short_id}")])
    keyboard.append([InlineKeyboardButton("➕ Add New", callback_data="add_cat_start")])
    keyboard.append([InlineKeyboardButton("🔙 Back", callback_data="admin_panel")])
    try:
        update.callback_query.edit_message_text("📁 Manage Categories\n\n✨ Emoji auto-added\n💰 Decimal rates supported", reply_markup=InlineKeyboardMarkup(keyboard))
    except:
        update.callback_query.message.reply_text("📁 Manage Categories\n\n✨ Emoji auto-added\n💰 Decimal rates supported", reply_markup=InlineKeyboardMarkup(keyboard))

def add_category_start(update, context):
    context.user_data['add_cat'] = 'name'
    keyboard = [[InlineKeyboardButton("Back", callback_data="manage_categories")]]
    try:
        update.callback_query.edit_message_text("➕ Add Category\n\n📝 Step 1/2: Send category name\n\nExample: `Twitter` or `Hotmail`", reply_markup=InlineKeyboardMarkup(keyboard))
    except:
        update.callback_query.message.reply_text("➕ Add Category\n\n📝 Step 1/2: Send category name", reply_markup=InlineKeyboardMarkup(keyboard))

def process_cat_name(update, context, name):
    context.user_data['new_cat_name'] = name.strip()
    context.user_data['add_cat'] = 'price'
    keyboard = [[InlineKeyboardButton("Back", callback_data="add_cat_start")]]
    update.message.reply_text(f"📝 Name: {name.strip()}\n\n💰 Step 2/2: Send rate (TK)\n\nExample: `30` or `7.5`", reply_markup=InlineKeyboardMarkup(keyboard))

def process_cat_price(update, context, price_text):
    try:
        price = float(price_text) if '.' in price_text else int(price_text)
        if price <= 0:
            update.message.reply_text("❌ Enter positive number!")
            return
        name = context.user_data.get('new_cat_name')
        if not name:
            update.message.reply_text("❌ Session expired!")
            return
        cat_id = name.lower().replace(" ", "_").replace("+", "_")
        emoji = get_emoji_for_category(name)
        categories = load_categories()
        if cat_id in categories:
            update.message.reply_text(f"❌ '{name}' already exists!")
            return
        categories[cat_id] = {"name": name, "price": price, "emoji": emoji}
        save_categories(categories)
        price_str = f"{price:.2f}".rstrip('0').rstrip('.') if isinstance(price, float) else str(price)
        update.message.reply_text(f"✅ Added!\n\n📁 {emoji} {name}\n💰 {price_str} TK")
        context.user_data['add_cat'] = None
        context.user_data['new_cat_name'] = None
        show_admin_categories_from_msg(update, update.effective_user.id)
    except ValueError:
        update.message.reply_text("❌ Send valid number! (e.g., 30, 7.5)")

def show_admin_categories_from_msg(update, user_id):
    categories = load_categories()
    keyboard = []
    for cat_id, cat_info in categories.items():
        price_str = f"{cat_info['price']:.2f}".rstrip('0').rstrip('.') if isinstance(cat_info['price'], float) else str(cat_info['price'])
        short_id = get_short_id(cat_id, category_short_map)
        keyboard.append([InlineKeyboardButton(f"✏️ {cat_info['emoji']} {cat_info['name']} - {price_str} TK", callback_data=f"edit_{short_id}")])
    keyboard.append([InlineKeyboardButton("➕ Add New", callback_data="add_cat_start")])
    keyboard.append([InlineKeyboardButton("🔙 Back", callback_data="admin_panel")])
    update.message.reply_text("📁 Manage Categories", reply_markup=InlineKeyboardMarkup(keyboard))

def show_edit_category(update, context, short_id):
    cat_id = get_long_id(short_id, category_short_map)
    categories = load_categories()
    category = categories.get(cat_id)
    if not category:
        show_admin_categories(update, update.effective_user.id)
        return
    price_str = f"{category['price']:.2f}".rstrip('0').rstrip('.') if isinstance(category['price'], float) else str(category['price'])
    keyboard = [[InlineKeyboardButton("📝 Change Name", callback_data=f"ren_{short_id}")],
                [InlineKeyboardButton("💰 Change Price", callback_data=f"prc_{short_id}")],
                [InlineKeyboardButton("🗑 Delete", callback_data=f"del_{short_id}")],
                [InlineKeyboardButton("🔙 Back", callback_data="manage_categories")]]
    try:
        update.callback_query.edit_message_text(f"✏️ {category['emoji']} {category['name']}\n💰 Price: {price_str} TK", reply_markup=InlineKeyboardMarkup(keyboard))
    except:
        update.callback_query.message.reply_text(f"✏️ {category['emoji']} {category['name']}\n💰 Price: {price_str} TK", reply_markup=InlineKeyboardMarkup(keyboard))

def show_admin_vip(update, user_id):
    keyboard = [[InlineKeyboardButton("➕ Add VIP", callback_data="add_vip_start")],
                [InlineKeyboardButton("📋 VIP List", callback_data="vip_list")],
                [InlineKeyboardButton("✏️ Edit Rates", callback_data="edit_vip_rates")],
                [InlineKeyboardButton("🗑 Remove VIP", callback_data="remove_vip_start")],
                [InlineKeyboardButton("🔙 Back", callback_data="admin_panel")]]
    try:
        update.callback_query.edit_message_text("👑 VIP User Management", reply_markup=InlineKeyboardMarkup(keyboard))
    except:
        update.callback_query.message.reply_text("👑 VIP User Management", reply_markup=InlineKeyboardMarkup(keyboard))

def add_vip_start(update, context):
    context.user_data['add_vip'] = True
    keyboard = [[InlineKeyboardButton("Back", callback_data="admin_vip")]]
    try:
        update.callback_query.edit_message_text("➕ Add VIP\n\nSend User ID:\nExample: `8555327754`", reply_markup=InlineKeyboardMarkup(keyboard))
    except:
        update.callback_query.message.reply_text("➕ Add VIP\n\nSend User ID", reply_markup=InlineKeyboardMarkup(keyboard))

def process_add_vip(update, context, user_id_text):
    if not user_id_text.isdigit():
        update.message.reply_text("❌ Send numeric User ID!")
        return
    vip_users = load_vip_users()
    if user_id_text in vip_users:
        update.message.reply_text(f"❌ User {user_id_text} is already VIP!")
        return
    vip_users[user_id_text] = {"user_id": user_id_text, "name": "Unknown", "username": "Unknown",
                               "added_by": str(update.effective_user.id), "added_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "custom_rates": {}}
    save_vip_users(vip_users)
    update.message.reply_text(f"✅ VIP Added: {user_id_text}\n\nNow set custom rates.")
    context.user_data['add_vip'] = False
    show_vip_rates_menu(update, context, user_id_text)

def show_vip_rates_menu(update, context, user_id_text):
    categories = load_categories()
    vip_users = load_vip_users()
    current = vip_users.get(user_id_text, {}).get('custom_rates', {})
    keyboard = []
    for cat_id, cat_info in categories.items():
        cur_rate = current.get(cat_id, cat_info['price'])
        rate_str = f"{cur_rate:.2f}".rstrip('0').rstrip('.') if isinstance(cur_rate, float) else str(cur_rate)
        short_id = get_short_id(cat_id, category_short_map)
        keyboard.append([InlineKeyboardButton(f"{cat_info['emoji']} {cat_info['name']} - {rate_str} TK", callback_data=f"vrate_{short_id}_{user_id_text}")])
    keyboard.append([InlineKeyboardButton("✅ Done", callback_data="admin_vip")])
    try:
        update.callback_query.edit_message_text(f"👑 Set Rates for VIP {user_id_text}", reply_markup=InlineKeyboardMarkup(keyboard))
    except:
        update.message.reply_text(f"👑 Set Rates for VIP {user_id_text}", reply_markup=InlineKeyboardMarkup(keyboard))

def ask_vip_rate(update, context, short_id, user_id_text):
    cat_id = get_long_id(short_id, category_short_map)
    categories = load_categories()
    category = categories.get(cat_id, {})
    context.user_data['vip_rate_user'] = user_id_text
    context.user_data['vip_rate_cat'] = cat_id
    keyboard = [[InlineKeyboardButton("Back", callback_data=f"back_vip_{user_id_text}")]]
    try:
        update.callback_query.edit_message_text(f"Set rate for {category['emoji']} {category['name']}\nDefault: {category['price']} TK\n\nSend custom rate:", reply_markup=InlineKeyboardMarkup(keyboard))
    except:
        update.callback_query.message.reply_text(f"Set rate for {category['emoji']} {category['name']}\nSend custom rate:", reply_markup=InlineKeyboardMarkup(keyboard))

def save_vip_rate(update, context, rate_text):
    try:
        rate = float(rate_text) if '.' in rate_text else int(rate_text)
        if rate <= 0:
            update.message.reply_text("❌ Positive number required!")
            return
        user_id_text = context.user_data.get('vip_rate_user')
        cat_id = context.user_data.get('vip_rate_cat')
        if not user_id_text or not cat_id:
            update.message.reply_text("❌ Session expired!")
            return
        vip_users = load_vip_users()
        if user_id_text not in vip_users:
            update.message.reply_text("❌ VIP not found!")
            return
        if 'custom_rates' not in vip_users[user_id_text]:
            vip_users[user_id_text]['custom_rates'] = {}
        vip_users[user_id_text]['custom_rates'][cat_id] = rate
        save_vip_users(vip_users)
        categories = load_categories()
        category = categories.get(cat_id, {})
        rate_str = f"{rate:.2f}".rstrip('0').rstrip('.') if isinstance(rate, float) else str(rate)
        update.message.reply_text(f"✅ Rate set!\n{category['emoji']} {category['name']}: {rate_str} TK")
        context.user_data['vip_rate_user'] = None
        context.user_data['vip_rate_cat'] = None
        show_vip_rates_menu(update, context, user_id_text)
    except ValueError:
        update.message.reply_text("❌ Send valid number! (e.g., 30, 7.5)")

def show_vip_list(update, user_id):
    vip_users = load_vip_users()
    keyboard = [[InlineKeyboardButton("Back", callback_data="admin_vip")]]
    if not vip_users:
        try:
            update.callback_query.edit_message_text("No VIP users.", reply_markup=InlineKeyboardMarkup(keyboard))
        except:
            update.callback_query.message.reply_text("No VIP users.", reply_markup=InlineKeyboardMarkup(keyboard))
        return
    text = "📋 VIP List:\n\n"
    for uid, info in vip_users.items():
        text += f"━━━━━━━━━━━━━━━━━━━━\n🆔 {uid}\n👤 {info.get('name', 'N/A')}\n📅 {info.get('added_time', 'N/A')}\n"
        if info.get('custom_rates'):
            text += f"⭐ Custom rates:\n"
            for cat_id, r in info['custom_rates'].items():
                text += f"   • {cat_id}: {r} TK\n"
        text += "\n"
    try:
        update.callback_query.edit_message_text(text[:4000], reply_markup=InlineKeyboardMarkup(keyboard))
    except:
        update.callback_query.message.reply_text(text[:4000], reply_markup=InlineKeyboardMarkup(keyboard))

def edit_vip_rates_list(update, context):
    vip_users = load_vip_users()
    if not vip_users:
        keyboard = [[InlineKeyboardButton("Back", callback_data="admin_vip")]]
        try:
            update.callback_query.edit_message_text("No VIP users.", reply_markup=InlineKeyboardMarkup(keyboard))
        except:
            update.callback_query.message.reply_text("No VIP users.", reply_markup=InlineKeyboardMarkup(keyboard))
        return
    keyboard = []
    for uid, info in vip_users.items():
        name = info.get('name', 'Unknown')
        keyboard.append([InlineKeyboardButton(f"👤 {name} ({uid})", callback_data=f"edit_vip_{uid}")])
    keyboard.append([InlineKeyboardButton("Back", callback_data="admin_vip")])
    try:
        update.callback_query.edit_message_text("Select VIP user:", reply_markup=InlineKeyboardMarkup(keyboard))
    except:
        update.callback_query.message.reply_text("Select VIP user:", reply_markup=InlineKeyboardMarkup(keyboard))

def remove_vip_start(update, context):
    context.user_data['remove_vip'] = True
    keyboard = [[InlineKeyboardButton("Back", callback_data="admin_vip")]]
    try:
        update.callback_query.edit_message_text("🗑 Remove VIP\n\nSend User ID to remove:", reply_markup=InlineKeyboardMarkup(keyboard))
    except:
        update.callback_query.message.reply_text("🗑 Remove VIP\n\nSend User ID:", reply_markup=InlineKeyboardMarkup(keyboard))

def process_remove_vip(update, context, user_id_text):
    if not user_id_text.isdigit():
        update.message.reply_text("❌ Send numeric User ID!")
        context.user_data['remove_vip'] = False
        return
    vip_users = load_vip_users()
    if user_id_text in vip_users:
        del vip_users[user_id_text]
        save_vip_users(vip_users)
        update.message.reply_text(f"✅ VIP removed: {user_id_text}")
    else:
        update.message.reply_text(f"❌ {user_id_text} not in VIP list")
    context.user_data['remove_vip'] = False
    show_main_menu_from_msg(update, update.effective_user.id)

def button_callback(update, context):
    query = update.callback_query
    query.answer()
    user_id = update.effective_user.id
    if not is_member(context.bot, user_id):
        query.edit_message_text("Please join the channel!")
        return
    data = query.data
    
    if data == "check_join":
        if is_member(context.bot, user_id):
            keyboard = [[InlineKeyboardButton("bKash", callback_data="method_bkash")],
                        [InlineKeyboardButton("Nagad", callback_data="method_nagad")],
                        [InlineKeyboardButton("Rocket", callback_data="method_rocket")],
                        [InlineKeyboardButton("Binance", callback_data="method_binance")]]
            query.edit_message_text("Select payment method:", reply_markup=InlineKeyboardMarkup(keyboard))
        else:
            query.edit_message_text("Not a member!")
    elif data.startswith("method_"):
        method = data.split("_")[1]
        context.user_data['temp_method'] = method
        query.message.reply_text(f"Enter {method} number:")
        context.user_data['waiting_for_payment_number'] = True
    elif data == "submit":
        show_categories_menu(update, user_id)
    elif data.startswith("s_"):
        short_id = data.replace("s_", "")
        ask_quantity(update, context, short_id)
    elif data == "support":
        show_support_menu(update, user_id)
    elif data == "setting":
        show_setting_menu(update, user_id)
    elif data == "change_payment":
        keyboard = [[InlineKeyboardButton("bKash", callback_data="method_bkash")],
                    [InlineKeyboardButton("Nagad", callback_data="method_nagad")],
                    [InlineKeyboardButton("Rocket", callback_data="method_rocket")],
                    [InlineKeyboardButton("Binance", callback_data="method_binance")],
                    [InlineKeyboardButton("Back", callback_data="setting")]]
        query.edit_message_text("Select new payment:", reply_markup=InlineKeyboardMarkup(keyboard))
    elif data == "admin_panel" and user_id in ADMIN_IDS:
        show_admin_panel(update, user_id)
    elif data == "admin_orders" and user_id in ADMIN_IDS:
        show_admin_orders(update, user_id)
    elif data == "manage_categories" and user_id in ADMIN_IDS:
        show_admin_categories(update, user_id)
    elif data == "add_cat_start" and user_id in ADMIN_IDS:
        add_category_start(update, context)
    elif data.startswith("edit_") and user_id in ADMIN_IDS:
        short_id = data.replace("edit_", "")
        show_edit_category(update, context, short_id)
    elif data.startswith("ren_") and user_id in ADMIN_IDS:
        short_id = data.replace("ren_", "")
        context.user_data['ren_cat'] = short_id
        query.edit_message_text("Send new name:")
    elif data.startswith("prc_") and user_id in ADMIN_IDS:
        short_id = data.replace("prc_", "")
        context.user_data['prc_cat'] = short_id
        query.edit_message_text("Send new price (e.g., 30, 7.5):")
    elif data.startswith("del_") and user_id in ADMIN_IDS:
        short_id = data.replace("del_", "")
        cat_id = get_long_id(short_id, category_short_map)
        categories = load_categories()
        if cat_id in categories:
            del categories[cat_id]
            save_categories(categories)
            query.edit_message_text("✅ Deleted!")
            show_admin_categories(update, user_id)
    elif data == "admin_vip" and user_id in ADMIN_IDS:
        show_admin_vip(update, user_id)
    elif data == "add_vip_start" and user_id in ADMIN_IDS:
        add_vip_start(update, context)
    elif data == "vip_list" and user_id in ADMIN_IDS:
        show_vip_list(update, user_id)
    elif data == "edit_vip_rates" and user_id in ADMIN_IDS:
        edit_vip_rates_list(update, context)
    elif data.startswith("edit_vip_") and user_id in ADMIN_IDS:
        uid = data.replace("edit_vip_", "")
        show_vip_rates_menu(update, context, uid)
    elif data.startswith("vrate_") and user_id in ADMIN_IDS:
        parts = data.split("_")
        short_id = parts[1]
        uid = parts[2]
        ask_vip_rate(update, context, short_id, uid)
    elif data.startswith("back_vip_") and user_id in ADMIN_IDS:
        uid = data.replace("back_vip_", "")
        show_vip_rates_menu(update, context, uid)
    elif data == "remove_vip_start" and user_id in ADMIN_IDS:
        remove_vip_start(update, context)
    elif data == "admin_broadcast" and user_id in ADMIN_IDS:
        context.user_data['broadcast'] = True
        keyboard = [[InlineKeyboardButton("Back", callback_data="admin_panel")]]
        query.edit_message_text("Send broadcast message:", reply_markup=InlineKeyboardMarkup(keyboard))
    elif data == "back_main":
        show_main_menu_from_msg(update, user_id)

def message_handler(update, context):
    user_id = update.effective_user.id
    text = update.message.text.strip()
    
    if context.user_data.get('waiting_for_payment_number'):
        save_payment_number(update, context)
    elif context.user_data.get('awaiting_quantity'):
        process_quantity(update, context, user_id, text)
    elif context.user_data.get('awaiting_link'):
        if text.startswith('http'):
            save_order_with_link(update, context, user_id, text)
        else:
            update.message.reply_text("Send valid HTTP/HTTPS link!")
    elif context.user_data.get('add_vip') and user_id in ADMIN_IDS:
        process_add_vip(update, context, text)
        context.user_data['add_vip'] = False
    elif context.user_data.get('remove_vip') and user_id in ADMIN_IDS:
        process_remove_vip(update, context, text)
        context.user_data['remove_vip'] = False
    elif context.user_data.get('vip_rate_user') and user_id in ADMIN_IDS:
        save_vip_rate(update, context, text)
    elif context.user_data.get('add_cat') == 'name' and user_id in ADMIN_IDS:
        process_cat_name(update, context, text)
    elif context.user_data.get('add_cat') == 'price' and user_id in ADMIN_IDS:
        process_cat_price(update, context, text)
    elif context.user_data.get('broadcast') and user_id in ADMIN_IDS:
        orders = load_orders()
        sent = 0
        for uid in orders.keys():
            try:
                context.bot.send_message(chat_id=int(uid), text=text)
                sent += 1
            except:
                pass
        update.message.reply_text(f"✅ Sent to {sent} users")
        context.user_data['broadcast'] = False
        show_main_menu_from_msg(update, user_id)
    elif context.user_data.get('ren_cat') and user_id in ADMIN_IDS:
        short_id = context.user_data['ren_cat']
        cat_id = get_long_id(short_id, category_short_map)
        categories = load_categories()
        if cat_id in categories:
            new_emoji = get_emoji_for_category(text)
            categories[cat_id]['name'] = text
            categories[cat_id]['emoji'] = new_emoji
            save_categories(categories)
            update.message.reply_text(f"✅ Name changed to: {text}\n✨ Emoji: {new_emoji}")
        context.user_data['ren_cat'] = None
        show_main_menu_from_msg(update, user_id)
    elif context.user_data.get('prc_cat') and user_id in ADMIN_IDS:
        short_id = context.user_data['prc_cat']
        cat_id = get_long_id(short_id, category_short_map)
        try:
            price = float(text) if '.' in text else int(text)
            if price <= 0:
                update.message.reply_text("Enter positive number!")
                return
            categories = load_categories()
            if cat_id in categories:
                categories[cat_id]['price'] = price
                save_categories(categories)
                price_str = f"{price:.2f}".rstrip('0').rstrip('.') if isinstance(price, float) else str(price)
                update.message.reply_text(f"✅ Price: {price_str} TK")
        except:
            update.message.reply_text("Send valid number! (e.g., 30, 7.5)")
        context.user_data['prc_cat'] = None
        show_main_menu_from_msg(update, user_id)

def main():
    updater = Updater(BOT_TOKEN, use_context=True)
    dp = updater.dispatcher
    dp.add_handler(CommandHandler("start", start))
    dp.add_handler(CallbackQueryHandler(button_callback))
    dp.add_handler(MessageHandler(Filters.text & ~Filters.command, message_handler))
    updater.start_polling()
    print("🤖 Bot is running!")
    print(f"📊 Google Sheets: {'✅ Active' if google_sheet else '❌ Failed'}")
    updater.idle()

if __name__ == "__main__":
    main()