#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import os

# Add virtual environment path
sys.path.insert(0, '/home/admin202678/.virtualenvs/mybot/lib/python3.8/site-packages')

import logging
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup
from telegram.ext import Updater, CommandHandler, CallbackQueryHandler, MessageHandler, Filters
import json
import os
from datetime import datetime
import hashlib
from pathlib import Path
import openpyxl
from openpyxl import Workbook, load_workbook
import io
import re
from collections import Counter
import shutil
import threading
from http.server import HTTPServer, BaseHTTPRequestHandler
import requests
import base64

# ================ GITHUB CONFIGURATION ================
# 👇 এখানে আপনার তথ্য দিন
GITHUB_REPO = "rubelofficial999/file submit"  # আপনার GitHub রেপো নাম দিন
GITHUB_BRANCH = "main"  # অথবা "master"
GITHUB_TOKEN = "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"  # আপনার Token টি বসান

# GitHub API URLs
GITHUB_API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/contents"

# ================ GITHUB FUNCTIONS ================
def upload_to_github(file_path, github_path, commit_message):
    """
    Upload a file to GitHub repository
    Returns: (success, file_url)
    """
    try:
        # Read file content
        with open(file_path, 'rb') as f:
            content = f.read()
        
        # Encode to base64
        encoded_content = base64.b64encode(content).decode('utf-8')
        
        # Prepare API request
        url = f"{GITHUB_API_URL}/{github_path}"
        
        # Get current file SHA if exists (for update)
        sha = None
        response = requests.get(url, headers={
            'Authorization': f'token {GITHUB_TOKEN}',
            'Accept': 'application/vnd.github.v3+json'
        })
        if response.status_code == 200:
            sha = response.json().get('sha')
        
        # Prepare data
        data = {
            'message': commit_message,
            'content': encoded_content,
            'branch': GITHUB_BRANCH
        }
        if sha:
            data['sha'] = sha
        
        # Upload to GitHub
        response = requests.put(url, 
            headers={
                'Authorization': f'token {GITHUB_TOKEN}',
                'Accept': 'application/vnd.github.v3+json'
            },
            json=data
        )
        
        if response.status_code in [200, 201]:
            file_url = response.json().get('content', {}).get('html_url', '')
            logging.info(f"✅ File uploaded to GitHub: {github_path}")
            return True, file_url
        else:
            logging.error(f"❌ GitHub upload failed: {response.text}")
            return False, None
            
    except Exception as e:
        logging.error(f"❌ GitHub upload error: {e}")
        return False, None

def get_github_file_url(github_path):
    """Get the raw URL of a file on GitHub"""
    return f"https://raw.githubusercontent.com/{GITHUB_REPO}/{GITHUB_BRANCH}/{github_path}"

# ================ HTTP SERVER FOR RENDER ================
class HealthHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b'OK')
    
    def do_HEAD(self):
        self.send_response(200)
        self.end_headers()
    
    def log_message(self, format, *args):
        pass

def run_http_server():
    port = int(os.environ.get('PORT', 10000))
    server = HTTPServer(('0.0.0.0', port), HealthHandler)
    print(f"🌐 HTTP server running on port {port}")
    server.serve_forever()

# ================ BOT CONFIGURATION ================
BOT_TOKEN = "8884157908:AAH3CjxZ-J2l6oUDczu_zD-8GKKh_TleNUM"
CHANNEL_USERNAME = "@quick_sell_bd"
CHANNEL_URL = "https://t.me/quick_sell_bd"
ADMIN_IDS = [8061006207]

# File paths - Using os.path.join for cross-platform compatibility
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, "user_data.json")
CATEGORIES_FILE = os.path.join(BASE_DIR, "categories.json")
ORDERS_FILE = os.path.join(BASE_DIR, "orders.json")
VIP_USERS_FILE = os.path.join(BASE_DIR, "vip_users.json")
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
PROCESSED_DIR = os.path.join(BASE_DIR, "processed")
TEXT_DIR = os.path.join(BASE_DIR, "text_files")
REPORT_DIR = os.path.join(BASE_DIR, "reports")
REPORT_RESULTS_DIR = os.path.join(BASE_DIR, "report_results")

def ensure_directory_exists(directory):
    """Ensure directory exists, create if not"""
    if not os.path.exists(directory):
        os.makedirs(directory)
        logging.info(f"Created directory: {directory}")
    return directory

# Create directories
ensure_directory_exists(UPLOAD_DIR)
ensure_directory_exists(PROCESSED_DIR)
ensure_directory_exists(TEXT_DIR)
ensure_directory_exists(REPORT_DIR)
ensure_directory_exists(REPORT_RESULTS_DIR)

# Global mapping for short IDs
category_short_map = {}
vip_short_map = {}

logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)

def get_short_id(long_id, map_dict):
    short_id = hashlib.md5(long_id.encode()).hexdigest()[:8]
    map_dict[short_id] = long_id
    return short_id

def get_long_id(short_id, map_dict):
    return map_dict.get(short_id, short_id)

def clean_filename(filename):
    """Clean filename to remove special characters and spaces"""
    filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
    filename = re.sub(r'[()\s\[\]]', '_', filename)
    filename = re.sub(r'_+', '_', filename)
    filename = filename.strip('_')
    return filename

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

def clean_number(value):
    """Remove .0 from numbers if present"""
    if value is None:
        return ''
    str_value = str(value)
    if str_value.endswith('.0'):
        return str_value[:-2]
    return str_value

def convert_excel_to_text(excel_path, output_dir=TEXT_DIR):
    """
    Convert Excel file to text format and save as .txt file
    """
    try:
        wb = load_workbook(excel_path, data_only=True)
        ws = wb.active
        
        base_name = os.path.splitext(os.path.basename(excel_path))[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        text_filename = f"{base_name}_converted_{timestamp}.txt"
        text_path = os.path.join(output_dir, text_filename)
        
        all_values = []
        for row in ws.iter_rows(values_only=True):
            if any(cell is not None and str(cell).strip() != '' for cell in row):
                cleaned_row = [clean_number(cell) for cell in row]
                line = '\t'.join(cleaned_row)
                all_values.append(line)
        
        with open(text_path, 'w', encoding='utf-8') as f:
            for line in all_values:
                f.write(line + '\n')
        
        row_count = len(all_values)
        return text_path, row_count
        
    except Exception as e:
        logging.error(f"Error converting Excel to text: {e}")
        return None, 0

def get_existing_texts_from_all_orders():
    """Get all unique texts from all orders (for duplicate checking)"""
    orders = load_orders()
    all_texts = set()
    
    for order_id, order in orders.items():
        text_file_path = order.get('text_file_path')
        if not text_file_path or not os.path.exists(text_file_path):
            continue
        
        try:
            with open(text_file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            for line in content.split('\n'):
                if line.strip():
                    parts = line.strip().split('\t')
                    for part in parts:
                        if part.strip():
                            all_texts.add(part.strip())
        except Exception as e:
            logging.error(f"Error reading order file {text_file_path}: {e}")
            continue
    
    return all_texts

def check_duplicates_in_text_file(text_file_path, existing_texts):
    """Check if any text in the text file matches existing texts"""
    try:
        with open(text_file_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        new_lines = []
        matched_values = []
        removed_count = 0
        
        for line in lines:
            if not line.strip():
                continue
            
            parts = line.strip().split('\t')
            should_remove = False
            matched_parts = []
            
            for part in parts:
                if part.strip() in existing_texts:
                    should_remove = True
                    matched_parts.append(part.strip())
            
            if should_remove:
                removed_count += 1
                matched_values.extend(matched_parts)
            else:
                new_lines.append(line)
        
        if removed_count == 0:
            return text_file_path, 0, [], len(lines)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        cleaned_filename = f"duplicate_removed_{timestamp}.txt"
        cleaned_path = os.path.join(TEXT_DIR, cleaned_filename)
        
        with open(cleaned_path, 'w', encoding='utf-8') as f:
            for line in new_lines:
                f.write(line)
        
        remaining_count = len(new_lines)
        unique_matched = list(set(matched_values))
        
        return cleaned_path, removed_count, unique_matched, remaining_count
        
    except Exception as e:
        logging.error(f"Error checking duplicates in text file: {e}")
        return text_file_path, 0, [], 0

def process_excel_file(file_path, user_id):
    """Remove duplicates within the file itself"""
    try:
        wb = load_workbook(file_path)
        ws = wb.active
        
        rows_to_delete = []
        seen = set()
        duplicate_values = []
        total_rows = 0
        
        for row_idx, row in enumerate(ws.iter_rows(min_col=1, max_col=1, values_only=True), start=1):
            if row[0] is not None:
                value = str(row[0]).strip()
                if value:
                    total_rows += 1
                    if value in seen:
                        rows_to_delete.append(row_idx)
                        duplicate_values.append(value)
                    else:
                        seen.add(value)
        
        if rows_to_delete:
            for row_idx in sorted(rows_to_delete, reverse=True):
                ws.delete_rows(row_idx)
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            cleaned_filename = f"cleaned_{user_id}_{timestamp}.xlsx"
            cleaned_path = os.path.join(PROCESSED_DIR, cleaned_filename)
            wb.save(cleaned_path)
            
            remaining_rows = 0
            for row in ws.iter_rows(values_only=True):
                if row[0] is not None and str(row[0]).strip():
                    remaining_rows += 1
            
            return cleaned_path, duplicate_values, remaining_rows
        else:
            return file_path, [], total_rows
            
    except Exception as e:
        logging.error(f"Error processing Excel file: {e}")
        return None, [], 0

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def save_data(data):
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
    # Upload to GitHub
    upload_to_github(DATA_FILE, "data/user_data.json", "Update user data")

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
    upload_to_github(CATEGORIES_FILE, "data/categories.json", "Update categories")

def load_orders():
    if os.path.exists(ORDERS_FILE):
        with open(ORDERS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def save_orders(orders):
    with open(ORDERS_FILE, 'w', encoding='utf-8') as f:
        json.dump(orders, f, ensure_ascii=False, indent=4)
    upload_to_github(ORDERS_FILE, "data/orders.json", "Update orders")

def load_vip_users():
    if os.path.exists(VIP_USERS_FILE):
        with open(VIP_USERS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def save_vip_users(vip_users):
    with open(VIP_USERS_FILE, 'w', encoding='utf-8') as f:
        json.dump(vip_users, f, ensure_ascii=False, indent=4)
    upload_to_github(VIP_USERS_FILE, "data/vip_users.json", "Update VIP users")

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

def format_order_for_admin(order):
    rate_str = f"{order['rate']:.2f}".rstrip('0').rstrip('.') if isinstance(order['rate'], float) else str(order['rate'])
    
    text = f"━━━━━━━━━━━━━━━━━━━━\n"
    text += f"🆔 Order: #{order['order_id']}\n"
    text += f"⏰ Time: {order.get('order_time', 'N/A')}\n"
    text += f"⭐️ Type: {order.get('user_type', 'Normal')}\n"
    text += f"📛 Name: {order.get('user_name', 'N/A')}\n"
    text += f"🔖 Username: {order.get('username', 'N/A')}\n"
    text += f"📁 Category: {order['category']}\n"
    text += f"🔢 Quantity: {order['quantity']}\n"
    text += f"💰 Rate: {rate_str} TK\n"
    text += f"💳 Payment: {order.get('payment_method', 'N/A')}\n"
    text += f"📞 Number: {order.get('payment_number', 'N/A')}\n"
    text += f"📊 Duplicates Removed: {order.get('duplicates_removed', 0)}\n"
    if order.get('admin_note'):
        text += f"📝 Note: {order['admin_note']}\n"
    return text

def get_status_text(status):
    status_map = {
        'pending': '⏳ Pending',
        'received': '✅ Received ✅',
        'cancelled': '❌ Cancelled ❌',
        'completed': '✅ Payment Confirmed ✅'
    }
    return status_map.get(status, '⏳ Pending')

def get_main_menu_keyboard(user_id):
    is_admin = user_id in ADMIN_IDS
    keyboard = [
        ["📝 File Submit"],
        ["🆘 Admin Support"],
        ["⚙️ Settings"]
    ]
    if is_admin:
        keyboard.append(["👑 Admin Panel"])
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=False)

def get_categories_keyboard(user_id):
    categories = load_categories()
    keyboard = []
    row = []
    for cat_id, cat_info in categories.items():
        price = get_user_price(cat_id, user_id) if is_vip_user(user_id) else cat_info['price']
        price_str = f"{price:.2f}".rstrip('0').rstrip('.') if isinstance(price, float) else str(price)
        button_text = f"{cat_info['emoji']} {cat_info['name']} - {price_str} TK"
        if is_vip_user(user_id):
            button_text += " ⭐"
        row.append(button_text)
        if len(row) == 2:
            keyboard.append(row)
            row = []
    if row:
        keyboard.append(row)
    keyboard.append(["🔙 Back to Main"])
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=False)

def get_admin_panel_keyboard():
    keyboard = [
        ["📁 Manage Categories", "👑 VIP Users"],
        ["🔍 Search User", "📢 Broadcast"],
        ["📊 Report System"],
        ["🔙 Back to Main"]
    ]
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=False)

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
        show_main_menu(update, user_id)
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
        show_main_menu(update, user_id)

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
    show_main_menu(update, user_id)

def show_main_menu(update, user_id):
    keyboard = get_main_menu_keyboard(user_id)
    try:
        if hasattr(update, 'callback_query') and update.callback_query:
            update.callback_query.message.reply_text("🏠 Main Menu", reply_markup=keyboard)
        else:
            update.message.reply_text("🏠 Main Menu", reply_markup=keyboard)
    except:
        update.message.reply_text("🏠 Main Menu", reply_markup=keyboard)

def show_categories_menu(update, user_id):
    keyboard = get_categories_keyboard(user_id)
    try:
        if hasattr(update, 'callback_query') and update.callback_query:
            update.callback_query.message.reply_text("📋 Select Service:", reply_markup=keyboard)
        else:
            update.message.reply_text("📋 Select Service:", reply_markup=keyboard)
    except:
        update.message.reply_text("📋 Select Service:", reply_markup=keyboard)

def save_order_with_file(update, context, user_id, file, category_name, price, user_type):
    categories = load_categories()
    user_data = load_data()
    user_info_full = get_user_info(update, context)
    
    file_name = file.file_name
    file_id = file.file_id
    
    clean_name = clean_filename(file_name)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    unique_filename = f"{user_id}_{timestamp}_{clean_name}"
    file_path = os.path.join(UPLOAD_DIR, unique_filename)
    
    try:
        new_file = context.bot.get_file(file_id)
        new_file.download(file_path)
    except Exception as e:
        logging.error(f"File download error: {e}")
        update.message.reply_text("❌ Failed to download file. Please try again!")
        return
    
    # STEP 1: Process the file (remove duplicates within the file)
    processed_path, file_duplicates, total_count_after_file_clean = process_excel_file(file_path, user_id)
    
    if processed_path is None:
        update.message.reply_text("❌ Failed to process file. Please try again!")
        return
    
    # STEP 2: Convert Excel to Text
    text_path, text_row_count = convert_excel_to_text(processed_path)
    
    if text_path is None:
        update.message.reply_text("❌ Failed to convert file to text. Please try again!")
        return
    
    # STEP 3: Check for duplicates with existing orders in the TEXT file
    existing_texts = get_existing_texts_from_all_orders()
    cleaned_text_path, removed_count, matched_values, remaining_count = check_duplicates_in_text_file(text_path, existing_texts)
    
    # If all data was duplicate (file became empty)
    if remaining_count == 0 and removed_count > 0:
        duplicates_list = "\n".join([f"{i+1}. {val}" for i, val in enumerate(matched_values[:10])])
        if len(matched_values) > 10:
            duplicates_list += f"\n... and {len(matched_values) - 10} more"
        
        update.message.reply_text(
            f"⚠️ ALL DATA ALREADY EXISTS!\n\n"
            f"🔴 Your file contains {removed_count} entries that already exist in our system.\n\n"
            f"📋 Duplicate values:\n{duplicates_list}\n\n"
            f"💡 Please upload a file with NEW data only."
        )
        try:
            os.remove(file_path)
            os.remove(processed_path)
            os.remove(text_path)
            if cleaned_text_path != text_path:
                os.remove(cleaned_text_path)
        except:
            pass
        context.user_data.clear()
        return
    
    # If no data left after processing
    if remaining_count == 0:
        update.message.reply_text(
            "⚠️ No valid data found in your file!\n\n"
            "💡 Please upload a file with valid data."
        )
        try:
            os.remove(file_path)
            os.remove(processed_path)
            os.remove(text_path)
            if cleaned_text_path != text_path:
                os.remove(cleaned_text_path)
        except:
            pass
        context.user_data.clear()
        return
    
    # Use the cleaned text file if duplicates were removed, otherwise use original
    final_text_path = cleaned_text_path if removed_count > 0 else text_path
    
    # Count final rows
    with open(final_text_path, 'r', encoding='utf-8') as f:
        final_rows = len(f.readlines())
    
    orders = load_orders()
    order_id = str(len(orders) + 1)
    price_str = f"{price:.2f}".rstrip('0').rstrip('.') if isinstance(price, float) else str(price)
    
    # Upload files to GitHub
    github_files = {}
    
    # Upload processed Excel file
    github_excel_path = f"orders/order_{order_id}/data.xlsx"
    success, excel_url = upload_to_github(processed_path, github_excel_path, f"Add order #{order_id} Excel data")
    if success:
        github_files['excel'] = excel_url
    
    # Upload text file
    github_text_path = f"orders/order_{order_id}/data.txt"
    success, text_url = upload_to_github(final_text_path, github_text_path, f"Add order #{order_id} text data")
    if success:
        github_files['text'] = text_url
    
    # Upload original file
    github_original_path = f"orders/order_{order_id}/original_{os.path.basename(file_path)}"
    success, original_url = upload_to_github(file_path, github_original_path, f"Add order #{order_id} original file")
    if success:
        github_files['original'] = original_url
    
    order_data = {
        "order_id": order_id,
        "user_id": user_id,
        "user_name": user_info_full["name"],
        "username": user_info_full["username"],
        "order_time": user_info_full["timestamp"],
        "category": category_name,
        "quantity": final_rows,
        "rate": price,
        "file_name": file_name,
        "file_path": processed_path,
        "text_file_path": final_text_path,
        "original_file_path": file_path,
        "file_id": file_id,
        "payment_method": user_data.get(str(user_id), {}).get('payment_method'),
        "payment_number": user_data.get(str(user_id), {}).get('payment_number'),
        "status": "pending",
        "user_type": user_type,
        "duplicates_removed": len(file_duplicates),
        "admin_note": "",
        "matched_count": 0,
        "matched_data": [],
        "existing_duplicates_removed": removed_count,
        "existing_duplicates_values": matched_values,
        "github_urls": github_files
    }
    
    orders[order_id] = order_data
    save_orders(orders)
    
    current_time = datetime.now().strftime("%d-%m-%Y %I:%M %p")
    
    # Build duplicate removal message
    existing_duplicates_text = ""
    if removed_count > 0:
        existing_duplicates_text = f"\n\n🔴 Existing Data Removed ({removed_count} rows):\n"
        existing_duplicates_text += f"These values already exist in previous orders:\n"
        for i, val in enumerate(matched_values[:10], 1):
            existing_duplicates_text += f"{i}. {val}\n"
        if len(matched_values) > 10:
            existing_duplicates_text += f"... and {len(matched_values) - 10} more\n"
    else:
        existing_duplicates_text = f"\n\n✅ No existing data found!"
    
    file_duplicates_text = ""
    if file_duplicates:
        file_duplicates_text = f"\n\n🔴 Duplicates Removed from file ({len(file_duplicates)} rows):\n"
        for i, dup in enumerate(file_duplicates[:10], 1):
            file_duplicates_text += f"{i}. {dup}\n"
        if len(file_duplicates) > 10:
            file_duplicates_text += f"... and {len(file_duplicates) - 10} more\n"
        file_duplicates_text += f"\n✅ Total unique entries: {final_rows}"
    else:
        file_duplicates_text = f"\n\n✅ No duplicates found in file!\n✅ Total entries: {final_rows}"
    
    # GitHub links
    github_links = ""
    if github_files:
        github_links = "\n\n📎 **GitHub Backups:**"
        if 'excel' in github_files:
            github_links += f"\n📊 Excel: [View on GitHub]({github_files['excel']})"
        if 'text' in github_files:
            github_links += f"\n📝 Text: [View on GitHub]({github_files['text']})"
        if 'original' in github_files:
            github_links += f"\n📁 Original: [View on GitHub]({github_files['original']})"
    
    user_text = f"━━━━━━━━━━━━━━━━━━━━\n"
    user_text += f"✅ ORDER PROCESSED\n"
    user_text += f"━━━━━━━━━━━━━━━━━━━━\n\n"
    user_text += f"🆔 Order: #{order_id}\n"
    user_text += f"⏰ Time: {user_info_full['timestamp']}\n"
    user_text += f"⭐️ Type: {user_type}\n"
    user_text += f"📛 Name: {user_info_full['name']}\n"
    user_text += f"🔖 Username: {user_info_full['username']}\n"
    user_text += f"📁 Category: {category_name}\n"
    user_text += f"🔢 Quantity: {final_rows}\n"
    user_text += f"💰 Rate: {price_str} TK\n"
    user_text += f"💳 Payment: {user_data.get(str(user_id), {}).get('payment_method', 'N/A')}\n"
    user_text += f"📞 Number: {user_data.get(str(user_id), {}).get('payment_number', 'N/A')}\n"
    user_text += f"📊 Status: ⏳ Pending\n"
    user_text += f"🕐 Submitted: {current_time}"
    user_text += existing_duplicates_text
    user_text += file_duplicates_text
    user_text += github_links
    user_text += f"\n\n📎 Your file has been submitted and backed up to GitHub!"
    
    keyboard = get_main_menu_keyboard(user_id)
    update.message.reply_text(user_text, reply_markup=keyboard, parse_mode='Markdown')
    
    for admin_id in ADMIN_IDS:
        try:
            order_text = format_order_for_admin(order_data)
            keyboard_inline = [
                [InlineKeyboardButton("✅ Receive", callback_data=f"receive_{order_id}"),
                 InlineKeyboardButton("❌ Cancel", callback_data=f"cancel_{order_id}")],
                [InlineKeyboardButton("💳 Payment Complete", callback_data=f"complete_{order_id}")],
                [InlineKeyboardButton("✏️ Edit Note", callback_data=f"note_{order_id}")]
            ]
            
            with open(final_text_path, 'r', encoding='utf-8') as f:
                text_content = f.read()
            
            text_file_io = io.BytesIO(text_content.encode('utf-8'))
            text_file_io.name = f"order_{order_id}_data.txt"
            
            note_display = ""
            if order_data.get('admin_note'):
                note_display = f"\n\n📝 Note: {order_data['admin_note']}"
            
            github_backup = ""
            if github_files:
                github_backup = f"\n\n📎 GitHub Backups:"
                for key, url in github_files.items():
                    github_backup += f"\n• {key}: {url}"
            
            caption = f"📦 New Order #{order_id}\n\n{order_text}\n\n📊 Existing Duplicates Removed: {removed_count}\n📊 File Duplicates Removed: {len(file_duplicates)}\n📁 Data: {final_rows} rows (cleaned){note_display}{github_backup}"
            
            context.bot.send_document(
                chat_id=admin_id,
                document=text_file_io,
                filename=f"order_{order_id}_data.txt",
                caption=caption,
                reply_markup=InlineKeyboardMarkup(keyboard_inline)
            )
            
        except Exception as e:
            logging.error(f"Error sending text file to admin {admin_id}: {e}")
    
    context.user_data.clear()

def handle_file_upload(update, context):
    user_id = update.effective_user.id
    file = update.message.document
    
    file_name = file.file_name.lower()
    
    if user_id in ADMIN_IDS and context.user_data.get('awaiting_report'):
        if file_name.endswith('.txt'):
            handle_text_report_upload(update, context)
        else:
            update.message.reply_text("❌ Please upload a .txt file for report!")
        return
    
    if not file_name.endswith(('.xlsx', '.xls')):
        update.message.reply_text("❌ Please upload a valid Excel file (.xlsx or .xls)!")
        return
    
    if 'selected_category' not in context.user_data:
        update.message.reply_text("❌ Please select a category first from the menu!")
        return
    
    cat_id = context.user_data.get('selected_category')
    categories = load_categories()
    category = categories.get(cat_id)
    
    if not category:
        update.message.reply_text("❌ Category not found! Please select again.")
        return
    
    price = get_user_price(cat_id, user_id)
    user_type = "VIP" if is_vip_user(user_id) else "Normal"
    
    save_order_with_file(update, context, user_id, file, category['name'], price, user_type)

def handle_text_report_upload(update, context):
    user_id = update.effective_user.id
    file = update.message.document
    
    try:
        new_file = context.bot.get_file(file.file_id)
        file_content = new_file.download_as_bytearray()
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_filename = f"report_{user_id}_{timestamp}.txt"
        report_path = os.path.join(REPORT_DIR, report_filename)
        
        with open(report_path, 'wb') as f:
            f.write(file_content)
        
        # Upload report to GitHub
        github_report_path = f"reports/report_{timestamp}.txt"
        upload_to_github(report_path, github_report_path, f"Add report {timestamp}")
        
        update.message.reply_text("⏳ Processing report... Please wait.")
        
        result = process_text_report(update, context, report_path, user_id)
        
        if result:
            report_id, matched_orders, total_report_entries = result
            
            if matched_orders:
                report_result = {
                    "report_id": report_id,
                    "report_file": report_filename,
                    "uploaded_by": user_id,
                    "upload_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "total_report_entries": total_report_entries,
                    "matched_orders_count": len(matched_orders),
                    "matched_orders": matched_orders
                }
                
                result_file = os.path.join(REPORT_RESULTS_DIR, f"result_{report_id}.json")
                with open(result_file, 'w', encoding='utf-8') as f:
                    json.dump(report_result, f, ensure_ascii=False, indent=4)
                
                # Upload result to GitHub
                github_result_path = f"reports/results/result_{report_id}.json"
                upload_to_github(result_file, github_result_path, f"Add report result {report_id}")
            
            # Continue with report processing
            if matched_orders:
                update.message.reply_text(
                    f"✅ Report Processed!\n\n"
                    f"📊 Total Report Entries: {total_report_entries}\n"
                    f"✅ Orders with Matches: {len(matched_orders)}\n\n"
                    f"📋 Orders with matches are shown below:"
                )
                
                for result in matched_orders:
                    order_id = result['order_id']
                    orders = load_orders()
                    order = orders.get(order_id, {})
                    
                    if order:
                        order_text = format_order_for_admin(order)
                        current_time = datetime.now().strftime("%d-%m-%Y %I:%M %p")
                        
                        matched_count = result['matched_count']
                        rate = order.get('rate', 0)
                        total_price = matched_count * rate
                        
                        status = order.get('status', 'pending')
                        if status == 'received':
                            receive_btn = InlineKeyboardButton("✅ Received ✅", callback_data="done")
                        else:
                            receive_btn = InlineKeyboardButton("✅ Receive", callback_data=f"receive_{order_id}")
                        
                        if status == 'cancelled':
                            cancel_btn = InlineKeyboardButton("❌ Cancelled ❌", callback_data="done")
                        else:
                            cancel_btn = InlineKeyboardButton("❌ Cancel", callback_data=f"cancel_{order_id}")
                        
                        if status == 'completed':
                            complete_btn = InlineKeyboardButton("✅ Payment Confirmed ✅", callback_data="done")
                        else:
                            complete_btn = InlineKeyboardButton("💳 Payment Complete", callback_data=f"complete_{order_id}")
                        
                        keyboard_inline = [
                            [receive_btn, cancel_btn],
                            [complete_btn],
                            [InlineKeyboardButton("✏️ Edit Note", callback_data=f"note_{order_id}")]
                        ]
                        
                        try:
                            text_file_path = order.get('text_file_path')
                            if text_file_path and os.path.exists(text_file_path):
                                with open(text_file_path, 'r', encoding='utf-8') as f:
                                    text_content = f.read()
                                
                                text_file_io = io.BytesIO(text_content.encode('utf-8'))
                                text_file_io.name = f"order_{order_id}_data.txt"
                                
                                caption = (
                                    f"📦 Order #{order_id} (Matched in Report)\n"
                                    f"🕐 {current_time}\n\n"
                                    f"{order_text}\n\n"
                                    f"📊 Status: {get_status_text(status)}\n"
                                    f"✅ Matched: {matched_count} entries\n"
                                    f"💰 Rate: {rate} TK\n"
                                    f"💵 Total Price: {total_price} TK"
                                )
                                
                                update.message.reply_document(
                                    document=text_file_io,
                                    filename=f"order_{order_id}_data.txt",
                                    caption=caption,
                                    reply_markup=InlineKeyboardMarkup(keyboard_inline)
                                )
                            else:
                                caption = (
                                    f"📦 Order #{order_id} (Matched in Report)\n"
                                    f"🕐 {current_time}\n\n"
                                    f"{order_text}\n\n"
                                    f"📊 Status: {get_status_text(status)}\n"
                                    f"✅ Matched: {matched_count} entries\n"
                                    f"💰 Rate: {rate} TK\n"
                                    f"💵 Total Price: {total_price} TK"
                                )
                                
                                update.message.reply_text(
                                    caption,
                                    reply_markup=InlineKeyboardMarkup(keyboard_inline)
                                )
                        except Exception as e:
                            logging.error(f"Error sending matched order {order_id}: {e}")
                            caption = f"📦 Order #{order_id} (Matched in Report)\n\n{order_text}\n\n📊 Status: {get_status_text(status)}"
                            update.message.reply_text(
                                caption,
                                reply_markup=InlineKeyboardMarkup(keyboard_inline)
                            )
            else:
                update.message.reply_text(
                    f"✅ Report Processed!\n\n"
                    f"📊 Total Report Entries: {total_report_entries}\n"
                    f"❌ No matches found in any order.\n\n"
                    f"💡 Tip: Make sure the text file contains data that matches any user's order file content."
                )
        else:
            update.message.reply_text("❌ Failed to process report file!")
        
        context.user_data['awaiting_report'] = False
        
    except Exception as e:
        logging.error(f"Error in handle_text_report_upload: {e}")
        update.message.reply_text(f"❌ Error processing report: {str(e)}")
        context.user_data['awaiting_report'] = False

def process_text_report(update, context, report_path, admin_id):
    """Process text file report and match with all orders"""
    try:
        with open(report_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        report_values = []
        for line in content.split('\n'):
            if line.strip():
                parts = re.split(r'[\t,\s]+', line.strip())
                for part in parts:
                    if part.strip():
                        report_values.append(part.strip())
        
        if not report_values:
            update.message.reply_text("❌ Report file is empty!")
            return None
        
        report_set = set(report_values)
        total_report_entries = len(report_set)
        
        orders = load_orders()
        matched_orders = []
        
        for order_id, order in orders.items():
            text_file_path = order.get('text_file_path')
            if not text_file_path or not os.path.exists(text_file_path):
                continue
            
            try:
                with open(text_file_path, 'r', encoding='utf-8') as f:
                    order_content = f.read()
            except Exception as e:
                logging.error(f"Error reading order file {text_file_path}: {e}")
                continue
            
            order_values = []
            for line in order_content.split('\n'):
                if line.strip():
                    parts = line.strip().split('\t')
                    for part in parts:
                        if part.strip():
                            order_values.append(part.strip())
            
            if not order_values:
                continue
            
            order_counter = Counter(order_values)
            matched_values = []
            
            for val in report_set:
                if val in order_counter:
                    matched_values.append(val)
            
            if len(matched_values) >= 1:
                matched_count = len(matched_values)
                rate = order.get('rate', 0)
                total_price = rate * matched_count
                
                matched_orders.append({
                    "order_id": order_id,
                    "user_name": order.get('user_name', 'N/A'),
                    "username": order.get('username', 'N/A'),
                    "category": order.get('category', 'N/A'),
                    "quantity": order.get('quantity', 0),
                    "matched_count": matched_count,
                    "rate": rate,
                    "total_price": total_price,
                    "order_time": order.get('order_time', 'N/A'),
                    "status": order.get('status', 'pending'),
                    "text_file_path": text_file_path,
                    "order_data": order
                })
        
        matched_orders.sort(key=lambda x: x['matched_count'], reverse=True)
        
        report_id = hashlib.md5(f"{admin_id}_{datetime.now().isoformat()}".encode()).hexdigest()[:8]
        
        return report_id, matched_orders, total_report_entries
        
    except Exception as e:
        logging.error(f"Error processing text report: {e}")
        update.message.reply_text(f"❌ Error processing report: {str(e)}")
        return None

def show_support_menu(update, user_id):
    keyboard = get_main_menu_keyboard(user_id)
    update.message.reply_text(
        "🆘 Admin Support\n\n"
        "For any assistance, please contact:\n"
        "👤 @Rubel_QSB\n\n"
        "Or send a message to our support team.",
        reply_markup=keyboard
    )

def show_setting_menu(update, user_id):
    user_data = load_data()
    pm = user_data.get(str(user_id), {})
    vip_status = "✅ VIP Member ⭐" if is_vip_user(user_id) else "Normal User"
    
    keyboard = [
        ["💳 Change Payment"],
        ["🔙 Back to Main"]
    ]
    reply_keyboard = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=False)
    
    update.message.reply_text(
        f"⚙️ Settings\n\n"
        f"👤 {vip_status}\n\n"
        f"💳 Payment: {pm.get('payment_method', 'N/A')}\n"
        f"📞 Number: {pm.get('payment_number', 'N/A')}",
        reply_markup=reply_keyboard
    )

def show_admin_panel(update, user_id):
    keyboard = get_admin_panel_keyboard()
    update.message.reply_text("👑 Admin Panel", reply_markup=keyboard)

def show_admin_categories(update, user_id):
    categories = load_categories()
    keyboard = []
    row = []
    for cat_id, cat_info in categories.items():
        price_str = f"{cat_info['price']:.2f}".rstrip('0').rstrip('.') if isinstance(cat_info['price'], float) else str(cat_info['price'])
        short_id = get_short_id(cat_id, category_short_map)
        button_text = f"✏️ {cat_info['emoji']} {cat_info['name']} - {price_str} TK"
        row.append(button_text)
        if len(row) == 2:
            keyboard.append(row)
            row = []
    if row:
        keyboard.append(row)
    keyboard.append(["➕ Add New Category"])
    keyboard.append(["🔙 Back to Admin Panel"])
    reply_keyboard = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=False)
    
    update.message.reply_text(
        "📁 Manage Categories\n\n"
        "✨ Click on a category to edit\n"
        "💰 Decimal rates supported\n\n"
        "➕ Add New Category - to add new",
        reply_markup=reply_keyboard
    )

def show_admin_vip(update, user_id):
    keyboard = [
        ["➕ Add VIP", "📋 VIP List"],
        ["✏️ Edit Rates", "🗑 Remove VIP"],
        ["🔙 Back to Admin Panel"]
    ]
    reply_keyboard = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=False)
    update.message.reply_text("👑 VIP User Management", reply_markup=reply_keyboard)

def show_report_system(update, user_id):
    keyboard = [
        ["📤 Upload Report"],
        ["📋 View Reports"],
        ["🔙 Back to Admin Panel"]
    ]
    reply_keyboard = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=False)
    update.message.reply_text(
        "📊 Report System\n\n"
        "Upload a .txt file to match with user orders.\n"
        "The system will compare ALL text in the report file\n"
        "with ALL text in each order file.\n\n"
        "✅ Each unique match counts as 1\n"
        "💵 Total = Matched × Rate\n\n"
        "You can manage orders directly from results.",
        reply_markup=reply_keyboard
    )

def start_report_upload(update, context):
    context.user_data['awaiting_report'] = True
    keyboard = [
        ["📤 Upload Report"],
        ["📋 View Reports"],
        ["🔙 Back to Admin Panel"]
    ]
    reply_keyboard = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=False)
    update.message.reply_text(
        "📤 Upload Report File\n\n"
        "Please upload a .txt file containing the report data.\n\n"
        "The system will compare ALL text from the report file\n"
        "with ALL text from each order file.\n"
        "✅ Each unique match counts as 1\n"
        "💵 Total = Matched × Rate\n\n"
        "⚠️ Only .txt files are accepted.",
        reply_markup=reply_keyboard
    )

def view_reports(update, user_id):
    report_files = os.listdir(REPORT_RESULTS_DIR)
    keyboard = [
        ["📤 Upload Report"],
        ["📋 View Reports"],
        ["🔙 Back to Admin Panel"]
    ]
    reply_keyboard = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=False)
    
    if not report_files:
        update.message.reply_text("📋 No reports found.", reply_markup=reply_keyboard)
        return
    
    text = "📋 Recent Reports:\n\n"
    for file in sorted(report_files, reverse=True)[:10]:
        try:
            with open(os.path.join(REPORT_RESULTS_DIR, file), 'r', encoding='utf-8') as f:
                data = json.load(f)
            report_id = data.get('report_id', 'Unknown')
            matched = data.get('matched_orders_count', 0)
            total = data.get('total_report_entries', 0)
            time = data.get('upload_time', 'Unknown')
            text += f"📊 Report #{report_id}\n"
            text += f"   Matched: {matched} orders\n"
            text += f"   Entries: {total}\n"
            text += f"   Time: {time}\n\n"
        except:
            continue
    
    update.message.reply_text(text, reply_markup=reply_keyboard)

def show_vip_list(update, user_id):
    vip_users = load_vip_users()
    keyboard = [
        ["➕ Add VIP", "📋 VIP List"],
        ["✏️ Edit Rates", "🗑 Remove VIP"],
        ["🔙 Back to Admin Panel"]
    ]
    reply_keyboard = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=False)
    
    if not vip_users:
        update.message.reply_text("📋 No VIP users found.", reply_markup=reply_keyboard)
        return
    
    text = "📋 VIP List:\n\n"
    for uid, info in vip_users.items():
        text += f"━━━━━━━━━━━━━━━━━━━━\n"
        text += f"🆔 ID: {uid}\n"
        text += f"👤 Name: {info.get('name', 'N/A')}\n"
        text += f"📅 Added: {info.get('added_time', 'N/A')}\n"
        if info.get('custom_rates'):
            text += f"⭐ Custom rates:\n"
            for cat_id, r in info['custom_rates'].items():
                text += f"   • {cat_id}: {r} TK\n"
        text += "\n"
    
    if len(text) > 4000:
        text = text[:4000] + "\n\n... (truncated)"
    
    update.message.reply_text(text, reply_markup=reply_keyboard)

def edit_vip_rates_list(update, context):
    vip_users = load_vip_users()
    keyboard = [
        ["➕ Add VIP", "📋 VIP List"],
        ["✏️ Edit Rates", "🗑 Remove VIP"],
        ["🔙 Back to Admin Panel"]
    ]
    reply_keyboard = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=False)
    
    if not vip_users:
        update.message.reply_text("No VIP users found.", reply_markup=reply_keyboard)
        return
    
    text = "👑 Select a VIP user to edit rates:\n\n"
    for uid, info in vip_users.items():
        name = info.get('name', 'Unknown')
        text += f"🆔 {uid} - {name}\n"
        text += f"   Send: /edit_vip_{uid}\n\n"
    
    update.message.reply_text(text, reply_markup=reply_keyboard)

def remove_vip_start(update, context):
    vip_users = load_vip_users()
    keyboard = [
        ["➕ Add VIP", "📋 VIP List"],
        ["✏️ Edit Rates", "🗑 Remove VIP"],
        ["🔙 Back to Admin Panel"]
    ]
    reply_keyboard = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=False)
    
    if not vip_users:
        update.message.reply_text("No VIP users found.", reply_markup=reply_keyboard)
        return
    
    text = "🗑 Select a VIP user to remove:\n\n"
    for uid, info in vip_users.items():
        name = info.get('name', 'Unknown')
        text += f"🆔 {uid} - {name}\n"
        text += f"   Send: /remove_vip_{uid}\n\n"
    
    update.message.reply_text(text, reply_markup=reply_keyboard)

def add_vip_start(update, context):
    keyboard = [
        ["➕ Add VIP", "📋 VIP List"],
        ["✏️ Edit Rates", "🗑 Remove VIP"],
        ["🔙 Back to Admin Panel"]
    ]
    reply_keyboard = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=False)
    update.message.reply_text(
        "➕ Add VIP\n\n"
        "Send User ID to add as VIP:\n"
        "Example: `8555327754`",
        reply_markup=reply_keyboard
    )

def process_add_vip(update, context, user_id_text):
    if not user_id_text.isdigit():
        update.message.reply_text("❌ Send numeric User ID!")
        return
    
    vip_users = load_vip_users()
    if user_id_text in vip_users:
        update.message.reply_text(f"❌ User {user_id_text} is already VIP!")
        return
    
    vip_users[user_id_text] = {
        "user_id": user_id_text,
        "name": "Unknown",
        "username": "Unknown",
        "added_by": str(update.effective_user.id),
        "added_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "custom_rates": {}
    }
    save_vip_users(vip_users)
    update.message.reply_text(f"✅ VIP Added: {user_id_text}")
    show_vip_rates_menu(update, context, user_id_text)

def show_vip_rates_menu(update, context, user_id_text):
    categories = load_categories()
    keyboard = []
    row = []
    for cat_id, cat_info in categories.items():
        button_text = f"{cat_info['emoji']} {cat_info['name']}"
        row.append(button_text)
        if len(row) == 2:
            keyboard.append(row)
            row = []
    if row:
        keyboard.append(row)
    keyboard.append(["🔙 Back to VIP Management"])
    reply_keyboard = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=False)
    
    context.user_data['vip_rate_user'] = user_id_text
    update.message.reply_text(
        f"👑 Set Rates for VIP {user_id_text}\n\n"
        f"Click on a category to set custom rate:",
        reply_markup=reply_keyboard
    )

def ask_vip_rate(update, context, cat_name):
    categories = load_categories()
    cat_id = None
    for cid, info in categories.items():
        if info['name'].lower() == cat_name.lower():
            cat_id = cid
            break
    
    if not cat_id:
        update.message.reply_text("❌ Category not found!")
        return
    
    context.user_data['vip_rate_cat'] = cat_id
    update.message.reply_text(
        f"Set rate for {cat_name}\n"
        f"Default: {categories[cat_id]['price']} TK\n\n"
        f"Send custom rate (e.g., 30, 7.5):"
    )

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
        show_vip_rates_menu(update, context, user_id_text)
        
    except ValueError:
        update.message.reply_text("❌ Send valid number! (e.g., 30, 7.5)")

def process_remove_vip(update, context, user_id_text):
    if not user_id_text.isdigit():
        update.message.reply_text("❌ Send numeric User ID!")
        return
    
    vip_users = load_vip_users()
    if user_id_text in vip_users:
        del vip_users[user_id_text]
        save_vip_users(vip_users)
        update.message.reply_text(f"✅ VIP removed: {user_id_text}")
    else:
        update.message.reply_text(f"❌ {user_id_text} not in VIP list")
    
    show_admin_vip(update, update.effective_user.id)

def search_user_start(update, context):
    context.user_data['search_user'] = True
    keyboard = get_admin_panel_keyboard()
    update.message.reply_text(
        "🔍 Search User Orders\n\n"
        "Send username to search (e.g., @username or username):",
        reply_markup=keyboard
    )

def search_user_orders(update, context, search_term):
    search_term = search_term.strip()
    if search_term.startswith('@'):
        search_term = search_term[1:]
    
    orders = load_orders()
    found_orders = []
    
    for order_id, order in orders.items():
        username = order.get('username', '').replace('@', '').lower()
        if search_term.lower() in username:
            found_orders.append(order)
    
    keyboard = get_admin_panel_keyboard()
    
    if not found_orders:
        update.message.reply_text(
            f"❌ No orders found for username: @{search_term}\n\n"
            f"Please check the username and try again.",
            reply_markup=keyboard
        )
        context.user_data['search_user'] = False
        return
    
    result_text = f"🔍 Search Results for: @{search_term}\n"
    result_text += f"📊 Found {len(found_orders)} order(s)\n\n"
    
    for order in found_orders:
        result_text += format_order_for_admin(order)
        result_text += f"📊 Status: {get_status_text(order['status'])}\n\n"
    
    if len(result_text) > 4000:
        first_part = result_text[:3800] + "\n\n... (More orders exist)"
        update.message.reply_text(first_part, reply_markup=keyboard)
        remaining = result_text[3800:]
        chunks = [remaining[i:i+4000] for i in range(0, len(remaining), 4000)]
        for chunk in chunks:
            update.message.reply_text(chunk)
    else:
        update.message.reply_text(result_text, reply_markup=keyboard)
    
    context.user_data['search_user'] = False

def broadcast_start(update, context):
    context.user_data['broadcast'] = True
    keyboard = get_admin_panel_keyboard()
    update.message.reply_text(
        "📢 Send broadcast message to ALL users:\n\n"
        "Type your message below:",
        reply_markup=keyboard
    )

def send_broadcast(update, context, message_text):
    user_data = load_data()
    sent = 0
    failed = 0
    
    for user_id_str in user_data.keys():
        try:
            context.bot.send_message(chat_id=int(user_id_str), text=message_text)
            sent += 1
        except Exception as e:
            failed += 1
            logging.error(f"Failed to send broadcast to {user_id_str}: {e}")
    
    keyboard = get_admin_panel_keyboard()
    response_text = f"✅ Broadcast Sent!\n\n📊 Sent to: {sent} users\n❌ Failed: {failed} users"
    update.message.reply_text(response_text, reply_markup=keyboard)
    context.user_data['broadcast'] = False

def show_edit_category(update, context, cat_name):
    categories = load_categories()
    cat_id = None
    for cid, info in categories.items():
        if info['name'].lower() == cat_name.lower():
            cat_id = cid
            break
    
    if not cat_id:
        update.message.reply_text("❌ Category not found!")
        return
    
    category = categories[cat_id]
    price_str = f"{category['price']:.2f}".rstrip('0').rstrip('.') if isinstance(category['price'], float) else str(category['price'])
    
    keyboard = [
        ["📝 Change Name", "💰 Change Price"],
        ["🗑 Delete Category"],
        ["🔙 Back to Categories"]
    ]
    reply_keyboard = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=False)
    
    context.user_data['editing_category'] = cat_id
    
    update.message.reply_text(
        f"✏️ {category['emoji']} {category['name']}\n"
        f"💰 Price: {price_str} TK\n\n"
        f"Select an option:",
        reply_markup=reply_keyboard
    )

def add_category_start(update, context):
    context.user_data['add_cat'] = 'name'
    keyboard = [
        ["🔙 Back to Admin Panel"]
    ]
    reply_keyboard = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=False)
    update.message.reply_text(
        "➕ Add Category\n\n"
        "📝 Step 1/2: Send category name\n\n"
        "Example: `Twitter` or `Hotmail`",
        reply_markup=reply_keyboard
    )

def process_cat_name(update, context, name):
    context.user_data['new_cat_name'] = name.strip()
    context.user_data['add_cat'] = 'price'
    keyboard = [
        ["🔙 Back to Admin Panel"]
    ]
    reply_keyboard = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=False)
    update.message.reply_text(
        f"📝 Name: {name.strip()}\n\n"
        f"💰 Step 2/2: Send rate (TK)\n\n"
        f"Example: `30` or `7.5`",
        reply_markup=reply_keyboard
    )

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
        show_admin_categories(update, update.effective_user.id)
        
    except ValueError:
        update.message.reply_text("❌ Send valid number! (e.g., 30, 7.5)")

def change_category_name(update, context, cat_id):
    context.user_data['ren_cat'] = cat_id
    keyboard = [
        ["🔙 Back to Categories"]
    ]
    reply_keyboard = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=False)
    update.message.reply_text("📝 Send new name:", reply_markup=reply_keyboard)

def process_category_rename(update, context, new_name):
    cat_id = context.user_data.get('ren_cat')
    if not cat_id:
        update.message.reply_text("❌ Session expired!")
        return
    
    categories = load_categories()
    if cat_id in categories:
        new_emoji = get_emoji_for_category(new_name)
        categories[cat_id]['name'] = new_name
        categories[cat_id]['emoji'] = new_emoji
        save_categories(categories)
        update.message.reply_text(f"✅ Name changed to: {new_name}\n✨ Emoji: {new_emoji}")
    
    context.user_data['ren_cat'] = None
    show_admin_categories(update, update.effective_user.id)

def change_category_price(update, context, cat_id):
    context.user_data['prc_cat'] = cat_id
    keyboard = [
        ["🔙 Back to Categories"]
    ]
    reply_keyboard = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=False)
    update.message.reply_text("💰 Send new price (e.g., 30, 7.5):", reply_markup=reply_keyboard)

def process_category_price(update, context, price_text):
    cat_id = context.user_data.get('prc_cat')
    if not cat_id:
        update.message.reply_text("❌ Session expired!")
        return
    
    try:
        price = float(price_text) if '.' in price_text else int(price_text)
        if price <= 0:
            update.message.reply_text("Enter positive number!")
            return
        
        categories = load_categories()
        if cat_id in categories:
            categories[cat_id]['price'] = price
            save_categories(categories)
            price_str = f"{price:.2f}".rstrip('0').rstrip('.') if isinstance(price, float) else str(price)
            update.message.reply_text(f"✅ Price: {price_str} TK")
        
        context.user_data['prc_cat'] = None
        show_admin_categories(update, update.effective_user.id)
        
    except ValueError:
        update.message.reply_text("Send valid number! (e.g., 30, 7.5)")

def delete_category(update, context, cat_id):
    categories = load_categories()
    if cat_id in categories:
        del categories[cat_id]
        save_categories(categories)
        update.message.reply_text("✅ Deleted!")
    
    show_admin_categories(update, update.effective_user.id)

def change_payment_start(update, context):
    keyboard = [
        ["bKash", "Nagad"],
        ["Rocket", "Binance"],
        ["🔙 Back to Settings"]
    ]
    reply_keyboard = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=False)
    update.message.reply_text("Select new payment method:", reply_markup=reply_keyboard)

def change_payment_number(update, context, method):
    context.user_data['change_payment_method'] = method
    keyboard = [
        ["🔙 Back to Settings"]
    ]
    reply_keyboard = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=False)
    update.message.reply_text(f"Enter your {method} number:", reply_markup=reply_keyboard)

def save_new_payment(update, context):
    user_id = update.effective_user.id
    number = update.message.text.strip()
    method = context.user_data.get('change_payment_method')
    
    if not method:
        update.message.reply_text("❌ Session expired!")
        return
    
    user_data = load_data()
    user_data[str(user_id)]["payment_method"] = method
    user_data[str(user_id)]["payment_number"] = number
    save_data(user_data)
    
    context.user_data['change_payment_method'] = None
    update.message.reply_text(f"✅ Payment updated!\n\n💳 {method}\n📞 {number}")
    show_main_menu(update, user_id)

# ================ BUTTON CALLBACK ================

def button_callback(update, context):
    query = update.callback_query
    query.answer()
    user_id = update.effective_user.id
    data = query.data
    
    if user_id not in ADMIN_IDS:
        if not is_member(context.bot, user_id):
            query.edit_message_text("Please join the channel!")
            return
    
    if data.startswith("note_") and user_id in ADMIN_IDS:
        order_id = data.replace("note_", "")
        edit_admin_note(update, context, order_id)
        return
    
    if data.startswith("back_to_order_"):
        order_id = data.replace("back_to_order_", "")
        orders = load_orders()
        if order_id in orders:
            order = orders[order_id]
            order_text = format_order_for_admin(order)
            
            status = order['status']
            if status == 'received':
                receive_btn = InlineKeyboardButton("✅ Received ✅", callback_data="done")
            else:
                receive_btn = InlineKeyboardButton("✅ Receive", callback_data=f"receive_{order_id}")
            
            if status == 'cancelled':
                cancel_btn = InlineKeyboardButton("❌ Cancelled ❌", callback_data="done")
            else:
                cancel_btn = InlineKeyboardButton("❌ Cancel", callback_data=f"cancel_{order_id}")
            
            if status == 'completed':
                complete_btn = InlineKeyboardButton("✅ Payment Confirmed ✅", callback_data="done")
            else:
                complete_btn = InlineKeyboardButton("💳 Payment Complete", callback_data=f"complete_{order_id}")
            
            keyboard_inline = [
                [receive_btn, cancel_btn],
                [complete_btn],
                [InlineKeyboardButton("✏️ Edit Note", callback_data=f"note_{order_id}")]
            ]
            
            try:
                query.message.delete()
                query.message.reply_text(
                    f"📦 Order #{order_id}\n\n{order_text}\n\n📊 Status: {get_status_text(order['status'])}",
                    reply_markup=InlineKeyboardMarkup(keyboard_inline)
                )
            except:
                query.message.reply_text(
                    f"📦 Order #{order_id}\n\n{order_text}\n\n📊 Status: {get_status_text(order['status'])}",
                    reply_markup=InlineKeyboardMarkup(keyboard_inline)
                )
        context.user_data['editing_note_order'] = None
        return
    
    if data.startswith("receive_") and user_id in ADMIN_IDS:
        order_id = data.replace("receive_", "")
        handle_receive_order(update, context, order_id)
        return
    elif data.startswith("cancel_") and user_id in ADMIN_IDS:
        order_id = data.replace("cancel_", "")
        handle_cancel_order(update, context, order_id)
        return
    elif data.startswith("complete_") and user_id in ADMIN_IDS:
        order_id = data.replace("complete_", "")
        handle_complete_order(update, context, order_id)
        return
    
    if data == "done":
        return
    
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
        query.message.reply_text(f"Enter your {method} number:")
        context.user_data['waiting_for_payment_number'] = True

# ================ ORDER MANAGEMENT FUNCTIONS ================

def edit_admin_note(update, context, order_id):
    orders = load_orders()
    
    if order_id not in orders:
        update.callback_query.answer("❌ Order not found!", show_alert=True)
        return
    
    context.user_data['editing_note_order'] = order_id
    
    current_note = orders[order_id].get('admin_note', '')
    
    if current_note:
        note_display = f"📝 Current Note: {current_note}"
    else:
        note_display = "📝 No note added yet"
    
    keyboard_inline = [
        [InlineKeyboardButton("🔙 Back to Order", callback_data=f"back_to_order_{order_id}")]
    ]
    
    try:
        update.callback_query.message.reply_text(
            f"✏️ Edit Note for Order #{order_id}\n\n"
            f"{note_display}\n\n"
            f"━━━━━━━━━━━━━━━━━━━━\n"
            f"📝 Type your note below and press Send:\n"
            f"(Note will only be visible to Admins)",
            reply_markup=InlineKeyboardMarkup(keyboard_inline)
        )
        update.callback_query.answer()
    except Exception as e:
        logging.error(f"Error in edit_admin_note: {e}")
        update.callback_query.message.reply_text(
            f"✏️ Edit Note for Order #{order_id}\n\n"
            f"Type your note and press Send:",
            reply_markup=InlineKeyboardMarkup(keyboard_inline)
        )

def save_admin_note(update, context, order_id, note_text):
    orders = load_orders()
    
    if order_id not in orders:
        update.message.reply_text("❌ Order not found!")
        return
    
    orders[order_id]['admin_note'] = note_text.strip()
    save_orders(orders)
    
    order = orders[order_id]
    
    status = order['status']
    if status == 'received':
        receive_btn = InlineKeyboardButton("✅ Received ✅", callback_data="done")
    else:
        receive_btn = InlineKeyboardButton("✅ Receive", callback_data=f"receive_{order_id}")
    
    if status == 'cancelled':
        cancel_btn = InlineKeyboardButton("❌ Cancelled ❌", callback_data="done")
    else:
        cancel_btn = InlineKeyboardButton("❌ Cancel", callback_data=f"cancel_{order_id}")
    
    if status == 'completed':
        complete_btn = InlineKeyboardButton("✅ Payment Confirmed ✅", callback_data="done")
    else:
        complete_btn = InlineKeyboardButton("💳 Payment Complete", callback_data=f"complete_{order_id}")
    
    keyboard_inline = [
        [receive_btn, cancel_btn],
        [complete_btn],
        [InlineKeyboardButton("✏️ Edit Note", callback_data=f"note_{order_id}")]
    ]
    
    order_text = format_order_for_admin(order)
    
    update.message.reply_text(
        f"✅ Note saved successfully for Order #{order_id}!\n\n"
        f"📦 Order #{order_id}\n\n"
        f"{order_text}\n\n"
        f"📊 Status: {get_status_text(order['status'])}",
        reply_markup=InlineKeyboardMarkup(keyboard_inline)
    )
    
    context.user_data['editing_note_order'] = None

def update_order_status(update, context, order_id, new_status, button_text, user_message):
    user_id = update.effective_user.id
    if user_id not in ADMIN_IDS:
        return False
    
    orders = load_orders()
    if order_id not in orders:
        if update.callback_query:
            update.callback_query.answer("❌ Order not found!", show_alert=True)
        return False
    
    orders[order_id]['status'] = new_status
    save_orders(orders)
    order = orders[order_id]
    
    status = order['status']
    
    if status == 'received':
        receive_btn = InlineKeyboardButton("✅ Received ✅", callback_data="done")
    else:
        receive_btn = InlineKeyboardButton("✅ Receive", callback_data=f"receive_{order_id}")
    
    if status == 'cancelled':
        cancel_btn = InlineKeyboardButton("❌ Cancelled ❌", callback_data="done")
    else:
        cancel_btn = InlineKeyboardButton("❌ Cancel", callback_data=f"cancel_{order_id}")
    
    if status == 'completed':
        complete_btn = InlineKeyboardButton("✅ Payment Confirmed ✅", callback_data="done")
    else:
        complete_btn = InlineKeyboardButton("💳 Payment Complete", callback_data=f"complete_{order_id}")
    
    keyboard_inline = [
        [receive_btn, cancel_btn],
        [complete_btn],
        [InlineKeyboardButton("✏️ Edit Note", callback_data=f"note_{order_id}")]
    ]
    
    try:
        order_text = format_order_for_admin(order)
        current_time = datetime.now().strftime("%d-%m-%Y %I:%M %p")
        
        if update.callback_query:
            try:
                update.callback_query.edit_message_caption(
                    caption=f"📦 Order #{order_id}\n🕐 Updated: {current_time}\n\n{order_text}\n\n📊 Status: {get_status_text(status)}",
                    reply_markup=InlineKeyboardMarkup(keyboard_inline)
                )
            except:
                update.callback_query.message.reply_text(
                    f"📦 Order #{order_id}\n🕐 Updated: {current_time}\n\n{order_text}\n\n📊 Status: {get_status_text(status)}",
                    reply_markup=InlineKeyboardMarkup(keyboard_inline)
                )
            update.callback_query.answer(f"{button_text}", show_alert=True)
        else:
            update.message.reply_text(
                f"📦 Order #{order_id}\n🕐 Updated: {current_time}\n\n{order_text}\n\n📊 Status: {get_status_text(status)}",
                reply_markup=InlineKeyboardMarkup(keyboard_inline)
            )
    except Exception as e:
        logging.error(f"Error updating admin message: {e}")
        if update.callback_query:
            update.callback_query.message.reply_text(f"✅ Order #{order_id} - {button_text}")
        else:
            update.message.reply_text(f"✅ Order #{order_id} - {button_text}")
    
    try:
        user_id = order['user_id']
        user_text = f"━━━━━━━━━━━━━━━━━━━━\n"
        user_text += f"📋 ORDER UPDATE\n"
        user_text += f"━━━━━━━━━━━━━━━━━━━━\n\n"
        user_text += f"🆔 Order: #{order['order_id']}\n"
        user_text += f"⏰ Time: {order.get('order_time', 'N/A')}\n"
        user_text += f"⭐️ Type: {order.get('user_type', 'Normal')}\n"
        user_text += f"📛 Name: {order.get('user_name', 'N/A')}\n"
        user_text += f"🔖 Username: {order.get('username', 'N/A')}\n"
        user_text += f"📁 Category: {order['category']}\n"
        user_text += f"🔢 Quantity: {order['quantity']}\n"
        user_text += f"💰 Rate: {order['rate']} TK\n"
        user_text += f"💳 Payment: {order.get('payment_method', 'N/A')}\n"
        user_text += f"📞 Number: {order.get('payment_number', 'N/A')}\n"
        user_text += f"📊 Status: {get_status_text(status)}\n"
        user_text += f"🕐 Updated: {current_time}\n\n"
        user_text += f"{user_message}\n"
        user_text += f"━━━━━━━━━━━━━━━━━━━━\n"
        user_text += f"Thank you for using our service! 🙏"
        
        context.bot.send_message(
            chat_id=user_id,
            text=user_text
        )
    except Exception as e:
        logging.error(f"Error sending user notification: {e}")
    
    return True

def handle_receive_order(update, context, order_id):
    update_order_status(
        update, context, order_id,
        'received',
        "✅ Received ✅",
        "Your order has been received by admin. We will process it shortly."
    )

def handle_cancel_order(update, context, order_id):
    update_order_status(
        update, context, order_id,
        'cancelled',
        "❌ Cancelled ❌",
        "Your order has been cancelled by admin. Please contact support if you have any questions."
    )

def handle_complete_order(update, context, order_id):
    update_order_status(
        update, context, order_id,
        'completed',
        "✅ Payment Confirmed ✅",
        "Payment confirmed for your order! Thank you for choosing our service."
    )

# ================ MESSAGE HANDLER ================

def message_handler(update, context):
    user_id = update.effective_user.id
    text = update.message.text.strip() if update.message.text else None
    
    if update.message.document:
        if context.user_data.get('awaiting_file') or context.user_data.get('awaiting_report'):
            handle_file_upload(update, context)
        else:
            if user_id in ADMIN_IDS:
                update.message.reply_text("❌ Please select an option from Admin Panel first.")
            else:
                update.message.reply_text("❌ Please select a category first from the menu.")
        return
    
    if not text:
        return
    
    if context.user_data.get('waiting_for_payment_number'):
        save_payment_number(update, context)
        return
    
    if context.user_data.get('change_payment_method'):
        save_new_payment(update, context)
        return
    
    if context.user_data.get('editing_note_order') and user_id in ADMIN_IDS:
        order_id = context.user_data['editing_note_order']
        save_admin_note(update, context, order_id, text)
        return
    
    if context.user_data.get('search_user') and user_id in ADMIN_IDS:
        search_user_orders(update, context, text)
        return
    
    if context.user_data.get('add_vip') and user_id in ADMIN_IDS:
        process_add_vip(update, context, text)
        context.user_data['add_vip'] = False
        return
    
    if context.user_data.get('vip_rate_cat') and user_id in ADMIN_IDS:
        save_vip_rate(update, context, text)
        return
    
    if context.user_data.get('add_cat') == 'name' and user_id in ADMIN_IDS:
        process_cat_name(update, context, text)
        return
    
    if context.user_data.get('add_cat') == 'price' and user_id in ADMIN_IDS:
        process_cat_price(update, context, text)
        return
    
    if context.user_data.get('ren_cat') and user_id in ADMIN_IDS:
        process_category_rename(update, context, text)
        return
    
    if context.user_data.get('prc_cat') and user_id in ADMIN_IDS:
        process_category_price(update, context, text)
        return
    
    if context.user_data.get('broadcast') and user_id in ADMIN_IDS:
        send_broadcast(update, context, text)
        return
    
    # Main menu commands
    if text == "📝 File Submit":
        show_categories_menu(update, user_id)
        return
    
    elif text == "🆘 Admin Support":
        show_support_menu(update, user_id)
        return
    
    elif text == "⚙️ Settings":
        show_setting_menu(update, user_id)
        return
    
    elif text == "👑 Admin Panel" and user_id in ADMIN_IDS:
        show_admin_panel(update, user_id)
        return
    
    elif text == "🔙 Back to Main" or text == "🔙 Back":
        show_main_menu(update, user_id)
        return
    
    elif text == "🔙 Back to Admin Panel" and user_id in ADMIN_IDS:
        show_admin_panel(update, user_id)
        return
    
    elif text == "🔙 Back to Settings":
        show_setting_menu(update, user_id)
        return
    
    elif text == "🔙 Back to Categories" and user_id in ADMIN_IDS:
        show_admin_categories(update, user_id)
        return
    
    elif text == "🔙 Back to VIP Management" and user_id in ADMIN_IDS:
        show_admin_vip(update, user_id)
        return
    
    # Admin panel options
    if user_id in ADMIN_IDS:
        if text == "📁 Manage Categories":
            show_admin_categories(update, user_id)
            return
        
        elif text == "👑 VIP Users":
            show_admin_vip(update, user_id)
            return
        
        elif text == "🔍 Search User":
            search_user_start(update, context)
            return
        
        elif text == "📢 Broadcast":
            broadcast_start(update, context)
            return
        
        elif text == "📊 Report System":
            show_report_system(update, user_id)
            return
        
        elif text == "📤 Upload Report":
            start_report_upload(update, context)
            return
        
        elif text == "📋 View Reports":
            view_reports(update, user_id)
            return
        
        elif text == "➕ Add VIP":
            add_vip_start(update, context)
            return
        
        elif text == "📋 VIP List":
            show_vip_list(update, user_id)
            return
        
        elif text == "✏️ Edit Rates":
            edit_vip_rates_list(update, context)
            return
        
        elif text == "🗑 Remove VIP":
            remove_vip_start(update, context)
            return
        
        elif text == "➕ Add New Category":
            add_category_start(update, context)
            return
        
        elif text.startswith("✏️") and "📁" in text:
            parts = text.split("📁")
            if len(parts) > 1:
                cat_name = parts[1].split(" -")[0].strip()
                show_edit_category(update, context, cat_name)
                return
        
        elif text == "📝 Change Name" and context.user_data.get('editing_category'):
            cat_id = context.user_data['editing_category']
            change_category_name(update, context, cat_id)
            return
        
        elif text == "💰 Change Price" and context.user_data.get('editing_category'):
            cat_id = context.user_data['editing_category']
            change_category_price(update, context, cat_id)
            return
        
        elif text == "🗑 Delete Category" and context.user_data.get('editing_category'):
            cat_id = context.user_data['editing_category']
            delete_category(update, context, cat_id)
            context.user_data['editing_category'] = None
            return
        
        elif context.user_data.get('vip_rate_user'):
            categories = load_categories()
            for cat_id, info in categories.items():
                if info['name'] in text:
                    ask_vip_rate(update, context, info['name'])
                    return
        
        elif text.startswith("/remove_vip_") and user_id in ADMIN_IDS:
            uid = text.replace("/remove_vip_", "")
            process_remove_vip(update, context, uid)
            return
        
        elif text.startswith("/edit_vip_") and user_id in ADMIN_IDS:
            uid = text.replace("/edit_vip_", "")
            show_vip_rates_menu(update, context, uid)
            return
    
    if text == "💳 Change Payment":
        change_payment_start(update, context)
        return
    
    elif text in ["bKash", "Nagad", "Rocket", "Binance"] and context.user_data.get('change_payment_method') is None:
        change_payment_number(update, context, text)
        return
    
    # Handle category selection
    categories = load_categories()
    for cat_id, cat_info in categories.items():
        price = get_user_price(cat_id, user_id) if is_vip_user(user_id) else cat_info['price']
        price_str = f"{price:.2f}".rstrip('0').rstrip('.') if isinstance(price, float) else str(price)
        button_text = f"{cat_info['emoji']} {cat_info['name']} - {price_str} TK"
        if is_vip_user(user_id):
            button_text += " ⭐"
        
        if text == button_text:
            context.user_data['selected_category'] = cat_id
            user_type = "VIP ⭐" if is_vip_user(user_id) else "Normal"
            
            keyboard = get_main_menu_keyboard(user_id)
            update.message.reply_text(
                f"📁 Selected: {cat_info['emoji']} {cat_info['name']}\n"
                f"👤 Type: {user_type}\n"
                f"💰 Price: {price_str} TK per piece\n\n"
                f"📎 Please upload your Excel file (.xlsx)\n\n"
                f"⚠️ Only .xlsx files are accepted.\n"
                f"🔄 Duplicates in Column A will be automatically removed.\n"
                f"🔍 Existing data will be checked and removed.",
                reply_markup=keyboard
            )
            context.user_data['awaiting_file'] = True
            return
    
    # Default
    show_main_menu(update, user_id)

# ================ MAIN ================

def main():
    # Start HTTP server in background thread
    http_thread = threading.Thread(target=run_http_server, daemon=True)
    http_thread.start()
    
    updater = Updater(BOT_TOKEN, use_context=True)
    dp = updater.dispatcher
    
    dp.add_handler(CommandHandler("start", start))
    dp.add_handler(CallbackQueryHandler(button_callback))
    dp.add_handler(MessageHandler(Filters.text & ~Filters.command, message_handler))
    dp.add_handler(MessageHandler(Filters.document, message_handler))
    
    updater.start_polling()
    print("🤖 Bot is running!")
    print(f"📁 Upload directory: {UPLOAD_DIR}")
    print(f"📁 Processed directory: {PROCESSED_DIR}")
    print(f"📁 Text directory: {TEXT_DIR}")
    print(f"📁 Report directory: {REPORT_DIR}")
    print(f"📁 Report Results directory: {REPORT_RESULTS_DIR}")
    print(f"👑 Admins: {ADMIN_IDS}")
    print(f"📦 GitHub Backup: Enabled")
    
    updater.idle()

if __name__ == '__main__':
    main()
