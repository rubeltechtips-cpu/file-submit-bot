import logging
import os
import io
import time
import re
import json
from datetime import datetime
from telegram import (
    Update, ReplyKeyboardMarkup, KeyboardButton, InputFile,
    ReplyKeyboardRemove, InlineKeyboardMarkup, InlineKeyboardButton,
    ChatMember
)
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    ContextTypes, ConversationHandler, CallbackQueryHandler, filters
)
from openpyxl import Workbook, load_workbook

# ================ CONFIG ================
BOT_TOKEN = "8349208659:AAEyJikjx1tUri_PztFGRca_lPT0WilJ0N0"
ADMIN_ID = 8061006207
ADMIN_USERNAME = "Rubel_QSB"
CHANNEL_USERNAME = "quick_sell_bd"
DATA_DIR = "categories"
TXT_DIR = "txt_files"
EXCEL_DIR = "excel_files"
LOG_DIR = "logs"

os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(TXT_DIR, exist_ok=True)
os.makedirs(EXCEL_DIR, exist_ok=True)
os.makedirs(LOG_DIR, exist_ok=True)

# ================ STATES ================
(
    MAIN_MENU,
    BUY_MENU,
    BUY_SUB_MENU,
    ADMIN_PANEL,
    ADD_MAIN_CAT,
    REMOVE_MAIN_CAT,
    MANAGE_CATEGORY,
    MANAGE_SUB_CATEGORY,
    ADD_SUB_CAT,
    REMOVE_SUB_CAT,
    ADD_ITEMS_TXT,
    EDIT_PAYMENT,
    EDIT_PRICE_MAIN,
    EDIT_PRICE_SUB,
    RECEIVE_NEW_PRICE,
    GET_QUANTITY,
    WAIT_SCREENSHOT,
    DEPOSIT,
    GET_DEPOSIT_AMOUNT,
    DASHBOARD,
    SEND_NOTICE,
    VIEW_USER_PROFILE,
    SEARCH_USER_PROFILE,
    MANAGE_PAYMENT_CATEGORIES,
    SEARCH_USER_FOR_BALANCE,
    BALANCE_EDIT_ACTION,
    RECEIVE_BALANCE_EDIT_AMOUNT,
    CONFIRM_ORDER,
    DEPOSIT_SELECT_METHOD,
    DEPOSIT_ENTER_AMOUNT,
    DEPOSIT_ENTER_TRXID,
    ADMIN_ADD_PAYMENT_METHOD,
    ADMIN_VIEW_PAYMENT_METHODS,
    ADMIN_DELETE_PAYMENT_METHOD,
    ADMIN_RECEIVE_PAYMENT_DETAILS
) = range(35)

# ================ LOGGER ================
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# SMS লগ ফাইল
SMS_LOG_FILE = os.path.join(LOG_DIR, "sms_log.txt")
DEPOSIT_LOG_FILE = os.path.join(LOG_DIR, "deposit_log.txt")

# ================ DATA ================
categories = {}
prices = {}
payment_methods = {}
balances = {}
user_sales = {}
user_deposits = {}
user_info = {}
total_deposits = 0
total_sales = 0
sales_count_per_category = {}
transaction_log = []
dashboard_message = "স্বাগতম! এটি আপনার বটের ড্যাশবোর্ড।"
MANUAL_DELIVERY_CATEGORIES = []

# অটো ডিপোজিট সিস্টেমের জন্য ডেটা
pending_deposits = {}  # {trxid: {"user_id": user_id, "amount": amount, "timestamp": time, "method": method, "username": username}}
processed_trxids = set()  # ইতিমধ্যে প্রসেস করা TRXID গুলো
sms_log = []  # সমস্ত SMS লগ

# ================ DATA PERSISTENCE ================
def load_user_data():
    global balances, user_sales, user_deposits, user_info, total_deposits, total_sales
    global transaction_log, categories, prices, sales_count_per_category
    global MANUAL_DELIVERY_CATEGORIES, payment_methods, pending_deposits, processed_trxids
    try:
        with open("user_data.json", "r", encoding='utf-8') as f:
            data = json.load(f)
            balances.update({int(k): v for k, v in data.get("balances", {}).items()})
            user_sales.update({int(k): v for k, v in data.get("user_sales", {}).items()})
            user_deposits.update({int(k): v for k, v in data.get("user_deposits", {}).items()})
            user_info.update({int(k): v for k, v in data.get("user_info", {}).items()})
            total_deposits = data.get("total_deposits", 0)
            total_sales = data.get("total_sales", 0)
            transaction_log = data.get("transaction_log", [])
            categories.update(data.get("categories", {}))
            prices.update(data.get("prices", {}))
            sales_count_per_category.update(data.get("sales_count_per_category", {}))
            MANUAL_DELIVERY_CATEGORIES = data.get("manual_delivery_categories", [])
            payment_methods.update(data.get("payment_methods", {}))
            pending_deposits.update(data.get("pending_deposits", {}))
            processed_trxids = set(data.get("processed_trxids", []))
    except FileNotFoundError:
        pass

def save_user_data():
    with open("user_data.json", "w", encoding='utf-8') as f:
        data = {
            "balances": balances,
            "user_sales": user_sales,
            "user_deposits": user_deposits,
            "user_info": user_info,
            "total_deposits": total_deposits,
            "total_sales": total_sales,
            "transaction_log": transaction_log,
            "categories": categories,
            "prices": prices,
            "sales_count_per_category": sales_count_per_category,
            "manual_delivery_categories": MANUAL_DELIVERY_CATEGORIES,
            "payment_methods": payment_methods,
            "pending_deposits": pending_deposits,
            "processed_trxids": list(processed_trxids)
        }
        json.dump(data, f, ensure_ascii=False, indent=4)

def log_sms(sms_text, trxid=None, amount=None, status="RECEIVED"):
    """SMS লগ সংরক্ষণ"""
    log_entry = {
        "timestamp": datetime.now().isoformat(),
        "sms_text": sms_text[:500],
        "trxid": trxid,
        "amount": amount,
        "status": status
    }
    sms_log.append(log_entry)
    
    # ফাইলে সংরক্ষণ
    with open(SMS_LOG_FILE, "a", encoding='utf-8') as f:
        f.write(f"[{log_entry['timestamp']}] STATUS: {status} | TRXID: {trxid} | Amount: {amount}\n")
        f.write(f"SMS: {sms_text[:300]}\n")
        f.write("-" * 50 + "\n")
    
    return log_entry

def log_deposit(user_id, amount, trxid, status="AUTO", method=None):
    """ডিপোজিট লগ সংরক্ষণ"""
    log_entry = {
        "timestamp": datetime.now().isoformat(),
        "user_id": user_id,
        "amount": amount,
        "trxid": trxid,
        "status": status,
        "method": method
    }
    
    with open(DEPOSIT_LOG_FILE, "a", encoding='utf-8') as f:
        f.write(f"[{log_entry['timestamp']}] {status} | User: {user_id} | Amount: {amount} | TRXID: {trxid}\n")
        f.write("-" * 50 + "\n")
    
    return log_entry

load_user_data()

# ================ HELPERS ================
def get_txt_path(main_cat: str, sub_cat: str) -> str:
    file_name = f"{main_cat}_{sub_cat}.txt".replace(" ", "_").replace("-", "_")
    return os.path.join(TXT_DIR, file_name)

def get_excel_path(main_cat: str, sub_cat: str) -> str:
    file_name = f"{main_cat}_{sub_cat}.xlsx".replace(" ", "_").replace("-", "_")
    return os.path.join(EXCEL_DIR, file_name)

def ensure_txt_file(main_cat: str, sub_cat: str):
    path = get_txt_path(main_cat, sub_cat)
    if not os.path.exists(path):
        with open(path, 'w', encoding='utf-8') as f:
            f.write("")

def add_items_from_txt(main_cat: str, sub_cat: str, txt_content: str):
    ensure_txt_file(main_cat, sub_cat)
    path = get_txt_path(main_cat, sub_cat)
    with open(path, 'a', encoding='utf-8') as f:
        f.write(txt_content + '\n')

def pop_items_from_txt(main_cat: str, sub_cat: str, qty: int) -> list:
    path = get_txt_path(main_cat, sub_cat)
    if not os.path.exists(path):
        return []
    
    with open(path, 'r', encoding='utf-8') as f:
        items = [line.strip() for line in f.readlines() if line.strip()]
    
    if len(items) < qty:
        return []
    
    result = items[:qty]
    remaining_items = items[qty:]
    
    with open(path, 'w', encoding='utf-8') as f:
        for item in remaining_items:
            f.write(item + '\n')
    
    return result

def count_items(main_cat: str, sub_cat: str) -> int:
    path = get_txt_path(main_cat, sub_cat)
    if not os.path.exists(path):
        return 0
    with open(path, 'r', encoding='utf-8') as f:
        items = [line.strip() for line in f.readlines() if line.strip()]
    return len(items)

def create_xlsx_file(items: list, file_name: str) -> io.BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "Items"
    for item in items:
        ws.append([item])
    
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def get_total_stock(main_cat: str) -> int:
    total = 0
    if main_cat in categories:
        for sub_cat in categories[main_cat]:
            total += count_items(main_cat, sub_cat)
    return total

def get_report_summary(transactions, days):
    end_timestamp = time.time()
    start_timestamp = end_timestamp - (days * 24 * 60 * 60)
    
    daily_deposits = 0
    daily_sales = 0
    
    for trans_type, _, amount, timestamp in transactions:
        if start_timestamp <= timestamp <= end_timestamp:
            if trans_type == 'deposit':
                daily_deposits += amount
            elif trans_type == 'sale':
                daily_sales += amount
    return daily_deposits, daily_sales
    
def get_user_transactions(user_id, transactions):
    return [t for t in transactions if t[1] == user_id]

# ================ SMS AUTO DEPOSIT PROCESSOR ================
def extract_trxid_from_bkash_sms(sms_text: str):
    """
    bKash SMS থেকে TRXID বের করে
    বিশেষ করে আপনার দেখানো ফরম্যাটের জন্য:
    "TrxID DFQ9POCKRB at 26/06/2026"
    """
    trxid = None
    
    print(f"\n🔍 SMS থেকে TRXID বের করা হচ্ছে:")
    print(f"SMS: {sms_text}")
    
    # প্রথমে TrxID খুঁজুন - আপনার দেখানো ফরম্যাটের জন্য
    # "TrxID DFQ9POCKRB" - এখানে space আছে
    match = re.search(r'TrxID\s+([A-Z0-9]{6,})', sms_text, re.IGNORECASE)
    if match:
        trxid = match.group(1).strip().upper()
        print(f"✅ TRXID পাওয়া গেছে (TrxID): {trxid}")
        return trxid
    
    # যদি না পায়, তাহলে TRXID খুঁজুন
    match = re.search(r'TRXID\s+([A-Z0-9]{6,})', sms_text, re.IGNORECASE)
    if match:
        trxid = match.group(1).strip().upper()
        print(f"✅ TRXID পাওয়া গেছে (TRXID): {trxid}")
        return trxid
    
    # অন্য প্যাটার্ন
    patterns = [
        r'TrxID[:\s]*([A-Z0-9]{6,})',
        r'TRXID[:\s]*([A-Z0-9]{6,})',
        r'ট্রানজেকশন আইডি[:\s]*([A-Z0-9]{6,})',
        r'আইডি[:\s]*([A-Z0-9]{6,})',
        r'ID[:\s]*([A-Z0-9]{6,})',
        r'([A-Z0-9]{8,12})(?:\s+at\s+|\s*$)',
    ]
    
    for pattern in patterns:
        match = re.search(pattern, sms_text, re.IGNORECASE)
        if match:
            potential_trxid = match.group(1).strip().upper()
            if len(potential_trxid) >= 6 and re.match(r'^[A-Z0-9]+$', potential_trxid):
                trxid = potential_trxid
                print(f"✅ TRXID পাওয়া গেছে: {trxid}")
                return trxid
    
    print(f"⚠️ TRXID বের করতে ব্যর্থ")
    return None

def extract_amount_from_bkash_sms(sms_text: str):
    """
    bKash SMS থেকে পরিমাণ বের করে
    বিশেষ করে আপনার দেখানো ফরম্যাটের জন্য:
    "You have received Tk 50.00 from 01920525242"
    """
    amount = None
    
    print(f"\n🔍 SMS থেকে পরিমাণ বের করা হচ্ছে:")
    print(f"SMS: {sms_text}")
    
    # প্রথমে "received Tk" প্যাটার্ন চেক করুন - আপনার দেখানো ফরম্যাট
    match = re.search(r'received\s+Tk\s*([\d,]+\.?\d*)', sms_text, re.IGNORECASE)
    if match:
        amount_str = match.group(1).replace(',', '').strip()
        try:
            amount = int(float(amount_str))
            print(f"✅ পরিমাণ পাওয়া গেছে (received Tk): {amount} টাকা")
            return amount
        except:
            pass
    
    # "Tk X.XX from" প্যাটার্ন
    match = re.search(r'Tk\s*([\d,]+\.?\d*)\s+from', sms_text, re.IGNORECASE)
    if match:
        amount_str = match.group(1).replace(',', '').strip()
        try:
            amount = int(float(amount_str))
            print(f"✅ পরিমাণ পাওয়া গেছে (Tk X from): {amount} টাকা")
            return amount
        except:
            pass
    
    # অন্যান্য প্যাটার্ন
    patterns = [
        r'Tk\s*([\d,]+\.?\d*)',
        r'([\d,]+)\s*টাকা',
        r'([\d,]+)\s*TK',
        r'([\d,]+)\s*BDT',
        r'অ্যাকাউন্টে\s*([\d,]+)',
        r'([\d,]+)\s*টাকা',
        r'আমান\s*([\d,]+)',
        r'([\d,]+)\s*/\s*-\s*',
        r'([\d,]+)\.\d{2}\s*Tk',
    ]
    
    for pattern in patterns:
        match = re.search(pattern, sms_text, re.IGNORECASE)
        if match:
            amount_str = match.group(1).replace(',', '').strip()
            try:
                if '.' in amount_str:
                    amt = float(amount_str)
                else:
                    amt = float(amount_str)
                
                if 1 <= amt <= 100000:
                    amount = int(round(amt))
                    print(f"✅ পরিমাণ পাওয়া গেছে: {amount} টাকা")
                    return amount
            except:
                continue
    
    print(f"⚠️ পরিমাণ বের করতে ব্যর্থ")
    return None

def process_bkash_sms(sms_text: str):
    """
    bKash SMS সম্পূর্ণ প্রসেস করে
    """
    print(f"\n{'='*50}")
    print(f"📱 নতুন SMS প্রসেস করা হচ্ছে:")
    print(f"{'='*50}")
    print(f"SMS: {sms_text}")
    print(f"{'='*50}")
    
    trxid = extract_trxid_from_bkash_sms(sms_text)
    amount = extract_amount_from_bkash_sms(sms_text)
    
    print(f"{'='*50}")
    print(f"📊 ফলাফল:")
    print(f"  TRXID: {trxid}")
    print(f"  পরিমাণ: {amount} টাকা")
    print(f"{'='*50}\n")
    
    return trxid, amount

async def auto_deposit_from_sms(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    SMS থেকে অটো ডিপোজিট প্রসেস করে
    """
    global total_deposits, transaction_log, user_deposits, balances, pending_deposits, processed_trxids
    
    # শুধুমাত্র অ্যাডমিনের চ্যাট থেকে আসা SMS প্রসেস করুন
    if update.effective_user.id != ADMIN_ID:
        print(f"⚠️ অ্যাডমিন নয়: {update.effective_user.id}")
        return
    
    message_text = update.message.text
    print(f"\n📱 নতুন SMS প্রাপ্ত:")
    print(f"SMS: {message_text}")
    
    # SMS লগ করুন
    log_sms(message_text)
    
    # প্রথমে রেসপন্স দিন যে SMS পাওয়া গেছে
    await update.message.reply_text(
        f"📱 SMS প্রাপ্ত হয়েছে। প্রসেস করা হচ্ছে...\n"
        f"🔍 TRXID এবং পরিমাণ বের করা হচ্ছে..."
    )
    
    # bKash SMS প্রসেস করুন
    trxid, amount = process_bkash_sms(message_text)
    
    if not trxid:
        await update.message.reply_text(
            f"⚠️ SMS থেকে TRXID বের করতে ব্যর্থ হয়েছে।\n"
            f"📱 SMS: {message_text[:100]}...\n\n"
            f"💡 TRXID ফরম্যাট: TrxID XXXXXXXX"
        )
        return
    
    if not amount:
        await update.message.reply_text(
            f"⚠️ SMS থেকে পরিমাণ বের করতে ব্যর্থ হয়েছে।\n"
            f"🔑 TRXID: {trxid}\n"
            f"📱 SMS: {message_text[:100]}..."
        )
        return
    
    # ইতিমধ্যে প্রসেস করা TRXID চেক করুন
    if trxid in processed_trxids:
        await update.message.reply_text(
            f"⏭️ এই TRXID ইতিমধ্যে প্রসেস করা হয়েছে:\n"
            f"🔑 {trxid}\n"
            f"💰 {amount} টাকা"
        )
        return
    
    # পেন্ডিং ডিপোজিটের সাথে মিলিয়ে দেখুন
    if trxid in pending_deposits:
        deposit_data = pending_deposits[trxid]
        user_id = deposit_data['user_id']
        expected_amount = deposit_data['amount']
        username = deposit_data.get('username', 'N/A')
        method = deposit_data.get('method', 'Unknown')
        
        print(f"\n🔍 পেন্ডিং ডিপোজিট পাওয়া গেছে:")
        print(f"  ইউজার আইডি: {user_id}")
        print(f"  ইউজারনেম: {username}")
        print(f"  প্রত্যাশিত পরিমাণ: {expected_amount}")
        print(f"  SMS পরিমাণ: {amount}")
        
        # পরিমাণ মিলিয়ে দেখুন
        if amount == expected_amount:
            # ব্যালেন্স অ্যাড করুন
            balances[user_id] = balances.get(user_id, 0) + amount
            total_deposits += amount
            transaction_log.append(('deposit', user_id, amount, time.time()))
            user_deposits[user_id] = user_deposits.get(user_id, 0) + amount
            
            # পেন্ডিং লিস্ট থেকে রিমুভ করুন
            del pending_deposits[trxid]
            processed_trxids.add(trxid)
            save_user_data()
            
            # ডিপোজিট লগ
            log_deposit(user_id, amount, trxid, "AUTO_SUCCESS", method)
            
            # ইউজারকে নোটিফিকেশন পাঠান
            try:
                await context.bot.send_message(
                    chat_id=user_id,
                    text=f"✅ আপনার {amount} টাকা ডিপোজিট স্বয়ংক্রিয়ভাবে সফল হয়েছে!\n"
                         f"🔑 TRXID: {trxid}\n"
                         f"💰 নতুন ব্যালেন্স: {balances[user_id]} টাকা।"
                )
                print(f"✅ ইউজার {user_id} কে নোটিফিকেশন পাঠানো হয়েছে")
            except Exception as e:
                logger.error(f"Failed to notify user {user_id}: {e}")
            
            # অ্যাডমিনকে জানান
            await update.message.reply_text(
                f"✅ অটো ডিপোজিট সফল!\n\n"
                f"👤 ইউজার: @{username} (ID: {user_id})\n"
                f"💰 পরিমাণ: {amount} টাকা\n"
                f"🔑 TRXID: {trxid}\n"
                f"💳 মেথড: {method}\n"
                f"📱 SMS: {message_text[:150]}..."
            )
            
        else:
            # পরিমাণ মিলছে না
            await update.message.reply_text(
                f"⚠️ TRXID মিলেছে কিন্তু পরিমাণ মিলছে না!\n\n"
                f"🔑 TRXID: {trxid}\n"
                f"💰 ইউজার দিয়েছে: {expected_amount} টাকা\n"
                f"💰 SMS এ আছে: {amount} টাকা\n"
                f"👤 ইউজার: @{username} (ID: {user_id})\n"
                f"💳 মেথড: {method}\n\n"
                f"⚠️ ম্যানুয়ালি চেক করুন।"
            )
            
            # ডিপোজিট লগ
            log_deposit(user_id, amount, trxid, "AUTO_FAIL_AMOUNT_MISMATCH", method)
    else:
        # TRXID পেন্ডিং লিস্টে নেই
        await update.message.reply_text(
            f"ℹ️ এই TRXID এর জন্য কোনো পেন্ডিং ডিপোজিট নেই:\n\n"
            f"🔑 TRXID: {trxid}\n"
            f"💰 পরিমাণ: {amount} টাকা\n"
            f"📱 SMS: {message_text[:150]}...\n\n"
            f"কোনো ইউজার এই TRXID ব্যবহার করে ডিপোজিট করেনি অথবা TRXID ভুল।"
        )
        
        # এই TRXID প্রসেসড হিসেবে চিহ্নিত করুন (পুনরায় প্রসেস না করার জন্য)
        processed_trxids.add(trxid)
        save_user_data()
        
        # ডিপোজিট লগ
        log_deposit(0, amount, trxid, "AUTO_NO_PENDING")

# ================ BACK TO MAIN MENU ================
async def back_to_main_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    return await start(update, context)

# ================ CHECK SUBSCRIPTION ================
async def check_subscription(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    try:
        chat_member = await context.bot.get_chat_member(chat_id=f"@{CHANNEL_USERNAME}", user_id=user_id)
        if chat_member.status in [ChatMember.MEMBER, ChatMember.ADMINISTRATOR, ChatMember.OWNER]:
            return True
        else:
            keyboard = InlineKeyboardMarkup([
                [InlineKeyboardButton("Join Channel", url=f"https://t.me/{CHANNEL_USERNAME}")]
            ])
            await update.message.reply_text(
                "❌ আপনি এখনো আমাদের চ্যানেলে জয়েন করেননি।\n"
                "বট ব্যবহার করার জন্য অনুগ্রহ করে নিচের বাটনে ক্লিক করে চ্যানেলে জয়েন করুন।",
                reply_markup=keyboard
            )
            return False
    except Exception as e:
        logger.error(f"Error checking subscription: {e}")
        await update.message.reply_text("চ্যানেল সদস্যতা পরীক্ষা করতে সমস্যা হচ্ছে। অনুগ্রহ করে পরে আবার চেষ্টা করুন।")
        return False

# ================ START ================
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await check_subscription(update, context):
        return ConversationHandler.END
    
    user_id = update.effective_user.id
    username = update.effective_user.username
    first_name = update.effective_user.first_name
    
    if user_id not in user_info:
        user_info[user_id] = {
            "username": username,
            "first_name": first_name,
            "last_name": update.effective_user.last_name,
            "id": user_id
        }
        save_user_data()

    current_balance = balances.get(user_id, 0)

    keyboard = [
        [KeyboardButton("🛒 Buy"), KeyboardButton("💰 Balance")],
        [KeyboardButton("💸 Deposit"), KeyboardButton("📞 Help")],
    ]
    if update.effective_user.id == ADMIN_ID:
        keyboard.append([KeyboardButton("⚙️ Admin Panel")])
    
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=False)
    await update.message.reply_text(f"👋 স্বাগতম! আপনার বর্তমান ব্যালেন্স: {current_balance} টাকা।", reply_markup=reply_markup)
    return MAIN_MENU

async def menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await check_subscription(update, context):
        return ConversationHandler.END

    text = update.message.text
    
    if text == "🛒 Buy":
        if not categories:
            await update.message.reply_text("⚠️ এখন কোনো ক্যাটাগরি নেই।")
            return MAIN_MENU
        
        keyboard = []
        for cat in categories.keys():
            keyboard.append([KeyboardButton(cat)])
        
        keyboard.append([KeyboardButton("🔙 Back to Main Menu")])
        await update.message.reply_text("🛒 ক্যাটাগরি বেছে নিন:", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))
        return BUY_MENU

    if text == "💰 Balance":
        user_id = update.effective_user.id
        current_balance = balances.get(user_id, 0)
        await update.message.reply_text(f"আপনার বর্তমান ব্যালেন্স: {current_balance} টাকা।")
        return MAIN_MENU

    if text == "💸 Deposit":
        return await deposit_select_method(update, context)

    if text == "📞 Help":
        keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("📞 Contact Admin", url=f"tg://user?id={ADMIN_ID}")]
        ])
        await update.message.reply_text(
            "📞 অ্যাডমিনের সাথে যোগাযোগ করতে নিচের বাটনে ক্লিক করুন।",
            reply_markup=keyboard
        )
        return MAIN_MENU

    if text == "⚙️ Admin Panel":
        if update.effective_user.id == ADMIN_ID:
            return await show_dashboard(update, context)
        else:
            await update.message.reply_text("❌ অননুমোদিত।")
            return MAIN_MENU
            
    if text == "🔙 Back to Main Menu":
        return await start(update, context)
    
    return MAIN_MENU

# ================ DEPOSIT SYSTEM ================
async def deposit_select_method(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not payment_methods:
        await update.message.reply_text(
            "❌ এখন কোনো পেমেন্ট মেথড নেই। অ্যাডমিনের সাথে যোগাযোগ করুন।",
            reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Back to Main Menu")]], resize_keyboard=True)
        )
        return MAIN_MENU
    
    keyboard = []
    for method_name, method_data in payment_methods.items():
        if method_data.get("active", True):
            keyboard.append([KeyboardButton(f"💳 {method_name}")])
    
    keyboard.append([KeyboardButton("🔙 Back to Main Menu")])
    
    await update.message.reply_text(
        "💸 অনুগ্রহ করে আপনার ডিপোজিট মেথড নির্বাচন করুন:",
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )
    return DEPOSIT_SELECT_METHOD

async def deposit_enter_amount(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    
    if text == "🔙 Back to Main Menu":
        return await start(update, context)
    
    # Extract method name from button text
    method_name = text.replace("💳 ", "").strip()
    
    if method_name not in payment_methods:
        await update.message.reply_text(
            "❌ অবৈধ পেমেন্ট মেথড। অনুগ্রহ করে আবার নির্বাচন করুন।",
            reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Back to Main Menu")]], resize_keyboard=True)
        )
        return DEPOSIT_SELECT_METHOD
    
    context.user_data['deposit_method'] = method_name
    
    # Show payment details
    method_details = payment_methods[method_name].get("details", "কোনো বিস্তারিত তথ্য নেই।")
    
    await update.message.reply_text(
        f"💳 পেমেন্ট মেথড: {method_name}\n"
        f"📋 বিস্তারিত:\n{method_details}\n\n"
        f"✍️ আপনি কত টাকা ডিপোজিট করতে চান তা লিখুন:",
        reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Back to Main Menu")]], resize_keyboard=True)
    )
    return DEPOSIT_ENTER_AMOUNT

async def deposit_receive_amount(update: Update, context: ContextTypes.DEFAULT_TYPE):
    amount_str = update.message.text.strip()
    
    if amount_str == "🔙 Back to Main Menu":
        return await start(update, context)
    
    if not amount_str.isdigit() or int(amount_str) <= 0:
        await update.message.reply_text(
            "❌ অনুগ্রহ করে একটি বৈধ সংখ্যা লিখুন।",
            reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Back to Main Menu")]], resize_keyboard=True)
        )
        return DEPOSIT_ENTER_AMOUNT
    
    amount = int(amount_str)
    context.user_data['deposit_amount'] = amount
    
    await update.message.reply_text(
        f"✅ আপনি {amount} টাকা ডিপোজিট করতে চান।\n\n"
        f"✍️ এখন আপনার ট্রানজেকশন আইডি (TRXID) লিখুন:\n"
        f"(যে আইডি পেমেন্টের সময় পেয়েছেন)",
        reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Back to Main Menu")]], resize_keyboard=True)
    )
    return DEPOSIT_ENTER_TRXID

async def deposit_receive_trxid(update: Update, context: ContextTypes.DEFAULT_TYPE):
    trxid = update.message.text.strip().upper()
    
    if trxid == "🔙 Back to Main Menu":
        return await start(update, context)
    
    if not trxid or len(trxid) < 6:
        await update.message.reply_text(
            "❌ অনুগ্রহ করে একটি বৈধ ট্রানজেকশন আইডি লিখুন (অন্তত ৬ অক্ষর)।",
            reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Back to Main Menu")]], resize_keyboard=True)
        )
        return DEPOSIT_ENTER_TRXID
    
    method_name = context.user_data.get('deposit_method', 'Unknown')
    amount = context.user_data.get('deposit_amount', 0)
    user = update.effective_user
    username = user.username if user.username else 'N/A'
    
    # TRXID টি পেন্ডিং লিস্টে সেভ করুন (অটো ডিপোজিটের জন্য)
    global pending_deposits
    pending_deposits[trxid] = {
        "user_id": user.id,
        "amount": amount,
        "timestamp": time.time(),
        "method": method_name,
        "username": username
    }
    save_user_data()
    
    # ডিপোজিট লগ
    log_deposit(user.id, amount, trxid, "PENDING", method_name)
    
    # অ্যাডমিনকে নোটিফিকেশন পাঠান (ব্যাকআপের জন্য)
    caption = (
        f"🔔 নতুন ডিপোজিট রিকোয়েস্ট! 🔔\n"
        f"👤 ব্যবহারকারী: @{username}\n"
        f"🆔 ইউজার আইডি: {user.id}\n"
        f"💳 মেথড: {method_name}\n"
        f"💰 পরিমাণ: {amount} টাকা\n"
        f"🔑 ট্রানজেকশন আইডি: {trxid}\n"
        f"⏰ সময়: {time.strftime('%Y-%m-%d %H:%M:%S')}\n\n"
        f"📱 এই TRXID অটো ভেরিফাই হবে যখন অ্যাডমিনের ফোনে SMS আসবে!\n"
        f"⚠️ যদি ৫ মিনিটের মধ্যে অটো না হয়, ম্যানুয়ালি কনফর্ম করুন।"
    )
    
    keyboard = InlineKeyboardMarkup([
        [
            InlineKeyboardButton("✅ Confirm (Manual)", callback_data=f"deposit_confirm:{user.id}:{amount}"),
            InlineKeyboardButton("❌ Cancel", callback_data=f"deposit_cancel:{user.id}")
        ]
    ])
    
    await context.bot.send_message(
        chat_id=ADMIN_ID,
        text=caption,
        reply_markup=keyboard
    )
    
    await update.message.reply_text(
        f"✅ আপনার ডিপোজিট রিকোয়েস্ট পাঠানো হয়েছে।\n\n"
        f"📱 যখন অ্যাডমিনের ফোনে আপনার ট্রানজেকশনের SMS আসবে, তখন স্বয়ংক্রিয়ভাবে আপনার ব্যালেন্স অ্যাড হবে!\n"
        f"⏳ সাধারণত ১-২ মিনিটের মধ্যে হয়ে যায়।\n\n"
        f"⚠️ যদি ৫ মিনিটের মধ্যে অটো অ্যাড না হয়, অ্যাডমিন ম্যানুয়ালি কনফর্ম করবেন।",
        reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Back to Main Menu")]], resize_keyboard=True)
    )
    
    context.user_data.clear()
    return MAIN_MENU

# ================ DASHBOARD ================
async def show_dashboard(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU

    global total_deposits, total_sales, balances, sales_count_per_category, user_info, transaction_log, dashboard_message

    total_users_count = len(user_info)
    
    stock_info = ""
    for main_cat, sub_cats in categories.items():
        stock_info += f"  - {main_cat}\n" 
        for sub_cat in sub_cats:
            count = count_items(main_cat, sub_cat)
            stock_info += f"    - {sub_cat}: {count} টি আইটেম\n"

    sorted_sales = sorted(sales_count_per_category.items(), key=lambda item: item[1], reverse=True)
    top_selling_info = ""
    for sub_cat, count in sorted_sales[:10]:
        top_selling_info += f"  - {sub_cat}: {count} বিক্রয়\n"
    if len(sorted_sales) > 10:
        top_selling_info += f"  - ... এবং আরও {len(sorted_sales) - 10}টি ক্যাটাগরি"

    recent_transactions = ""
    last_5_transactions = transaction_log[-5:]
    if last_5_transactions:
        for trans in reversed(last_5_transactions):
            trans_type, user_id, amount, timestamp = trans
            date_str = time.strftime('%H:%M %b %d', time.localtime(timestamp))
            user_data = user_info.get(user_id, {})
            username = user_data.get("username", "N/A")
            if trans_type == 'deposit':
                recent_transactions += f"  - 💸 ডিপোজিট: {amount} টাকা (@{username}) {date_str}\n" 
            elif trans_type == 'sale':
                recent_transactions += f"  - 🛒 বিক্রয়: {amount} টাকা (@{username}) {date_str}\n"
    else:
        recent_transactions = "  - কোনো সাম্প্রতিক লেনদেন নেই।\n"
    
    daily_deposits, daily_sales = get_report_summary(transaction_log, 1)
    weekly_deposits, weekly_sales = get_report_summary(transaction_log, 7)
    monthly_deposits, monthly_sales = get_report_summary(transaction_log, 30)

    # পেন্ডিং ডিপোজিটের সংখ্যা
    pending_count = len(pending_deposits)
    
    # প্রসেস করা TRXID সংখ্যা
    processed_count = len(processed_trxids)

    dashboard_text = (
        f"📝 ড্যাশবোর্ড মেসেজ:\n"
        f"{dashboard_message}\n"
        "---------------------------\n"
        "📊 ড্যাশবোর্ড সামারি\n"
        f"👥 মোট ব্যবহারকারী: {total_users_count}\n"
        f"💰 মোট ব্যালেন্স: {sum(balances.values())} টাকা\n"
        f"🛒 মোট বিক্রয়: {total_sales} টাকা\n"
        f"💸 মোট ডিপোজিট: {total_deposits} টাকা\n"
        f"⏳ পেন্ডিং ডিপোজিট: {pending_count} টি\n"
        f"✅ প্রসেসড TRXID: {processed_count} টি\n"
        "---------------------------\n"
        "📈 দৈনিক/সাপ্তাহিক/মাসিক রিপোর্ট\n"
        f"গত ২৪ ঘণ্টা:\n"
        f"  - ডিপোজিট: {daily_deposits} টাকা\n"
        f"  - বিক্রয়: {daily_sales} টাকা\n"
        f"গত ৭ দিন:\n"
        f"  - ডিপোজিট: {weekly_deposits} টাকা\n"
        f"  - বিক্রয়: {weekly_sales} টাকা\n"
        f"গত ৩০ দিন:\n"
        f"  - ডিপোজিট: {monthly_deposits} টাকা\n"
        f"  - বিক্রয়: {monthly_sales} টাকা\n"
        "---------------------------\n"
        "📦 বর্তমান স্টক তথ্য:\n"
        f"{stock_info or '  - কোনো ক্যাটাগরি পাওয়া যায়নি।'}\n"
        "---------------------------\n"
        "📈 সর্বাধিক বিক্রীত ক্যাটাগরি:\n"
        f"{top_selling_info or '  - এখনো কোনো বিক্রয় হয়নি।'}\n"
        "---------------------------\n"
        "📜 সর্বশেষ লেনদেন:\n"
        f"{recent_transactions}\n"
    )
    
    if len(dashboard_text) > 4000:
        parts = [dashboard_text[i:i+4000] for i in range(0, len(dashboard_text), 4000)]
        for part in parts:
            await update.message.reply_text(part)
    else:
        await update.message.reply_text(dashboard_text)
    
    keyboard = [
        [KeyboardButton("🔄 Refresh Dashboard"), KeyboardButton("👥 User Profile")],
        [KeyboardButton("📂 Manage Categories"), KeyboardButton("💰 Edit Price")],
        [KeyboardButton("✏️ Edit Balance"), KeyboardButton("📢 Send Notice")],
        [KeyboardButton("💳 Payment Methods"), KeyboardButton("💳 Payment Categories")],
        [KeyboardButton("📱 SMS Auto Deposit"), KeyboardButton("📋 SMS Log")],
        [KeyboardButton("🔙 Back to Main Menu")]
    ]
    
    await update.message.reply_text("⚙️ অ্যাডমিন প্যানেল:", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))
    return ADMIN_PANEL

async def handle_dashboard_refresh(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU

    if update.message.text == "🔄 Refresh Dashboard":
        return await show_dashboard(update, context)
    return await back_to_admin_panel_handler(update, context)

# ================ USER PROFILE ================
async def view_user_profile(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    await update.message.reply_text("✍️ যে ব্যবহারকারীর প্রোফাইল দেখতে চান তার ইউজারনাম বা ইউজার আইডি লিখুন:", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
    return SEARCH_USER_PROFILE

async def search_and_show_user_profile(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
        
    search_term = update.message.text.strip().lstrip('@')
    
    if search_term == "🔙 Admin Panel":
        return await back_to_admin_panel_handler(update, context)

    found_user_id = None
    
    if search_term.isdigit():
        search_id = int(search_term)
        if search_id in user_info:
            found_user_id = search_id
    
    if not found_user_id:
        for user_id, info in user_info.items():
            if info.get("username") and info["username"].lower() == search_term.lower():
                found_user_id = user_id
                break

    if not found_user_id:
        await update.message.reply_text("❌ এই ইউজারনেম বা ইউজার আইডি এর কোনো ব্যবহারকারী পাওয়া যায়নি।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
        return SEARCH_USER_PROFILE

    user_transactions = get_user_transactions(found_user_id, transaction_log)
    
    balance = balances.get(found_user_id, 0)
    deposits = user_deposits.get(found_user_id, 0)
    sales = user_sales.get(found_user_id, 0)

    daily_deposits, daily_sales = get_report_summary(user_transactions, 1)
    weekly_deposits, weekly_sales = get_report_summary(user_transactions, 7)
    monthly_deposits, monthly_sales = get_report_summary(user_transactions, 30)
    yearly_deposits, yearly_sales = get_report_summary(user_transactions, 365)
    
    user_data = user_info.get(found_user_id, {})
    full_name = user_data.get("first_name", "") + (f" {user_data['last_name']}" if user_data.get("last_name") else "")
    username = user_data.get('username', 'N/A')
    
    profile_text = (
        f"👤 ব্যবহারকারী প্রোফাইল:\n"
        f"নাম: {full_name}\n"
        f"ইউজারনেম: @{username}\n"
        f"আইডি: {found_user_id}\n"
        "---------------------------\n"
        f"💰 বর্তমান ব্যালেন্স: {balance} টাকা\n"
        f"💸 মোট ডিপোজিট: {deposits} টাকা\n"
        f"🛒 মোট খরচ: {sales} টাকা\n"
        "---------------------------\n"
        "📈 লেনদেনের রিপোর্ট\n"
        f"গত ২৪ ঘণ্টা:\n"
        f"  - ডিপোজিট: {daily_deposits} টাকা\n"
        f"  - খরচ: {daily_sales} টাকা\n"
        f"গত ৭ দিন:\n"
        f"  - ডিপোজিট: {weekly_deposits} টাকা\n"
        f"  - খরচ: {weekly_sales} টাকা\n"
        f"গত ৩০ দিন:\n"
        f"  - ডিপোজিট: {monthly_deposits} টাকা\n"
        f"  - খরচ: {monthly_sales} টাকা\n"
        f"গত ১ বছর:\n"
        f"  - ডিপোজিট: {yearly_deposits} টাকা\n"
        f"  - খরচ: {yearly_sales} টাকা\n"
    )
    
    await update.message.reply_text(profile_text, reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
    return SEARCH_USER_PROFILE

# ================ BALANCE EDIT ================
async def edit_user_balance_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    
    await update.message.reply_text(
        "✍️ যে ব্যবহারকারীর ব্যালেন্স পরিবর্তন করতে চান তার ইউজারনাম বা ইউজার আইডি লিখুন:",
        reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True)
    )
    return SEARCH_USER_FOR_BALANCE

async def search_user_for_balance(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
        
    search_term = update.message.text.strip().lstrip('@')
    
    if search_term == "🔙 Admin Panel":
        return await back_to_admin_panel_handler(update, context)

    found_user_id = None
    
    if search_term.isdigit():
        search_id = int(search_term)
        if search_id in user_info:
            found_user_id = search_id
    
    if not found_user_id:
        for user_id, info in user_info.items():
            if info.get("username") and info["username"].lower() == search_term.lower():
                found_user_id = user_id
                break

    if not found_user_id:
        await update.message.reply_text(
            "❌ এই ইউজারনেম বা ইউজার আইডি এর কোনো ব্যবহারকারী পাওয়া যায়নি। আবার চেষ্টা করুন।",
            reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True)
        )
        return SEARCH_USER_FOR_BALANCE

    context.user_data['edit_balance_user_id'] = found_user_id
    
    user_data = user_info.get(found_user_id, {})
    username = user_data.get('username', 'N/A')
    current_balance = balances.get(found_user_id, 0)

    keyboard = [
        [KeyboardButton("➕ Add Balance"), KeyboardButton("➖ Remove Balance")],
        [KeyboardButton("✍️ Set New Balance")],
        [KeyboardButton("🔙 Admin Panel")]
    ]
    
    await update.message.reply_text(
        f"👤 ব্যবহারকারী: @{username} (আইডি: {found_user_id})\n"
        f"💰 বর্তমান ব্যালেন্স: {current_balance} টাকা\n\n"
        f"আপনি কী করতে চান?",
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )
    return BALANCE_EDIT_ACTION

async def balance_edit_action_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU

    action_text = update.message.text
    
    if action_text == "🔙 Admin Panel":
        return await back_to_admin_panel_handler(update, context)

    if action_text not in ["➕ Add Balance", "➖ Remove Balance", "✍️ Set New Balance"]:
        await update.message.reply_text("❌ অনুগ্রহ করে নিচের বাটন থেকে একটি অপশন বেছে নিন।")
        return BALANCE_EDIT_ACTION

    context.user_data['balance_edit_action'] = action_text
    
    if action_text == "➕ Add Balance":
        prompt = "✍️ কত টাকা যোগ করতে চান তা লিখুন:"
    elif action_text == "➖ Remove Balance":
        prompt = "✍️ কত টাকা সরাতে চান তা লিখুন:"
    else:
        prompt = "✍️ নতুন ব্যালেন্স কত হবে তা লিখুন:"
        
    await update.message.reply_text(
        prompt,
        reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True)
    )
    return RECEIVE_BALANCE_EDIT_AMOUNT

async def receive_balance_edit_amount(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU

    amount_str = update.message.text.strip()
    
    if amount_str == "🔙 Admin Panel":
        context.user_data.pop('edit_balance_user_id', None)
        context.user_data.pop('balance_edit_action', None)
        return await back_to_admin_panel_handler(update, context)

    if not amount_str.isdigit() or float(amount_str) < 0:
        await update.message.reply_text(
            "❌ অনুগ্রহ করে একটি ধনাত্মক সংখ্যা লিখুন।",
            reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True)
        )
        return RECEIVE_BALANCE_EDIT_AMOUNT
        
    amount = float(amount_str)
    user_id = context.user_data.get('edit_balance_user_id')
    action = context.user_data.get('balance_edit_action')

    if not user_id or not action:
        await update.message.reply_text("❌ একটি ত্রুটি ঘটেছে। অনুগ্রহ করে আবার শুরু করুন।")
        return await back_to_admin_panel_handler(update, context)

    old_balance = balances.get(user_id, 0)
    new_balance = 0
    
    if action == "➕ Add Balance":
        new_balance = old_balance + amount
        balances[user_id] = new_balance
        user_message = f"✅ অ্যাডমিন আপনার ব্যালেন্সে {amount} টাকা যোগ করেছে।\nআপনার নতুন ব্যালেন্স: {new_balance} টাকা।"
        admin_message = f"✅ ব্যবহারকারীর ব্যালেন্সে {amount} টাকা যোগ করা হয়েছে।\nনতুন ব্যালেন্স: {new_balance} টাকা।"

    elif action == "➖ Remove Balance":
        new_balance = old_balance - amount
        if new_balance < 0:
            new_balance = 0
        balances[user_id] = new_balance
        user_message = f"✅ অ্যাডমিন আপনার ব্যালেন্স থেকে {amount} টাকা সরিয়ে নিয়েছে।\nআপনার নতুন ব্যালেন্স: {new_balance} টাকা।"
        admin_message = f"✅ ব্যবহারকারীর ব্যালেন্স থেকে {amount} টাকা সরানো হয়েছে।\nনতুন ব্যালেন্স: {new_balance} টাকা।"

    elif action == "✍️ Set New Balance":
        new_balance = amount
        balances[user_id] = new_balance
        user_message = f"✅ অ্যাডমিন আপনার নতুন ব্যালেন্স {new_balance} টাকা সেট করেছে।"
        admin_message = f"✅ ব্যবহারকারীর নতুন ব্যালেন্স {new_balance} টাকা সেট করা হয়েছে।"

    save_user_data()

    try:
        await context.bot.send_message(chat_id=user_id, text=user_message)
    except Exception as e:
        logger.error(f"Failed to notify user {user_id} about balance change: {e}")
        admin_message += "\n⚠️ ব্যবহারকারীকে নোটিশ পাঠানো যায়নি।"

    await update.message.reply_text(admin_message)
    
    context.user_data.pop('edit_balance_user_id', None)
    context.user_data.pop('balance_edit_action', None)
    
    return await show_dashboard(update, context)

# ================ ADMIN PANEL ================
async def back_to_admin_panel_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    return await show_dashboard(update, context)

async def admin_panel_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU

    text = update.message.text
    
    if text == "🔙 Back to Main Menu":
        return await start(update, context)
    
    if text == "🔄 Refresh Dashboard":
        return await show_dashboard(update, context)
    
    if text == "👥 User Profile":
        return await view_user_profile(update, context)

    if text == "✏️ Edit Balance":
        return await edit_user_balance_start(update, context)
        
    if text == "📢 Send Notice":
        await update.message.reply_text("✍️ যে নোটিশটি পাঠাতে চান তা লিখুন:", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
        return SEND_NOTICE
        
    if text == "📂 Manage Categories":
        keyboard = []
        for cat in categories.keys():
            stock_count = get_total_stock(cat)
            keyboard.append([KeyboardButton(f"{cat} ({stock_count})")])
        
        keyboard.append([KeyboardButton("➕ Add Main Category")])
        keyboard.append([KeyboardButton("➖ Remove Main Category")])
        keyboard.append([KeyboardButton("🔙 Admin Panel")])
        await update.message.reply_text("⚙️ প্রধান ক্যাটাগরি ব্যবস্থাপনা:", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))
        return MANAGE_CATEGORY

    if text == "💰 Edit Price":
        keyboard = [[KeyboardButton(cat)] for cat in categories.keys()]
        keyboard.append([KeyboardButton("🔙 Admin Panel")])
        await update.message.reply_text("✍️ কোন ক্যাটাগরির মূল্য পরিবর্তন করবেন?", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))
        return EDIT_PRICE_MAIN

    if text == "💳 Payment Methods":
        return await admin_manage_payment_methods(update, context)

    if text == "💳 Payment Categories":
        return await manage_payment_categories_handler(update, context)
    
    if text == "📱 SMS Auto Deposit":
        return await show_pending_deposits(update, context)
    
    if text == "📋 SMS Log":
        return await show_sms_log(update, context)
         
    return ADMIN_PANEL

# ================ SMS LOG VIEWER ================
async def show_sms_log(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    
    try:
        with open(SMS_LOG_FILE, "r", encoding='utf-8') as f:
            lines = f.readlines()
        
        if not lines:
            await update.message.reply_text("📋 এখনো কোনো SMS লগ নেই।")
            return ADMIN_PANEL
        
        # সর্বশেষ ২০টি লগ দেখান
        recent_lines = lines[-20:]
        log_text = "📋 সর্বশেষ SMS লগ:\n\n" + "".join(recent_lines)
        
        if len(log_text) > 4000:
            parts = [log_text[i:i+4000] for i in range(0, len(log_text), 4000)]
            for part in parts:
                await update.message.reply_text(part)
        else:
            await update.message.reply_text(log_text)
            
    except FileNotFoundError:
        await update.message.reply_text("📋 এখনো কোনো SMS লগ নেই।")
    
    return ADMIN_PANEL

# ================ SHOW PENDING DEPOSITS ================
async def show_pending_deposits(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    
    if not pending_deposits:
        await update.message.reply_text(
            "✅ কোনো পেন্ডিং ডিপোজিট নেই।\n\n"
            "📱 SMS Auto Deposit সিস্টেম সক্রিয় আছে।\n"
            "যখন অ্যাডমিনের ফোনে bKash SMS আসবে, TRXID মিলিয়ে অটো ডিপোজিট হবে।",
            reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True)
        )
        return ADMIN_PANEL
    
    pending_text = "⏳ পেন্ডিং ডিপোজিট সমূহ:\n\n"
    for trxid, data in pending_deposits.items():
        user_id = data['user_id']
        user = user_info.get(user_id, {})
        username = user.get('username', 'N/A')
        pending_text += f"🔑 TRXID: {trxid}\n"
        pending_text += f"👤 ইউজার: @{username} (ID: {user_id})\n"
        pending_text += f"💰 পরিমাণ: {data['amount']} টাকা\n"
        pending_text += f"💳 মেথড: {data.get('method', 'Unknown')}\n"
        pending_text += f"⏰ সময়: {time.strftime('%H:%M %b %d', time.localtime(data['timestamp']))}\n"
        pending_text += "---------------------------\n"
    
    pending_text += f"\n📱 মোট পেন্ডিং: {len(pending_deposits)} টি"
    
    if len(pending_text) > 4000:
        parts = [pending_text[i:i+4000] for i in range(0, len(pending_text), 4000)]
        for part in parts:
            await update.message.reply_text(part)
    else:
        await update.message.reply_text(pending_text)
    
    return ADMIN_PANEL

# ================ ADMIN PAYMENT METHODS MANAGEMENT ================
async def admin_manage_payment_methods(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    
    keyboard = [
        [KeyboardButton("➕ Add Payment Method")],
        [KeyboardButton("📋 View All Methods")],
        [KeyboardButton("🗑️ Delete Payment Method")],
        [KeyboardButton("🔙 Admin Panel")]
    ]
    
    await update.message.reply_text(
        "💳 পেমেন্ট মেথড ম্যানেজমেন্ট\n\n"
        "আপনি নতুন পেমেন্ট মেথড যোগ করতে পারেন, "
        "সব মেথড দেখতে পারেন, অথবা কোনো মেথড ডিলিট করতে পারেন।",
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )
    return ADMIN_VIEW_PAYMENT_METHODS

async def admin_add_payment_method(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    
    await update.message.reply_text(
        "✍️ নতুন পেমেন্ট মেথডের নাম লিখুন:\n"
        "(যেমন: বিকাশ, নগদ, রকেট, ব্যাংক ট্রান্সফার ইত্যাদি)",
        reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True)
    )
    return ADMIN_ADD_PAYMENT_METHOD

async def admin_receive_payment_method_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    
    method_name = update.message.text.strip()
    
    if method_name == "🔙 Admin Panel":
        return await back_to_admin_panel_handler(update, context)
    
    if method_name in payment_methods:
        await update.message.reply_text(
            f"⚠️ '{method_name}' নামে একটি পেমেন্ট মেথড ইতিমধ্যে আছে।\n"
            f"আলাদা নাম ব্যবহার করুন।",
            reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True)
        )
        return ADMIN_ADD_PAYMENT_METHOD
    
    context.user_data['new_payment_method_name'] = method_name
    
    await update.message.reply_text(
        f"✍️ '{method_name}' এর জন্য বিস্তারিত তথ্য লিখুন:\n"
        f"(যেমন: নাম্বার, একাউন্ট নাম, রেফারেন্স ইত্যাদি)\n\n"
        f"📝 প্রতিটি লাইনে একটি করে তথ্য লিখুন।",
        reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True)
    )
    return ADMIN_RECEIVE_PAYMENT_DETAILS

async def admin_receive_payment_method_details(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    
    details = update.message.text.strip()
    
    if details == "🔙 Admin Panel":
        context.user_data.pop('new_payment_method_name', None)
        return await back_to_admin_panel_handler(update, context)
    
    method_name = context.user_data.get('new_payment_method_name')
    
    if not method_name:
        await update.message.reply_text("❌ একটি ত্রুটি ঘটেছে। অনুগ্রহ করে আবার শুরু করুন।")
        return await back_to_admin_panel_handler(update, context)
    
    payment_methods[method_name] = {
        "details": details,
        "active": True
    }
    
    save_user_data()
    
    await update.message.reply_text(
        f"✅ '{method_name}' পেমেন্ট মেথড সফলভাবে যোগ হয়েছে!\n\n"
        f"📋 বিস্তারিত:\n{details}",
        reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True)
    )
    
    context.user_data.pop('new_payment_method_name', None)
    return await back_to_admin_panel_handler(update, context)

async def admin_view_payment_methods(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    
    text = update.message.text
    
    if text == "🔙 Admin Panel":
        return await back_to_admin_panel_handler(update, context)
    
    if text == "➕ Add Payment Method":
        return await admin_add_payment_method(update, context)
    
    if text == "📋 View All Methods":
        if not payment_methods:
            await update.message.reply_text(
                "⚠️ কোনো পেমেন্ট মেথড নেই।",
                reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True)
            )
            return ADMIN_VIEW_PAYMENT_METHODS
        
        methods_text = "📋 সকল পেমেন্ট মেথড:\n\n"
        for idx, (name, data) in enumerate(payment_methods.items(), 1):
            status = "✅ সক্রিয়" if data.get("active", True) else "❌ নিষ্ক্রিয়"
            methods_text += f"{idx}. {name}\n"
            methods_text += f"   📝 {data['details']}\n"
            methods_text += f"   📊 অবস্থা: {status}\n\n"
        
        if len(methods_text) > 4000:
            parts = [methods_text[i:i+4000] for i in range(0, len(methods_text), 4000)]
            for part in parts:
                await update.message.reply_text(part)
        else:
            await update.message.reply_text(methods_text)
        
        return ADMIN_VIEW_PAYMENT_METHODS
    
    if text == "🗑️ Delete Payment Method":
        if not payment_methods:
            await update.message.reply_text(
                "⚠️ কোনো পেমেন্ট মেথড নেই।",
                reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True)
            )
            return ADMIN_VIEW_PAYMENT_METHODS
        
        keyboard = []
        for method_name in payment_methods.keys():
            keyboard.append([KeyboardButton(f"❌ {method_name}")])
        keyboard.append([KeyboardButton("🔙 Admin Panel")])
        
        await update.message.reply_text(
            "🗑️ কোন পেমেন্ট মেথড ডিলিট করতে চান?",
            reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        )
        return ADMIN_DELETE_PAYMENT_METHOD
    
    return ADMIN_VIEW_PAYMENT_METHODS

async def admin_delete_payment_method(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    
    text = update.message.text
    
    if text == "🔙 Admin Panel":
        return await back_to_admin_panel_handler(update, context)
    
    method_name = text.replace("❌ ", "").strip()
    
    if method_name in payment_methods:
        del payment_methods[method_name]
        save_user_data()
        await update.message.reply_text(
            f"✅ '{method_name}' পেমেন্ট মেথড ডিলিট করা হয়েছে।",
            reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True)
        )
    else:
        await update.message.reply_text(
            f"❌ '{method_name}' পেমেন্ট মেথড পাওয়া যায়নি।",
            reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True)
        )
    
    return await back_to_admin_panel_handler(update, context)

# ================ PAYMENT CATEGORIES ================
async def manage_payment_categories_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
        
    keyboard_inline = []
    if not categories:
        await update.message.reply_text("⚠️ কোনো ক্যাটাগরি নেই।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
        return ADMIN_PANEL

    for main_cat in categories:
        current_status = "Manual 💳" if main_cat in MANUAL_DELIVERY_CATEGORIES else "Balance 💰"
        button_text = "Switch to Balance" if main_cat in MANUAL_DELIVERY_CATEGORIES else "Switch to Manual"
        callback_data = f"toggle_payment:{main_cat}"
        
        keyboard_inline.append([
            InlineKeyboardButton(f"{main_cat} ({current_status})", callback_data="ignore"),
            InlineKeyboardButton(button_text, callback_data=callback_data)
        ])
    
    reply_markup_inline = InlineKeyboardMarkup(keyboard_inline)
    
    await update.message.reply_text(
        "⚡️ পেমেন্ট ক্যাটাগরি নিয়ন্ত্রণ\n\n"
        "নিচের তালিকা থেকে প্রতিটি ক্যাটাগরির জন্য পেমেন্ট পদ্ধতি পরিবর্তন করতে পারেন।\n"
        "Balance Payment মানে ব্যবহারকারী তার ব্যালেন্স থেকে কিনতে পারবে।\n"
        "Manual Payment মানে ব্যবহারকারীকে সরাসরি পেমেন্ট করে স্ক্রিনশট পাঠাতে হবে।",
        reply_markup=reply_markup_inline
    )

    reply_markup_text = ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True)
    await update.message.reply_text(
        "🔙 অ্যাডমিন প্যানেলে ফিরে যেতে নিচের বাটনটি চাপুন।",
        reply_markup=reply_markup_text
    )

    return MANAGE_PAYMENT_CATEGORIES

async def toggle_payment_method(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    
    if update.effective_user.id != ADMIN_ID:
        await query.edit_message_caption("❌ অননুমোদিত।", reply_markup=None)
        return
        
    if data.startswith("toggle_payment:"):
        _, cat_name = data.split(":", 1)
        
        if cat_name in MANUAL_DELIVERY_CATEGORIES:
            MANUAL_DELIVERY_CATEGORIES.remove(cat_name)
        else:
            MANUAL_DELIVERY_CATEGORIES.append(cat_name)
        
        save_user_data()
        
        keyboard = []
        for main_cat in categories:
            current_status = "Manual 💳" if main_cat in MANUAL_DELIVERY_CATEGORIES else "Balance 💰"
            button_text = "Switch to Balance" if main_cat in MANUAL_DELIVERY_CATEGORIES else "Switch to Manual"
            callback_data = f"toggle_payment:{main_cat}"
            keyboard.append([
                InlineKeyboardButton(f"{main_cat} ({current_status})", callback_data="ignore"),
                InlineKeyboardButton(button_text, callback_data=callback_data)
            ])
            
        reply_markup = InlineKeyboardMarkup(keyboard)
        await query.edit_message_text(
            "⚡️ পেমেন্ট ক্যাটাগরি নিয়ন্ত্রণ\n\n"
            "নিচের তালিকা থেকে প্রতিটি ক্যাটাগরির জন্য পেমেন্ট পদ্ধতি পরিবর্তন করতে পারেন।\n"
            "Balance Payment মানে ব্যবহারকারী তার ব্যালেন্স থেকে কিনতে পারবে।\n"
            "Manual Payment মানে ব্যবহারকারীকে সরাসরি পেমেন্ট করে স্ক্রিনশট পাঠাতে হবে।",
            reply_markup=reply_markup
        )
    return

# ================ SEND NOTICE ================
async def send_notice_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
    
    notice_text = update.message.text
    
    notice_count = 0
    failed_users = []
    
    users_to_notify = [uid for uid in user_info.keys() if uid != ADMIN_ID]
    
    if not users_to_notify:
        await update.message.reply_text("⚠️ কোনো ব্যবহারকারীকে নোটিশ পাঠানো হয়নি। সম্ভবত অন্য কোনো ব্যবহারকারী এখনও বট শুরু করেনি।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
        return await back_to_admin_panel_handler(update, context)

    for user_id in users_to_notify:
        try:
            await context.bot.send_message(chat_id=user_id, text=f"📢 নোটিশ:\n\n{notice_text}")
            notice_count += 1
        except Exception:
            failed_users.append(user_id)
            pass

    await update.message.reply_text(f"✅ নোটিশ পাঠানো হয়েছে।\n\nসফল: {notice_count} জন ব্যবহারকারীকে\nব্যর্থ: {len(failed_users)} জন ব্যবহারকারীকে", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
    return await back_to_admin_panel_handler(update, context)

# ================ MANAGE CATEGORIES ================
async def back_to_manage_main_categories_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if "active_main_cat" in context.user_data:
        del context.user_data["active_main_cat"]
    if "active_sub_cat" in context.user_data:
        del context.user_data["active_sub_cat"]
    return await manage_category_handler(update, context)

async def manage_category_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU

    text = update.message.text
    
    if text == "🔙 Admin Panel":
        return await back_to_admin_panel_handler(update, context)

    if text == "➕ Add Main Category":
        await update.message.reply_text("✍️ নতুন প্রধান ক্যাটাগরির নাম পাঠান:", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
        return ADD_MAIN_CAT

    if text == "➖ Remove Main Category":
        keyboard = []
        for cat in categories.keys():
            stock_count = get_total_stock(cat)
            keyboard.append([KeyboardButton(f"{cat} ({stock_count})")])
        
        keyboard.append([KeyboardButton("🔙 Admin Panel")])
        await update.message.reply_text("➖ কোন প্রধান ক্যাটাগরি সরাতে চান?", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))
        return REMOVE_MAIN_CAT
    
    original_cat = text.split(" (")[0]
    
    if original_cat in categories:
        context.user_data["active_main_cat"] = original_cat
        
        keyboard = []
        if categories.get(original_cat):
            for sub_cat in categories[original_cat]:
                stock_count = count_items(original_cat, sub_cat)
                keyboard.append([KeyboardButton(f"{sub_cat} ({stock_count})")])
        else:
            keyboard.append([KeyboardButton("⚠️ No Sub Categories")])
        
        keyboard.append([KeyboardButton("➕ Add Sub Category")])
        keyboard.append([KeyboardButton("➖ Remove Sub Category")])
        keyboard.append([KeyboardButton("➕ Add Items (TXT)")])
        keyboard.append([KeyboardButton("🔙 Manage Categories"), KeyboardButton("🔙 Admin Panel")])
        
        await update.message.reply_text(f"⚙️ {original_cat} এর সাব-ক্যাটাগরি:", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))
        return MANAGE_SUB_CATEGORY
    
    return MANAGE_CATEGORY

async def add_main_category(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU

    new_cat = update.message.text.strip()
    
    if new_cat == "🔙 Admin Panel":
        return await back_to_admin_panel_handler(update, context)
        
    if new_cat in categories:
        await update.message.reply_text("⚠️ এই প্রধান ক্যাটাগরি আগে থেকেই আছে।")
    else:
        categories[new_cat] = []
        save_user_data()
        await update.message.reply_text(f"✅ প্রধান ক্যাটাগরি '{new_cat}' যোগ হয়েছে।")
    
    return await manage_category_handler(update, context)

async def remove_main_category(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU

    cat_to_remove = update.message.text.split(" (")[0].strip()
    
    if cat_to_remove == "🔙 Admin Panel":
        return await back_to_admin_panel_handler(update, context)
        
    if cat_to_remove in categories:
        if cat_to_remove in categories:
            for sub_cat in categories[cat_to_remove]:
                txt_path = get_txt_path(cat_to_remove, sub_cat)
                if os.path.exists(txt_path):
                    os.remove(txt_path)
                excel_path = get_excel_path(cat_to_remove, sub_cat)
                if os.path.exists(excel_path):
                    os.remove(excel_path)
                if cat_to_remove in prices and sub_cat in prices[cat_to_remove]:
                    del prices[cat_to_remove][sub_cat]
        if cat_to_remove in prices:
            del prices[cat_to_remove]
        
        del categories[cat_to_remove]
        if cat_to_remove in MANUAL_DELIVERY_CATEGORIES:
            MANUAL_DELIVERY_CATEGORIES.remove(cat_to_remove)
            
        save_user_data()
        await update.message.reply_text(f"✅ প্রধান ক্যাটাগরি '{cat_to_remove}' সরানো হয়েছে।")
    else:
        await update.message.reply_text("⚠️ এই ক্যাটাগরি পাওয়া যায়নি।")
    
    return await manage_category_handler(update, context)

async def manage_sub_category_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
        
    text = update.message.text
    main_cat = context.user_data.get("active_main_cat")
    
    if text == "🔙 Manage Categories":
        if "active_main_cat" in context.user_data:
            del context.user_data["active_main_cat"]
        if "active_sub_cat" in context.user_data:
            del context.user_data["active_sub_cat"]
        return await manage_category_handler(update, context)
    
    if text == "🔙 Admin Panel":
        return await back_to_admin_panel_handler(update, context)

    if text == "➕ Add Sub Category":
        await update.message.reply_text("✍️ নতুন সাব-ক্যাটাগরির নাম পাঠান:", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Manage Categories"), KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
        return ADD_SUB_CAT
        
    if text == "➖ Remove Sub Category":
        if not categories.get(main_cat, []):
            await update.message.reply_text("⚠️ কোনো সাব-ক্যাটাগরি নেই।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Manage Categories")]], resize_keyboard=True))
            return MANAGE_SUB_CATEGORY
            
        keyboard = [[KeyboardButton(sub_cat)] for sub_cat in categories.get(main_cat, [])]
        keyboard.append([KeyboardButton("🔙 Manage Categories"), KeyboardButton("🔙 Admin Panel")])
        await update.message.reply_text("➖ কোন সাব-ক্যাটাগরি সরাতে চান?", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))
        return REMOVE_SUB_CAT
        
    if text == "➕ Add Items (TXT)":
        if not categories.get(main_cat, []):
            await update.message.reply_text("⚠️ কোনো সাব-ক্যাটাগরি নেই। আগে সাব-ক্যাটাগরি তৈরি করুন।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Manage Categories")]], resize_keyboard=True))
            return MANAGE_SUB_CATEGORY
            
        context.user_data["active_sub_cat"] = None
        keyboard = [[KeyboardButton(sub_cat)] for sub_cat in categories.get(main_cat, [])]
        keyboard.append([KeyboardButton("🔙 Manage Categories"), KeyboardButton("🔙 Admin Panel")])
        await update.message.reply_text(
            f"📁 {main_cat} ক্যাটাগরির জন্য সাব-ক্যাটাগরি নির্বাচন করুন:\n\n"
            f"নিচের সাব-ক্যাটাগরি থেকে একটি নির্বাচন করুন যেখানে আইটেম যোগ করতে চান:",
            reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        )
        return ADD_ITEMS_TXT
    
    original_sub_cat = text.split(" (")[0]
    
    if original_sub_cat in categories.get(main_cat, []):
        context.user_data["active_sub_cat"] = original_sub_cat
        count = count_items(main_cat, original_sub_cat)
        
        txt_path = get_txt_path(main_cat, original_sub_cat)
        preview = "কোনো আইটেম নেই"
        if os.path.exists(txt_path):
            with open(txt_path, 'r', encoding='utf-8') as f:
                items = [line.strip() for line in f.readlines() if line.strip()]
            if items:
                preview = "\n".join(items[:5])
                if len(items) > 5:
                    preview += f"\n... এবং আরও {len(items) - 5}টি আইটেম"
        
        keyboard = [
            [KeyboardButton("➕ Add Items (TXT)")],
            [KeyboardButton("🔙 Manage Categories"), KeyboardButton("🔙 Admin Panel")]
        ]
        await update.message.reply_text(
            f"⚙️ সাব-ক্যাটাগরি: {original_sub_cat}\n"
            f"📦 মোট আইটেম: {count} টি\n"
            f"📄 আইটেম প্রিভিউ:\n{preview}\n\n"
            f"আইটেম যোগ করতে '➕ Add Items (TXT)' চাপুন।",
            reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        )
        return ADD_ITEMS_TXT
        
    return MANAGE_SUB_CATEGORY

async def add_sub_category(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU

    new_sub_cat = update.message.text.strip()
    main_cat = context.user_data.get("active_main_cat")
    
    if new_sub_cat == "🔙 Manage Categories":
        await back_to_manage_main_categories_handler(update, context)
        return MANAGE_CATEGORY

    if new_sub_cat == "🔙 Admin Panel":
        return await back_to_admin_panel_handler(update, context)

    if main_cat and new_sub_cat not in categories[main_cat]:
        categories[main_cat].append(new_sub_cat)
        ensure_txt_file(main_cat, new_sub_cat)
        save_user_data()
        await update.message.reply_text(f"✅ সাব-ক্যাটাগরি '{new_sub_cat}' যোগ হয়েছে।")
    else:
        await update.message.reply_text("⚠️ এই সাব-ক্যাটাগরি আগে থেকেই আছে অথবা কোনো প্রধান ক্যাটাগরি নির্বাচন করা হয়নি।")
    
    keyboard = []
    if categories.get(main_cat):
        for sub_cat in categories[main_cat]:
            stock_count = count_items(main_cat, sub_cat)
            keyboard.append([KeyboardButton(f"{sub_cat} ({stock_count})")])
    
    keyboard.append([KeyboardButton("➕ Add Sub Category")])
    keyboard.append([KeyboardButton("➖ Remove Sub Category")])
    keyboard.append([KeyboardButton("➕ Add Items (TXT)")])
    keyboard.append([KeyboardButton("🔙 Manage Categories"), KeyboardButton("🔙 Admin Panel")])
    
    await update.message.reply_text(
        f"⚙️ {main_cat} এর সাব-ক্যাটাগরি:", 
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )
    return MANAGE_SUB_CATEGORY

async def remove_sub_category(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU

    sub_cat_to_remove = update.message.text.split(" (")[0].strip()
    main_cat = context.user_data.get("active_main_cat")
    
    if sub_cat_to_remove == "🔙 Manage Categories":
        await back_to_manage_main_categories_handler(update, context)
        return MANAGE_CATEGORY

    if sub_cat_to_remove == "🔙 Admin Panel":
        return await back_to_admin_panel_handler(update, context)
        
    if main_cat and sub_cat_to_remove in categories.get(main_cat, []):
        categories[main_cat].remove(sub_cat_to_remove)
        txt_path = get_txt_path(main_cat, sub_cat_to_remove)
        if os.path.exists(txt_path):
            os.remove(txt_path)
        excel_path = get_excel_path(main_cat, sub_cat_to_remove)
        if os.path.exists(excel_path):
            os.remove(excel_path)
        if main_cat in prices and sub_cat_to_remove in prices[main_cat]:
            del prices[main_cat][sub_cat_to_remove]
        save_user_data()
        await update.message.reply_text(f"✅ সাব-ক্যাটাগরি '{sub_cat_to_remove}' সরানো হয়েছে।")
    else:
        await update.message.reply_text("⚠️ এই সাব-ক্যাটাগরি পাওয়া যায়নি।")
    
    keyboard = []
    if categories.get(main_cat):
        for sub_cat in categories[main_cat]:
            stock_count = count_items(main_cat, sub_cat)
            keyboard.append([KeyboardButton(f"{sub_cat} ({stock_count})")])
    else:
        keyboard.append([KeyboardButton("⚠️ No Sub Categories")])
    
    keyboard.append([KeyboardButton("➕ Add Sub Category")])
    keyboard.append([KeyboardButton("➖ Remove Sub Category")])
    keyboard.append([KeyboardButton("➕ Add Items (TXT)")])
    keyboard.append([KeyboardButton("🔙 Manage Categories"), KeyboardButton("🔙 Admin Panel")])
    
    await update.message.reply_text(
        f"⚙️ {main_cat} এর সাব-ক্যাটাগরি:", 
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )
    return MANAGE_SUB_CATEGORY

# ================ ADD ITEMS VIA TXT ================
async def add_items_txt_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU

    text = update.message.text
    
    if text == "🔙 Manage Categories":
        await back_to_manage_main_categories_handler(update, context)
        return MANAGE_CATEGORY

    if text == "🔙 Admin Panel":
        return await back_to_admin_panel_handler(update, context)
    
    main_cat = context.user_data.get("active_main_cat")
    
    if not context.user_data.get("active_sub_cat"):
        if text in categories.get(main_cat, []):
            context.user_data["active_sub_cat"] = text
            await update.message.reply_text(
                f"📁 {main_cat} → {text}\n\n"
                f"এখন টেক্সট ফাইল আপলোড করুন যা আইটেম ধারণ করে।\n"
                f"প্রতি লাইনে একটি আইটেম থাকতে হবে।\n\n"
                f"ফাইল আপলোড করুন অথবা '✅ Done' চাপুন।",
                reply_markup=ReplyKeyboardMarkup(
                    [[KeyboardButton("✅ Done"), KeyboardButton("🔙 Manage Categories"), KeyboardButton("🔙 Admin Panel")]],
                    resize_keyboard=True
                )
            )
            return ADD_ITEMS_TXT
        else:
            await update.message.reply_text(
                "❌ অনুগ্রহ করে একটি বৈধ সাব-ক্যাটাগরি নির্বাচন করুন।",
                reply_markup=ReplyKeyboardMarkup(
                    [[KeyboardButton(sub_cat)] for sub_cat in categories.get(main_cat, [])] + 
                    [[KeyboardButton("🔙 Manage Categories"), KeyboardButton("🔙 Admin Panel")]],
                    resize_keyboard=True
                )
            )
            return ADD_ITEMS_TXT
    
    sub_cat = context.user_data.get("active_sub_cat")
    
    if text == "✅ Done":
        count = count_items(main_cat, sub_cat)
        await update.message.reply_text(
            f"✅ '{sub_cat}' তে মোট {count} টি আইটেম আছে।",
            reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True)
        )
        context.user_data.pop("active_sub_cat", None)
        return await back_to_admin_panel_handler(update, context)
    
    if update.message.document:
        file = update.message.document
        if file.file_name.endswith('.txt'):
            try:
                file_obj = await file.get_file()
                file_content = await file_obj.download_as_bytearray()
                txt_content = file_content.decode('utf-8')
                
                add_items_from_txt(main_cat, sub_cat, txt_content)
                
                count = count_items(main_cat, sub_cat)
                await update.message.reply_text(
                    f"✅ '{sub_cat}' তে আইটেম যোগ হয়েছে।\n"
                    f"বর্তমান মোট আইটেম: {count} টি\n\n"
                    f"আরও ফাইল আপলোড করতে পারেন অথবা '✅ Done' চাপুন।",
                    reply_markup=ReplyKeyboardMarkup(
                        [[KeyboardButton("✅ Done"), KeyboardButton("🔙 Manage Categories"), KeyboardButton("🔙 Admin Panel")]],
                        resize_keyboard=True
                    )
                )
                return ADD_ITEMS_TXT
                
            except Exception as e:
                logger.error(f"Error processing TXT file: {e}")
                await update.message.reply_text(
                    "❌ ফাইল প্রসেস করতে সমস্যা হয়েছে। অনুগ্রহ করে সঠিক টেক্সট ফাইল আপলোড করুন।",
                    reply_markup=ReplyKeyboardMarkup(
                        [[KeyboardButton("✅ Done"), KeyboardButton("🔙 Manage Categories"), KeyboardButton("🔙 Admin Panel")]],
                        resize_keyboard=True
                    )
                )
                return ADD_ITEMS_TXT
        else:
            await update.message.reply_text(
                "❌ অনুগ্রহ করে শুধু .txt ফাইল আপলোড করুন।",
                reply_markup=ReplyKeyboardMarkup(
                    [[KeyboardButton("✅ Done"), KeyboardButton("🔙 Manage Categories"), KeyboardButton("🔙 Admin Panel")]],
                    resize_keyboard=True
                )
            )
            return ADD_ITEMS_TXT
    
    if text and text not in ["✅ Done", "🔙 Manage Categories", "🔙 Admin Panel"]:
        add_items_from_txt(main_cat, sub_cat, text)
        count = count_items(main_cat, sub_cat)
        await update.message.reply_text(
            f"✅ '{sub_cat}' তে আইটেম যোগ হয়েছে।\n"
            f"বর্তমান মোট আইটেম: {count} টি\n\n"
            f"আরও টেক্সট পাঠাতে পারেন অথবা '✅ Done' চাপুন।",
            reply_markup=ReplyKeyboardMarkup(
                [[KeyboardButton("✅ Done"), KeyboardButton("🔙 Manage Categories"), KeyboardButton("🔙 Admin Panel")]],
                resize_keyboard=True
            )
        )
        return ADD_ITEMS_TXT
    
    return ADD_ITEMS_TXT

# ================ EDIT PRICE ================
async def edit_payment_info(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
        
    global payment_info
    new_info = update.message.text.strip()
    
    if new_info == "🔙 Admin Panel":
        return await back_to_admin_panel_handler(update, context)
        
    payment_info = new_info
    await update.message.reply_text("✅ পেমেন্ট তথ্য সফলভাবে আপডেট হয়েছে।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
    return await back_to_admin_panel_handler(update, context)

async def edit_price_main_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU

    text = update.message.text.strip()
    
    if text == "🔙 Admin Panel":
        return await back_to_admin_panel_handler(update, context)
    
    original_cat = text.split(" (")[0]
        
    if original_cat in categories.keys():
        context.user_data['temp_main_cat_for_price'] = original_cat
        keyboard = [[KeyboardButton(sub_cat)] for sub_cat in categories[original_cat]]
        keyboard.append([KeyboardButton("🔙 Admin Panel")])
        await update.message.reply_text(f"✍️ কোন সাব-ক্যাটাগরির মূল্য পরিবর্তন করবেন?", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))
        return EDIT_PRICE_SUB
        
    await update.message.reply_text("❌ এই ক্যাটাগরি পাওয়া যায়নি।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
    return EDIT_PRICE_MAIN

async def edit_price_sub_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU
        
    text = update.message.text.strip()
    main_cat = context.user_data.get('temp_main_cat_for_price')
    
    if text == "🔙 Admin Panel":
        if 'temp_main_cat_for_price' in context.user_data:
            del context.user_data['temp_main_cat_for_price']
        return await back_to_admin_panel_handler(update, context)
        
    if main_cat and text in categories.get(main_cat, []):
        context.user_data['temp_sub_cat_for_price'] = text
        current_price = prices.get(main_cat, {}).get(text, "সেট করা হয়নি")
        await update.message.reply_text(f"✍️ '{text}' এর বর্তমান মূল্য: {current_price}\nনতুন মূল্য লিখুন:", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
        return RECEIVE_NEW_PRICE
        
    await update.message.reply_text("❌ এই সাব-ক্যাটাগরি পাওয়া যায়নি।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
    return EDIT_PRICE_SUB

async def receive_price(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return MAIN_MENU

    price_text = update.message.text.strip()
    
    if price_text == "🔙 Admin Panel":
        if 'temp_main_cat_for_price' in context.user_data:
            del context.user_data['temp_main_cat_for_price']
        if 'temp_sub_cat_for_price' in context.user_data:
            del context.user_data['temp_sub_cat_for_price']
        return await back_to_admin_panel_handler(update, context)

    try:
        new_price = float(price_text)
        main_cat = context.user_data.get('temp_main_cat_for_price')
        sub_cat = context.user_data.get('temp_sub_cat_for_price')
        if main_cat and sub_cat:
            if main_cat not in prices:
                prices[main_cat] = {}
            prices[main_cat][sub_cat] = new_price
            save_user_data()
            await update.message.reply_text(f"✅ '{sub_cat}' এর মূল্য এখন {new_price} টাকা।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
            if 'temp_main_cat_for_price' in context.user_data:
                del context.user_data['temp_main_cat_for_price']
            if 'temp_sub_cat_for_price' in context.user_data:
                del context.user_data['temp_sub_cat_for_price']
            return await back_to_admin_panel_handler(update, context)
    except ValueError:
        await update.message.reply_text("❌ মূল্য শুধুমাত্র সংখ্যায় লিখুন।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Admin Panel")]], resize_keyboard=True))
        return RECEIVE_NEW_PRICE
        
    return await back_to_admin_panel_handler(update, context)

# ================ BUY FLOW ================
async def back_to_categories_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await check_subscription(update, context):
        return ConversationHandler.END

    keyboard = []
    for cat in categories.keys():
        keyboard.append([KeyboardButton(cat)])
    
    keyboard.append([KeyboardButton("🔙 Back to Main Menu")])
    await update.message.reply_text("🛒 ক্যাটাগরি বেছে নিন:", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))
    return BUY_MENU

async def back_to_subcategories_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await check_subscription(update, context):
        return ConversationHandler.END
    
    main_cat = context.user_data.get('temp_main_cat_for_buy')
    if not main_cat:
        return await back_to_categories_handler(update, context)
        
    keyboard = []
    for sub_cat in categories[main_cat]:
        stock_count = count_items(main_cat, sub_cat)
        keyboard.append([KeyboardButton(f"{sub_cat} ({stock_count})")])
        
    keyboard.append([KeyboardButton("🔙 Back to Categories")])
    keyboard.append([KeyboardButton("🔙 Back to Main Menu")])

    await update.message.reply_text(f"🛒 {main_cat} এর সাব-ক্যাটাগরি বেছে নিন:", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))
    return BUY_SUB_MENU

async def user_choose_category(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await check_subscription(update, context):
        return ConversationHandler.END
        
    text = update.message.text.strip()
    original_cat = text.split(" (")[0]
    
    if original_cat == "🔙 Back to Main Menu":
        return await start(update, context)

    context.user_data.clear()

    if original_cat in categories.keys():
        context.user_data["temp_main_cat_for_buy"] = original_cat
        
        keyboard = []
        for sub_cat in categories[original_cat]:
            stock_count = count_items(original_cat, sub_cat)
            keyboard.append([KeyboardButton(f"{sub_cat} ({stock_count})")])
        
        keyboard.append([KeyboardButton("🔙 Back to Categories")])
        keyboard.append([KeyboardButton("🔙 Back to Main Menu")])
        
        await update.message.reply_text(f"🛒 {original_cat} এর সাব-ক্যাটাগরি বেছে নিন:", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))
        return BUY_SUB_MENU

    await update.message.reply_text("❌ এই ক্যাটাগরি পাওয়া যায়নি।")
    return BUY_MENU

async def user_choose_subcategory(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await check_subscription(update, context):
        return ConversationHandler.END
        
    text = update.message.text.strip()
    original_sub_cat = text.split(" (")[0]
    main_cat = context.user_data.get('temp_main_cat_for_buy')
    
    if text == "🔙 Back to Categories":
        return await back_to_categories_handler(update, context)
    if text == "🔙 Back to Main Menu":
        return await start(update, context)
        
    if main_cat and original_sub_cat in categories.get(main_cat, []):
        context.user_data["order"] = {"main_cat": main_cat, "sub_cat": original_sub_cat}
        
        price = prices.get(main_cat, {}).get(original_sub_cat, "মূল্য এখনো সেট করা হয়নি।")
        
        keyboard = [
            [KeyboardButton("🔙 Back to Sub Categories")],
            [KeyboardButton("🔙 Back to Main Menu")]
        ]
        
        await update.message.reply_text(f"✅ আপনি {original_sub_cat} সিলেক্ট করেছেন।\n💰 প্রতিটির দাম: {price} টাকা\n\n✍️ অনুগ্রহ করে আপনি কতগুলি চান তা সংখ্যায় লিখুন:", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))
        return GET_QUANTITY

    await update.message.reply_text("❌ এই সাব-ক্যাটাগরি পাওয়া যায়নি।")
    return BUY_SUB_MENU

async def receive_quantity(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await check_subscription(update, context):
        return ConversationHandler.END

    qty = update.message.text.strip()

    if qty == "🔙 Back to Sub Categories":
        return await back_to_subcategories_handler(update, context)
    if qty == "🔙 Back to Main Menu":
        return await start(update, context)
    
    if not qty.isdigit():
        await update.message.reply_text("❌ অনুগ্রহ করে শুধু সংখ্যা লিখুন।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Back to Sub Categories"), KeyboardButton("🔙 Back to Main Menu")]], resize_keyboard=True))
        return GET_QUANTITY
        
    qty = int(qty)
    order = context.user_data.get("order", {})
    order["qty"] = qty
    
    main_cat = order['main_cat']
    sub_cat = order['sub_cat']
    
    price_per_item = prices.get(main_cat, {}).get(sub_cat, 0)
    total_price = price_per_item * qty
    
    order["price"] = total_price
    context.user_data["order"] = order

    is_manual = main_cat in MANUAL_DELIVERY_CATEGORIES

    if not is_manual:
        user_id = update.effective_user.id
        current_balance = balances.get(user_id, 0)
    
        if current_balance >= total_price:
            keyboard = [
                [KeyboardButton("✅ Confirm Purchase")],
                [KeyboardButton("🔙 Back to Sub Categories")],
                [KeyboardButton("🔙 Back to Main Menu")]
            ]
            await update.message.reply_text(
                f"✅ অর্ডার তৈরি হয়েছে।\n"
                f"ক্যাটাগরি: {order['sub_cat']}\n"
                f"পরিমাণ: {order['qty']} টি\n"
                f"মোট দাম: {total_price} টাকা\n"
                f"আপনার বর্তমান ব্যালেন্স: {current_balance} টাকা\n\n"
                f"আপনার ব্যালেন্স থেকে পেমেন্ট করতে '✅ Confirm Purchase' চাপুন।",
                reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
            )
            return CONFIRM_ORDER
        else:
            await update.message.reply_text(f"❌ আপনার যথেষ্ট ব্যালেন্স নেই। আপনার প্রয়োজন {total_price} টাকা কিন্তু ব্যালেন্স আছে {current_balance} টাকা।", reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Back to Sub Categories"), KeyboardButton("🔙 Back to Main Menu")]], resize_keyboard=True))
            return BUY_SUB_MENU
    
    # For manual delivery categories, show payment methods if available
    if payment_methods:
        keyboard = []
        for method_name in payment_methods.keys():
            keyboard.append([KeyboardButton(f"💳 {method_name}")])
        keyboard.append([KeyboardButton("🔙 Back to Sub Categories")])
        keyboard.append([KeyboardButton("🔙 Back to Main Menu")])
        
        await update.message.reply_text(
            f"✅ অর্ডার তৈরি হয়েছে।\n"
            f"ক্যাটাগরি: {order['sub_cat']}\n"
            f"পরিমাণ: {order['qty']} টি\n"
            f"মোট দাম: {total_price} টাকা\n\n"
            f"💳 অনুগ্রহ করে একটি পেমেন্ট মেথড নির্বাচন করুন:",
            reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        )
        return WAIT_SCREENSHOT
    else:
        keyboard = [
            [KeyboardButton("🔙 Back to Sub Categories")],
            [KeyboardButton("🔙 Back to Main Menu")]
        ]
        await update.message.reply_text(
            f"✅ অর্ডার তৈরি হয়েছে।\n"
            f"ক্যাটাগরি: {order['sub_cat']}\n"
            f"পরিমাণ: {order['qty']} টি\n"
            f"মোট দাম: {total_price} টাকা\n\n"
            f"⚠️ কোনো পেমেন্ট মেথড নেই। অ্যাডমিনের সাথে যোগাযোগ করুন।",
            reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        )
        return WAIT_SCREENSHOT

async def confirm_order(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await check_subscription(update, context):
        return ConversationHandler.END

    text = update.message.text
    
    if text == "🔙 Back to Sub Categories":
        return await back_to_subcategories_handler(update, context)
    if text == "🔙 Back to Main Menu":
        return await start(update, context)
    
    if text == "✅ Confirm Purchase":
        order = context.user_data.get("order", {})
        if not order:
            return ConversationHandler.END
            
        user_id = update.effective_user.id
        main_cat = order['main_cat']
        sub_cat = order['sub_cat']
        total_price = order['price']
        qty = order['qty']
        current_balance = balances.get(user_id, 0)
        
        if current_balance >= total_price:
            balances[user_id] = current_balance - total_price
            
            items = pop_items_from_txt(main_cat, sub_cat, qty)
            if not items:
                balances[user_id] = balances.get(user_id, 0) + total_price 
                await update.message.reply_text("❌ যথেষ্ট আইটেম স্টকে নেই। আপনার ব্যালেন্স ফেরত দেওয়া হয়েছে।", reply_markup=ReplyKeyboardRemove())
                return await start(update, context)
            
            excel_buffer = create_xlsx_file(items, f"{sub_cat}_order.xlsx")
            
            await context.bot.send_document(
                chat_id=user_id,
                document=InputFile(excel_buffer, filename=f"{sub_cat}_order_{int(time.time())}.xlsx"),
                caption=f"✅ আপনার অর্ডার সম্পূর্ণ হয়েছে!\n"
                        f"📦 {sub_cat} - {qty} টি আইটেম\n"
                        f"💰 মোট: {total_price} টাকা\n"
                        f"📄 এক্সেল ফাইলে আপনার অর্ডার সংযুক্ত আছে।"
            )
            
            global total_sales, sales_count_per_category, transaction_log, user_sales
            total_sales += total_price
            sales_count_per_category[sub_cat] = sales_count_per_category.get(sub_cat, 0) + qty
            transaction_log.append(('sale', user_id, total_price, time.time()))
            user_sales[user_id] = user_sales.get(user_id, 0) + total_price

            await update.message.reply_text("✅ ব্যালেন্স ব্যবহার করে আপনার অর্ডার সফল হয়েছে!", reply_markup=ReplyKeyboardRemove())
            context.user_data.clear()
            save_user_data()
            return await start(update, context)
        else:
            await update.message.reply_text("❌ আপনার যথেষ্ট ব্যালেন্স নেই।", reply_markup=ReplyKeyboardRemove())
            return await start(update, context)
    
    return CONFIRM_ORDER

async def user_send_screenshot(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await check_subscription(update, context):
        return ConversationHandler.END

    text = update.message.text
    
    if text == "🔙 Back to Sub Categories":
        return await back_to_subcategories_handler(update, context)
    if text == "🔙 Back to Main Menu":
        return await start(update, context)

    order = context.user_data.get("order", {})
    if not order:
        return ConversationHandler.END
        
    main_cat = order['main_cat']
    sub_cat = order['sub_cat']
    
    # If user selects a payment method
    if text.startswith("💳 "):
        method_name = text.replace("💳 ", "").strip()
        if method_name in payment_methods:
            method_details = payment_methods[method_name].get("details", "কোনো বিস্তারিত তথ্য নেই।")
            context.user_data['selected_payment_method'] = method_name
            
            await update.message.reply_text(
                f"💳 পেমেন্ট মেথড: {method_name}\n"
                f"📋 বিস্তারিত:\n{method_details}\n\n"
                f"📸 পেমেন্টের পর স্ক্রিনশট পাঠান।",
                reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Back to Sub Categories"), KeyboardButton("🔙 Back to Main Menu")]], resize_keyboard=True)
            )
            return WAIT_SCREENSHOT

    if not update.message.photo:
        await update.message.reply_text(
            "❌ অনুগ্রহ করে শুধু একটি ছবি পাঠান অথবা একটি পেমেন্ট মেথড নির্বাচন করুন।",
            reply_markup=ReplyKeyboardMarkup([[KeyboardButton("🔙 Back to Sub Categories"), KeyboardButton("🔙 Back to Main Menu")]], resize_keyboard=True)
        )
        return WAIT_SCREENSHOT
        
    user = update.effective_user
    username = user.username if user.username else 'N/A'
    method_name = context.user_data.get('selected_payment_method', 'অনির্দিষ্ট')
    
    caption = (
        f"🔔 নতুন অর্ডার! 🔔\n"
        f"ব্যবহারকারী: @{username}\n"
        f"ক্যাটাগরি: {order['sub_cat']}\n"
        f"পরিমাণ: {order['qty']} টি\n"
        f"মূল্য: {order['price']} টাকা\n"
        f"পেমেন্ট মেথড: {method_name}\n"
        f"ইউজার আইডি: {user.id}"
    )

    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ Confirm Order", callback_data=f"confirm_manual:{user.id}:{order['main_cat']}:{order['sub_cat']}:{order['qty']}:{order['price']}"),
         InlineKeyboardButton("❌ Cancel Order", callback_data=f"cancel_manual:{user.id}")]
    ])
    
    await context.bot.send_photo(chat_id=ADMIN_ID, photo=update.message.photo[-1].file_id, caption=caption, reply_markup=keyboard)
    
    await update.message.reply_text("✅ স্ক্রিনশট অ্যাডমিনকে পাঠানো হয়েছে। অ্যাডমিন নিশ্চিত করার পর আপনার অর্ডার প্রক্রিয়াকৃত হবে।")
    await start(update, context)
    context.user_data.clear()
    return ConversationHandler.END

# ================ ADMIN ORDER ACTIONS ================
async def admin_order_action(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global total_sales, sales_count_per_category, transaction_log, user_sales, prices
    query = update.callback_query
    await query.answer()
    data = query.data
    
    if update.effective_user.id != ADMIN_ID:
        await query.edit_message_caption("❌ অননুমোদিত।", reply_markup=None)
        return
    
    if data.startswith("confirm_manual:"):
        parts = data.split(":")
        uid = int(parts[1])
        main_cat = parts[2]
        sub_cat = parts[3]
        qty = int(parts[4])
        total_price = float(parts[5])
        
        stock_count = count_items(main_cat, sub_cat)

        if stock_count < qty:
            keyboard = InlineKeyboardMarkup([
                [InlineKeyboardButton("✅ Force Confirm", callback_data=f"force_confirm:{uid}:{main_cat}:{sub_cat}:{qty}:{total_price}"),
                 InlineKeyboardButton("❌ Cancel", callback_data=f"cancel_manual:{uid}")]
            ])
            await query.edit_message_caption(query.message.caption + "\n\n⚠️ স্টকে পর্যাপ্ত আইটেম নেই। আপনি কি নিশ্চিত করতে চান?\n\nযদি নিশ্চিত করেন, আপনাকে ম্যানুয়ালি আইটেমটি পাঠাতে হবে।", reply_markup=keyboard)
            return
            
        items = pop_items_from_txt(main_cat, sub_cat, qty)
        
        if not items:
            await query.edit_message_caption(query.message.caption + "\n\n❌ স্টকে আইটেম নেই।", reply_markup=None)
            return
        
        excel_buffer = create_xlsx_file(items, f"{sub_cat}_order.xlsx")
        
        await context.bot.send_document(
            chat_id=uid,
            document=InputFile(excel_buffer, filename=f"{sub_cat}_order_{int(time.time())}.xlsx"),
            caption=f"✅ আপনার অর্ডার নিশ্চিত করা হয়েছে!\n"
                    f"📦 {sub_cat} - {qty} টি আইটেম\n"
                    f"💰 মোট: {total_price} টাকা\n"
                    f"📄 এক্সেল ফাইলে আপনার অর্ডার সংযুক্ত আছে।"
        )
        
        total_sales += total_price
        sales_count_per_category[sub_cat] = sales_count_per_category.get(sub_cat, 0) + qty
        transaction_log.append(('sale', uid, total_price, time.time()))
        user_sales[uid] = user_sales.get(uid, 0) + total_price
        save_user_data()

        await query.edit_message_caption(query.message.caption + f"\n\n✅ অ্যাডমিন দ্বারা নিশ্চিত। এক্সেল ফাইল ইউজারকে পাঠানো হয়েছে।", reply_markup=None)

    elif data.startswith("force_confirm:"):
        parts = data.split(":")
        uid = int(parts[1])
        main_cat = parts[2]
        sub_cat = parts[3]
        qty = int(parts[4])
        total_price = float(parts[5])

        await context.bot.send_message(
            chat_id=uid, 
            text="✅ আপনার অর্ডার নিশ্চিত করা হয়েছে। অ্যাডমিন শীঘ্রই আপনাকে বিস্তারিত তথ্য পাঠাবেন।\n\n"
                 "⚠️ দয়া করে অপেক্ষা করুন, অ্যাডমিন আপনার আইটেম প্রস্তুত করছে।"
        )
        
        await context.bot.send_message(
            chat_id=ADMIN_ID,
            text=f"⚠️ ফোর্স নিশ্চিত অর্ডার\n"
                 f"ব্যবহারকারী: {uid}\n"
                 f"ক্যাটাগরি: {sub_cat}\n"
                 f"পরিমাণ: {qty} টি\n"
                 f"মোট: {total_price} টাকা\n\n"
                 f"স্টকে পর্যাপ্ত আইটেম নেই। অনুগ্রহ করে ম্যানুয়ালি আইটেম পাঠান।"
        )
        
        total_sales += total_price
        sales_count_per_category[sub_cat] = sales_count_per_category.get(sub_cat, 0) + qty
        transaction_log.append(('sale', uid, total_price, time.time()))
        user_sales[uid] = user_sales.get(uid, 0) + total_price
        save_user_data()

        await query.edit_message_caption(query.message.caption + f"\n\n✅ ফোর্স নিশ্চিত করা হয়েছে। ইউজারকে ম্যানুয়ালি যোগাযোগ করা হবে।", reply_markup=None)

    elif data.startswith("cancel_manual:"):
        _, uid = data.split(":")
        
        await context.bot.send_message(
            chat_id=uid,
            text="❌ দুঃখিত, আপনার অর্ডারটি বাতিল করা হয়েছে।"
        )
        await query.edit_message_caption(query.message.caption + "\n\n❌ অ্যাডমিন দ্বারা বাতিল", reply_markup=None)
        
    else:
        _, uid = data.split(":")
        
        await context.bot.send_message(
            chat_id=uid,
            text="❌ দুঃখিত, আপনার অর্ডারটি বাতিল করা হয়েছে।"
        )
        await query.edit_message_caption(query.message.caption + "\n\n❌ অ্যাডমিন দ্বারা বাতিল", reply_markup=None)

# ================ ADMIN DEPOSIT ACTIONS ================
async def admin_deposit_action(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global total_deposits, transaction_log, user_deposits, balances, pending_deposits, processed_trxids
    query = update.callback_query
    await query.answer()
    data = query.data

    if update.effective_user.id != ADMIN_ID:
        try:
            await query.edit_message_caption("❌ অননুমোদিত।", reply_markup=None)
        except:
            pass
        return
    
    if data.startswith("deposit_confirm:"):
        parts = data.split(":")
        uid = int(parts[1])
        amount = float(parts[2])
        
        balances[uid] = balances.get(uid, 0) + amount
        total_deposits += amount
        transaction_log.append(('deposit', uid, amount, time.time()))
        user_deposits[uid] = user_deposits.get(uid, 0) + amount
        
        # পেন্ডিং লিস্ট থেকে TRXID রিমুভ করুন (যদি থাকে)
        for trxid, deposit in list(pending_deposits.items()):
            if deposit['user_id'] == uid and deposit['amount'] == amount:
                del pending_deposits[trxid]
                processed_trxids.add(trxid)
                break
        
        save_user_data()

        await context.bot.send_message(
            chat_id=uid,
            text=f"✅ আপনার ডিপোজিট সফল হয়েছে। আপনার ব্যালেন্সে {amount} টাকা যোগ হয়েছে।\n"
                 f"নতুন ব্যালেন্স: {balances[uid]} টাকা।"
        )
        
        try:
            await query.edit_message_caption(
                query.message.caption + 
                f"\n\n✅ নিশ্চিত করা হয়েছে। {amount} টাকা ইউজার {uid} এর ব্যালেন্সে যোগ করা হয়েছে। "
                f"বর্তমান মোট ব্যালেন্স: {balances[uid]} টাকা।", 
                reply_markup=None
            )
        except Exception as e:
            logger.error(f"Failed to edit message caption on deposit confirmation: {e}")
            await context.bot.send_message(
                chat_id=ADMIN_ID, 
                text=f"⚠️ ডিপোজিট নিশ্চিত বার্তা আপডেট করতে ব্যর্থ। ইউজার আইডি: {uid}, পরিমাণ: {amount}"
            )
            
    else:
        _, uid = data.split(":")
        uid = int(uid)
        
        # পেন্ডিং লিস্ট থেকে TRXID রিমুভ করুন
        for trxid, deposit in list(pending_deposits.items()):
            if deposit['user_id'] == uid:
                del pending_deposits[trxid]
                break
        
        save_user_data()
        
        await context.bot.send_message(
            chat_id=uid,
            text="❌ দুঃখিত, আপনার ডিপোজিট রিকোয়েস্ট বাতিল করা হয়েছে।"
        )
        
        try:
            await query.edit_message_caption(
                query.message.caption + "\n\n❌ অ্যাডমিন দ্বারা ডিপোজিট বাতিল", 
                reply_markup=None
            )
        except Exception as e:
            logger.error(f"Failed to edit message caption on deposit cancellation: {e}")
            await context.bot.send_message(
                chat_id=ADMIN_ID, 
                text=f"⚠️ ডিপোজিট বাতিল বার্তা আপডেট করতে ব্যর্থ। ইউজার আইডি: {uid}"
            )

# ================ MAIN ================
def main():
    application = Application.builder().token(BOT_TOKEN).build()

    conv = ConversationHandler(
        entry_points=[
            CommandHandler("start", start),
            MessageHandler(filters.TEXT & ~filters.COMMAND, menu_handler)
        ],
        states={
            MAIN_MENU: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Main Menu"), start),
                MessageHandler(filters.TEXT & ~filters.COMMAND, menu_handler)
            ],
            BUY_MENU: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Main Menu"), start),
                MessageHandler(filters.TEXT & ~filters.COMMAND, user_choose_category)
            ],
            BUY_SUB_MENU: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Categories"), back_to_categories_handler),
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Main Menu"), start),
                MessageHandler(filters.TEXT & ~filters.COMMAND, user_choose_subcategory)
            ],
            GET_QUANTITY: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Sub Categories"), back_to_subcategories_handler),
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Main Menu"), start),
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_quantity)
            ],
            CONFIRM_ORDER: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Sub Categories"), back_to_subcategories_handler),
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Main Menu"), start),
                MessageHandler(filters.TEXT & filters.Regex("✅ Confirm Purchase"), confirm_order),
                MessageHandler(filters.TEXT & ~filters.COMMAND, confirm_order)
            ],
            WAIT_SCREENSHOT: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Sub Categories"), back_to_subcategories_handler),
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Main Menu"), start),
                MessageHandler(filters.PHOTO | filters.TEXT, user_send_screenshot),
            ],
            ADMIN_PANEL: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Main Menu"), start),
                MessageHandler(filters.TEXT & ~filters.COMMAND, admin_panel_handler)
            ],
            MANAGE_CATEGORY: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                MessageHandler(filters.TEXT & ~filters.COMMAND, manage_category_handler)
            ],
            MANAGE_SUB_CATEGORY: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                MessageHandler(filters.TEXT & filters.Regex("🔙 Manage Categories"), back_to_manage_main_categories_handler),
                MessageHandler(filters.TEXT & ~filters.COMMAND, manage_sub_category_handler)
            ],
            ADD_MAIN_CAT: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                MessageHandler(filters.TEXT & ~filters.COMMAND, add_main_category)
            ],
            REMOVE_MAIN_CAT: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                MessageHandler(filters.TEXT & ~filters.COMMAND, remove_main_category)
            ],
            ADD_SUB_CAT: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                MessageHandler(filters.TEXT & filters.Regex("🔙 Manage Categories"), back_to_manage_main_categories_handler),
                MessageHandler(filters.TEXT & ~filters.COMMAND, add_sub_category)
            ],
            REMOVE_SUB_CAT: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                MessageHandler(filters.TEXT & filters.Regex("🔙 Manage Categories"), back_to_manage_main_categories_handler),
                MessageHandler(filters.TEXT & ~filters.COMMAND, remove_sub_category)
            ],
            ADD_ITEMS_TXT: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                MessageHandler(filters.TEXT & filters.Regex("🔙 Manage Categories"), back_to_manage_main_categories_handler),
                MessageHandler(filters.TEXT & filters.Regex("✅ Done"), add_items_txt_handler),
                MessageHandler(filters.Document.ALL | filters.TEXT, add_items_txt_handler),
            ],
            EDIT_PAYMENT: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                MessageHandler(filters.TEXT & ~filters.COMMAND, edit_payment_info)
            ],
            EDIT_PRICE_MAIN: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                MessageHandler(filters.TEXT & ~filters.COMMAND, edit_price_main_handler)
            ],
            EDIT_PRICE_SUB: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                MessageHandler(filters.TEXT & ~filters.COMMAND, edit_price_sub_handler)
            ],
            RECEIVE_NEW_PRICE: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_price)
            ],
            DEPOSIT_SELECT_METHOD: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Main Menu"), start),
                MessageHandler(filters.TEXT & ~filters.COMMAND, deposit_enter_amount)
            ],
            DEPOSIT_ENTER_AMOUNT: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Main Menu"), start),
                MessageHandler(filters.TEXT & ~filters.COMMAND, deposit_receive_amount)
            ],
            DEPOSIT_ENTER_TRXID: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Back to Main Menu"), start),
                MessageHandler(filters.TEXT & ~filters.COMMAND, deposit_receive_trxid)
            ],
            DASHBOARD: [
                MessageHandler(filters.TEXT & filters.Regex("🔄 Refresh Dashboard"), handle_dashboard_refresh),
                MessageHandler(filters.TEXT & filters.Regex("🔙 Admin Panel"), back_to_admin_panel_handler)
            ],
            SEND_NOTICE: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                MessageHandler(filters.TEXT & ~filters.COMMAND, send_notice_text)
            ],
            SEARCH_USER_PROFILE: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                MessageHandler(filters.TEXT & ~filters.COMMAND, search_and_show_user_profile)
            ],
            MANAGE_PAYMENT_CATEGORIES: [
                CallbackQueryHandler(toggle_payment_method),
                MessageHandler(filters.TEXT & filters.Regex("🔙 Admin Panel"), back_to_admin_panel_handler)
            ],
            SEARCH_USER_FOR_BALANCE: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                MessageHandler(filters.TEXT & ~filters.COMMAND, search_user_for_balance)
            ],
            BALANCE_EDIT_ACTION: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                MessageHandler(filters.TEXT & ~filters.COMMAND, balance_edit_action_handler)
            ],
            RECEIVE_BALANCE_EDIT_AMOUNT: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_balance_edit_amount)
            ],
            ADMIN_VIEW_PAYMENT_METHODS: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                MessageHandler(filters.TEXT & ~filters.COMMAND, admin_view_payment_methods)
            ],
            ADMIN_ADD_PAYMENT_METHOD: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                MessageHandler(filters.TEXT & ~filters.COMMAND, admin_receive_payment_method_name)
            ],
            ADMIN_RECEIVE_PAYMENT_DETAILS: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                MessageHandler(filters.TEXT & ~filters.COMMAND, admin_receive_payment_method_details)
            ],
            ADMIN_DELETE_PAYMENT_METHOD: [
                MessageHandler(filters.TEXT & filters.Regex("🔙 Admin Panel"), back_to_admin_panel_handler),
                MessageHandler(filters.TEXT & ~filters.COMMAND, admin_delete_payment_method)
            ]
        },
        fallbacks=[CommandHandler("start", start)]
    )

    application.add_handler(conv)
    application.add_handler(CallbackQueryHandler(admin_order_action, pattern="^(confirm:|cancel:|confirm_manual:|cancel_manual:|force_confirm:)"))
    application.add_handler(CallbackQueryHandler(admin_deposit_action, pattern="^(deposit_confirm:|deposit_cancel:)"))
    application.add_handler(CallbackQueryHandler(toggle_payment_method, pattern="^toggle_payment:"))
    
    # SMS অটো ডিপোজিট হ্যান্ডলার - অ্যাডমিনের চ্যাটে আসা সব টেক্সট মেসেজ চেক করবে
    application.add_handler(MessageHandler(
        filters.TEXT & ~filters.COMMAND & filters.Chat(chat_id=ADMIN_ID),
        auto_deposit_from_sms
    ))
    
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND & filters.Regex("🔙 Back to Main Menu"), start))
    
    logger.info("🤖 বট চলছে...")
    application.run_polling()

if __name__ == "__main__":
    main()
