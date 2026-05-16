#!/usr/bin/env python3
import base64, io, json, logging, os, re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict

import fitz
import gspread
import PIL.Image
from groq import Groq
from google.oauth2.service_account import Credentials
from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import (Application, CallbackQueryHandler, CommandHandler,
                           ContextTypes, MessageHandler, filters)
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

logging.basicConfig(format="%(asctime)s | %(levelname)s | %(message)s", level=logging.INFO)
log = logging.getLogger(__name__)

BOT_TOKEN       = os.getenv("BOT_TOKEN", "")
ALLOWED_USER_IDS = set(int(x.strip()) for x in os.getenv("ALLOWED_USER_IDS", "0").split(",") if x.strip().isdigit())
GROQ_API_KEY    = os.getenv("GROQ_API_KEY", "")
SPREADSHEET_ID  = "1YlY7Mbmpfq64aLGNLmpk-6kaf2tLX5Tok5JvNiyPyM4"
SHEET_TAB       = os.getenv("SHEET_TAB", "Database")
CREDS_FILE      = os.getenv("GOOGLE_CREDS", "credentials.json")

CURRENCIES = ["USD","EGP","TRY","SAR","Gold","Silver"]
BANKS      = ["Cash","Real Cash-Karim","Real Cash-Wessam","Qnb-Wessam","Qnb-Karim",
              "Kuveyt Turk","Kuveyt Turk Credit Card","Vodafone Cash-Wessam",
              "Vodafone Cash-Karim","Wise-Karim"]
PROJECTS   = ["Work","Family Home","Turkey Home","Transportation","Charity",
              "Qur2An","Karim","Wessam","Mother","Walaa"]
COMPANIES  = ["Karim","Wessam","Charity","Qur2An","Mother","Walaa"]
NAMES      = ["Karim","Wessam","Maria","Father","Manar","Turkey","Old Cairo",
              "New Cairo","Charity","IHH","Wael","Dina","Youtube"]
CATEGORIES = ["Home","Transportation","Rent Expense","Charity","Household",
              "Utilities","Socialization","Food","Health","Education","Income","Other"]
SUB_CATS   = ["Groceries","Fuel","Taxi","Maintenance","Gebze emlak konut",
              "General Sadakah","Subscription","Internet","Mobile","Going Out",
              "Salary","Freelance","Transfer","Other"]

HINTS = {
    "bim":("Home","Groceries"),"migros":("Home","Groceries"),"tatbak":("Home","Groceries"),
    "groceries":("Home","Groceries"),"diesel":("Transportation","Fuel"),"fuel":("Transportation","Fuel"),
    "petrol":("Transportation","Fuel"),"uber":("Transportation","Taxi"),"taxi":("Transportation","Taxi"),
    "peugeot":("Transportation","Maintenance"),"206":("Transportation","Maintenance"),
    "car":("Transportation","Maintenance"),"rent":("Rent Expense","Gebze emlak konut"),
    "emlak":("Rent Expense","Gebze emlak konut"),"sadakah":("Charity","General Sadakah"),
    "charity":("Charity","General Sadakah"),"youtube":("Household","Subscription"),
    "netflix":("Household","Subscription"),"google":("Household","Subscription"),
    "internet":("Utilities","Internet"),"vodafone":("Utilities","Mobile"),
    "mobile":("Utilities","Mobile"),"coffee":("Socialization","Going Out"),
    "sweets":("Socialization","Going Out"),
}

# ── Data model ────────────────────────────────────────────────
@dataclass
class Transaction:
    date:str=""; description:str=""; amount_in:float=0.0; amount_out:float=0.0
    currency:str=""; bank:str=""; category:str=""; sub_category:str=""
    name:str=""; project:str=""; company:str=""
    raw_amount:float=0.0; direction:str="OUT"; guessed:dict=field(default_factory=dict)
    is_transfer:bool=False; transfer_to_bank:str=""

    def missing_required(self):
        m=[]
        if not self.description:            m.append("description")
        if self.raw_amount==0.0:            m.append("amount")
        if self.currency not in CURRENCIES: m.append("currency")
        if self.bank not in BANKS:          m.append("bank")
        if self.is_transfer and self.transfer_to_bank not in BANKS:
                                            m.append("transfer_to_bank")
        if self.company not in COMPANIES:   m.append("company")
        return m

    def to_sheet_rows(self):
        """Returns list of rows — 2 rows for transfers, 1 for normal."""
        if self.is_transfer:
            base = [self.date, self.description, "", "", self.currency,
                    "", "Transfer", "Transfer", "", "", self.company]
            out_row = base.copy(); out_row[3] = self.raw_amount; out_row[5] = self.bank
            in_row  = base.copy(); in_row[2]  = self.raw_amount; in_row[5] = self.transfer_to_bank
            return [out_row, in_row]
        return [[self.date, self.description,
                 self.amount_in or "", self.amount_out or "",
                 self.currency, self.bank, self.category, self.sub_category,
                 self.name, self.project, self.company]]

    def summary(self):
        if self.is_transfer:
            lines = ["*Transfer Preview*",
                     f"Date: {self.date}",
                     f"Description: {self.description}",
                     f"Amount: {self.raw_amount} {self.currency}",
                     f"From Bank: {self.bank}",
                     f"To Bank: {self.transfer_to_bank}",
                     f"Company: {self.company}",
                     "_Creates 2 rows: OUT + IN_"]
        else:
            d = (f"IN + {self.amount_in} {self.currency}" if self.direction=="IN"
                 else f"OUT - {self.amount_out} {self.currency}")
            lines = ["*Transaction Preview*",
                     f"Date: {self.date}",
                     f"Description: {self.description}",
                     f"Amount: {d}",
                     f"Bank: {self.bank}",
                     f"Company: {self.company}"]
            if self.category: lines.append(f"Category: {self.category} / {self.sub_category or '-'}")
        return "\n".join(lines)

pending: Dict[int, Transaction] = {}
pending_step: Dict[int, str] = {}

# ── Google Sheets — lazy singleton ────────────────────────────
_ws = None
def get_ws():
    global _ws
    if _ws is None:
        scopes = ["https://www.googleapis.com/auth/spreadsheets",
                  "https://www.googleapis.com/auth/drive"]
        creds_json = os.getenv("GOOGLE_CREDS_JSON")
        if creds_json:
            creds = Credentials.from_service_account_info(json.loads(creds_json), scopes=scopes)
        else:
            creds = Credentials.from_service_account_file(CREDS_FILE, scopes=scopes)
        gc  = gspread.authorize(creds)
        sh  = gc.open_by_key(SPREADSHEET_ID)
        _ws = sh.worksheet(SHEET_TAB)
        log.info("Sheets connected — tab: %s", SHEET_TAB)
    return _ws

def save_transaction(t: Transaction):
    """Write 1 or 2 rows and return row number(s)."""
    ws   = get_ws()
    rows = t.to_sheet_rows()
    nums = []
    for row in rows:
        result = ws.append_row(row, value_input_option="USER_ENTERED",
                               insert_data_option="INSERT_ROWS")
        try:
            nums.append(int(re.search(r"(\d+):", result["updates"]["updatedRange"]).group(1)))
        except Exception:
            nums.append(-1)
    return nums

# ── Groq LLM parser ───────────────────────────────────────────
_SYSTEM = (
    "You are a precise financial data extraction engine.\n"
    "OUTPUT: valid JSON ONLY — no markdown fences, no explanation.\n"
    'Schema: {"date":"DD-MMM-YY or empty","description":"concise vendor/item/purpose",'
    '"amount":number,"direction":"OUT|IN|TRANSFER",'
    '"currency":"USD|EGP|TRY|SAR|Gold|Silver or empty",'
    '"bank":"Cash|Real Cash-Karim|Real Cash-Wessam|Qnb-Wessam|Qnb-Karim|Kuveyt Turk|Kuveyt Turk Credit Card|Vodafone Cash-Wessam|Vodafone Cash-Karim|Wise-Karim or empty",'
    '"transfer_to_bank":"destination bank for transfers, same allowed values, or empty",'
    '"category":"Home|Transportation|Rent Expense|Charity|Household|Utilities|Socialization|Food|Health|Education|Income|Other or empty",'
    '"sub_category":"string or empty",'
    '"name":"Karim|Wessam|Maria|Father|Manar|Turkey|Old Cairo|New Cairo|Charity|IHH|Wael|Dina|Youtube or empty",'
    '"project":"Work|Family Home|Turkey Home|Transportation|Charity|Qur2An|Karim|Wessam|Mother|Walaa or empty",'
    '"company":"Karim|Wessam|Charity|Qur2An|Mother|Walaa or empty"}\n'
    "Rules:\n"
    "- direction=TRANSFER if money moves between accounts/banks\n"
    "- direction=IN for income/received/salary\n"
    "- direction=OUT for all expenses (default)\n"
    "- Use today's date if none found\n"
    "- Match all dropdown fields strictly to allowed values; empty string if unsure\n"
    "- For receipts/invoices: OUT unless stated otherwise\n"
    "- Extract ALL visible amounts, dates, vendor names from receipts"
)

class GroqParser:
    def __init__(self):
        self.client       = Groq(api_key=GROQ_API_KEY)
        self.text_model   = "llama-3.3-70b-versatile"
        self.vision_model = "llama-3.2-90b-vision-preview"

    def _today(self): return datetime.now().strftime("%d-%b-%y")

    def _build(self, raw):
        raw = raw.strip()
        raw = re.sub(r"^```[a-z]*\n?", "", raw)
        raw = re.sub(r"\n?```$", "", raw)
        d   = json.loads(raw)
        t   = Transaction()
        t.date            = d.get("date","")            or self._today()
        t.description     = (d.get("description","")    or "").strip()
        t.raw_amount      = float(d.get("amount",0)     or 0)
        direction         = (d.get("direction","OUT")   or "OUT").upper()
        t.is_transfer     = direction == "TRANSFER"
        t.direction       = direction if direction in ("IN","OUT") else "OUT"
        t.currency        = d.get("currency","")        if d.get("currency","")        in CURRENCIES else ""
        t.bank            = d.get("bank","")            if d.get("bank","")            in BANKS      else ""
        t.transfer_to_bank= d.get("transfer_to_bank","")if d.get("transfer_to_bank","")in BANKS      else ""
        t.category        = d.get("category","")        if d.get("category","")        in CATEGORIES else ""
        t.sub_category    = d.get("sub_category","")    or ""
        t.name            = d.get("name","")            if d.get("name","")            in NAMES      else ""
        t.project         = d.get("project","")         if d.get("project","")         in PROJECTS   else ""
        t.company         = d.get("company","")         if d.get("company","")         in COMPANIES  else ""
        if not t.is_transfer:
            if t.direction=="IN": t.amount_in  = t.raw_amount
            else:                 t.amount_out = t.raw_amount
        if not t.category:
            for kw,(cat,sub) in HINTS.items():
                if kw in t.description.lower():
                    t.category=cat; t.sub_category=t.sub_category or sub
                    t.guessed["category"]=True; break
        return t

    def _chat(self, messages):
        r = self.client.chat.completions.create(
            model=self.text_model, messages=messages,
            response_format={"type":"json_object"}, temperature=0.1, max_tokens=512)
        return r.choices[0].message.content

    def _vision_call(self, messages):
        r = self.client.chat.completions.create(
            model=self.vision_model, messages=messages,
            temperature=0.1, max_tokens=768)
        return r.choices[0].message.content

    def parse_text(self, text):
        return self._build(self._chat([
            {"role":"system","content":_SYSTEM},
            {"role":"user","content":f"Today: {self._today()}\n\n{text}"}]))

    def parse_image(self, image_bytes):
        # Ensure JPEG, resize to 1568px max (Groq vision optimal)
        img = PIL.Image.open(io.BytesIO(image_bytes)).convert("RGB")
        img.thumbnail((1568, 1568), PIL.Image.LANCZOS)
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=90)
        b64 = base64.b64encode(buf.getvalue()).decode()
        return self._build(self._vision_call([
            {"role":"system","content":_SYSTEM},
            {"role":"user","content":[
                {"type":"text","text":(
                    f"Today: {self._today()}\n"
                    "This is a receipt or invoice image. "
                    "Extract ALL transaction data visible. Return JSON only.")},
                {"type":"image_url","image_url":{"url":f"data:image/jpeg;base64,{b64}","detail":"high"}}
            ]}]))

    def parse_pdf(self, raw_bytes):
        pdf  = fitz.open(stream=raw_bytes, filetype="pdf")
        text = "\n".join(p.get_text() for p in pdf)[:5000]
        return self._build(self._chat([
            {"role":"system","content":_SYSTEM},
            {"role":"user","content":f"Today: {self._today()}\nPDF content:\n{text}"}]))

parser = GroqParser()

# ── Keyboards ─────────────────────────────────────────────────
def _grid(items, prefix, cols=2):
    btns = [InlineKeyboardButton(i, callback_data=f"{prefix}:{i}") for i in items]
    return InlineKeyboardMarkup([btns[i:i+cols] for i in range(0,len(btns),cols)])

def kb_direction(): return InlineKeyboardMarkup([[
    InlineKeyboardButton("Expense (OUT)",  callback_data="dir:OUT"),
    InlineKeyboardButton("Income (IN)",    callback_data="dir:IN"),
    InlineKeyboardButton("Transfer",       callback_data="dir:TRANSFER")]])

def kb_confirm(): return InlineKeyboardMarkup([[
    InlineKeyboardButton("Save to sheet",  callback_data="confirm:yes"),
    InlineKeyboardButton("Edit field",     callback_data="confirm:edit"),
    InlineKeyboardButton("Discard",        callback_data="confirm:no")]])

def kb_edit(): return InlineKeyboardMarkup([
    [InlineKeyboardButton("Description",   callback_data="edit:description"),
     InlineKeyboardButton("IN/OUT/Transfer",callback_data="edit:direction")],
    [InlineKeyboardButton("Currency",      callback_data="edit:currency"),
     InlineKeyboardButton("Bank",          callback_data="edit:bank")],
    [InlineKeyboardButton("Company",       callback_data="edit:company"),
     InlineKeyboardButton("Category",      callback_data="edit:category")]])

# ── Conversation flow ─────────────────────────────────────────
def _allowed(uid): return 0 in ALLOWED_USER_IDS or uid in ALLOWED_USER_IDS

async def advance(uid, ctx, chat_id):
    t = pending[uid]
    m = t.missing_required()

    if "currency" in m:
        pending_step[uid] = "currency"
        await ctx.bot.send_message(chat_id, "💱 Which currency?",
                                   reply_markup=_grid(CURRENCIES,"currency",3)); return

    if "bank" in m:
        pending_step[uid] = "bank"
        label = "🏦 Which bank is the *source* (FROM)?" if t.is_transfer else "🏦 Which bank account?"
        await ctx.bot.send_message(chat_id, label, parse_mode="Markdown",
                                   reply_markup=_grid(BANKS,"bank",2)); return

    if "transfer_to_bank" in m:
        pending_step[uid] = "transfer_to_bank"
        await ctx.bot.send_message(chat_id, "🏦 Which bank is the *destination* (TO)?",
                                   parse_mode="Markdown",
                                   reply_markup=_grid(BANKS,"transfer_to_bank",2)); return

    if "company" in m:
        pending_step[uid] = "company"
        await ctx.bot.send_message(chat_id, "🏢 Which company?",
                                   reply_markup=_grid(COMPANIES,"company",3)); return

    # All required fields present — show confirm
    pending_step[uid] = "confirm"
    note = ("\n\n_(Auto-guessed: " + ", ".join(t.guessed.keys()) + ")_") if t.guessed else ""
    await ctx.bot.send_message(chat_id, t.summary()+note,
                               parse_mode="Markdown", reply_markup=kb_confirm())

# ── Command handlers ──────────────────────────────────────────
async def cmd_start(u,c):
    await u.message.reply_text(
        f"Finance bot ready.\nYour Telegram ID: `{u.effective_user.id}`\n\n"
        "Send any expense, income, transfer, receipt image, or PDF.",
        parse_mode="Markdown")

async def cmd_whoami(u,c):
    await u.message.reply_text(f"Your Telegram ID: `{u.effective_user.id}`", parse_mode="Markdown")

async def cmd_cancel(u,c):
    uid=u.effective_user.id; pending.pop(uid,None); pending_step.pop(uid,None)
    await u.message.reply_text("Cancelled.")

async def cmd_ping(u,c):
    await u.message.reply_text("Pong! Bot is running.")

# ── Message handlers ──────────────────────────────────────────
async def handle_text(u,c):
    uid=u.effective_user.id
    if not _allowed(uid): return
    if pending_step.get(uid)=="edit_description" and uid in pending:
        pending[uid].description=u.message.text.strip(); pending_step[uid]=None
        await advance(uid,c,u.effective_chat.id); return
    await u.message.reply_text("⏳ Parsing...")
    try:
        t=parser.parse_text(u.message.text); pending[uid]=t
        kind = "Transfer" if t.is_transfer else t.direction
        await u.message.reply_text(
            f"🔍 *{kind}* `{t.raw_amount}` — _{t.description}_", parse_mode="Markdown")
        await advance(uid,c,u.effective_chat.id)
    except Exception as e:
        log.exception(e); await u.message.reply_text(f"Parse error: {e}")

async def handle_photo(u,c):
    uid=u.effective_user.id
    if not _allowed(uid): return
    await u.message.reply_text("📷 Reading receipt...")
    try:
        f=await c.bot.get_file(u.message.photo[-1].file_id)
        t=parser.parse_image(bytes(await f.download_as_bytearray())); pending[uid]=t
        await u.message.reply_text(
            f"🔍 *{t.direction}* `{t.raw_amount}` `{t.currency}` — _{t.description}_",
            parse_mode="Markdown")
        await advance(uid,c,u.effective_chat.id)
    except Exception as e:
        log.exception(e); await u.message.reply_text(f"Image error: {e}")

async def handle_document(u,c):
    uid=u.effective_user.id
    if not _allowed(uid): return
    doc=u.message.document; mime=doc.mime_type or ""
    await u.message.reply_text("📄 Processing document...")
    try:
        f=await c.bot.get_file(doc.file_id); raw=bytes(await f.download_as_bytearray())
        t=(parser.parse_image(raw) if "image" in mime
           else parser.parse_pdf(raw) if "pdf" in mime else None)
        if not t: await u.message.reply_text("Unsupported type. Send PDF or image."); return
        pending[uid]=t
        await u.message.reply_text(
            f"🔍 *{t.direction}* `{t.raw_amount}` `{t.currency}` — _{t.description}_",
            parse_mode="Markdown")
        await advance(uid,c,u.effective_chat.id)
    except Exception as e:
        log.exception(e); await u.message.reply_text(f"Document error: {e}")

async def handle_callback(u,c):
    q=u.callback_query; uid=q.from_user.id
    if not _allowed(uid): return
    await q.answer()
    if uid not in pending:
        await q.edit_message_text("No active transaction. Send a new one."); return
    prefix,value=q.data.split(":",1); t=pending[uid]

    if   prefix=="currency":          t.currency=value
    elif prefix=="bank":              t.bank=value
    elif prefix=="transfer_to_bank":  t.transfer_to_bank=value
    elif prefix=="company":           t.company=value
    elif prefix=="category":
        t.category=value
        for kw,(cat,sub) in HINTS.items():
            if cat==value and not t.sub_category: t.sub_category=sub; break
    elif prefix=="sub_category":      t.sub_category=value
    elif prefix=="name":              t.name=value
    elif prefix=="project":           t.project=value
    elif prefix=="dir":
        if value=="TRANSFER":
            t.is_transfer=True; t.direction="OUT"
        else:
            t.is_transfer=False; t.direction=value
            if value=="IN":  t.amount_in=t.raw_amount;  t.amount_out=0.0
            else:            t.amount_out=t.raw_amount; t.amount_in=0.0

    elif prefix=="confirm":
        if value=="yes":
            try:
                nums=save_transaction(t)
                rows_txt = " & ".join(str(n) for n in nums)
                label = "Transfer" if t.is_transfer else "Transaction"
                await q.edit_message_text(
                    f"✅ *{label} saved to row(s) {rows_txt}!*\n\n{t.summary()}",
                    parse_mode="Markdown")
            except Exception as e:
                log.exception(e); await q.edit_message_text(f"Sheet error: {e}")
            finally: pending.pop(uid,None); pending_step.pop(uid,None)
        elif value=="no":
            pending.pop(uid,None); pending_step.pop(uid,None)
            await q.edit_message_text("Discarded.")
        elif value=="edit":
            await q.edit_message_text("Which field to edit?", reply_markup=kb_edit())
        return

    elif prefix=="edit":
        if   value=="description":  pending_step[uid]="edit_description"; await q.edit_message_text("Type new description:")
        elif value=="direction":    await q.edit_message_text("Type?", reply_markup=kb_direction())
        elif value=="currency":     await q.edit_message_text("Currency:", reply_markup=_grid(CURRENCIES,"currency",3))
        elif value=="bank":         await q.edit_message_text("Bank:", reply_markup=_grid(BANKS,"bank",2))
        elif value=="company":      await q.edit_message_text("Company:", reply_markup=_grid(COMPANIES,"company",3))
        elif value=="category":     await q.edit_message_text("Category:", reply_markup=_grid(CATEGORIES,"category",2))
        return

    await advance(uid,c,q.message.chat_id)

# ── Entry point ───────────────────────────────────────────────
def main():
    log.info("Bot starting — ALLOWED_USER_ID=%s SHEET_TAB=%s", ALLOWED_USER_ID, SHEET_TAB)
    app=Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start",   cmd_start))
    app.add_handler(CommandHandler("whoami",  cmd_whoami))
    app.add_handler(CommandHandler("cancel",  cmd_cancel))
    app.add_handler(CommandHandler("ping",    cmd_ping))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    app.add_handler(MessageHandler(filters.PHOTO,        handle_photo))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(CallbackQueryHandler(handle_callback))
    app.run_polling(drop_pending_updates=True)

if __name__=="__main__":
    main()
