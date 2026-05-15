#!/usr/bin/env python3
import asyncio, io, json, logging, os, re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict

import fitz
import gspread
import PIL.Image
import google.generativeai as genai
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

BOT_TOKEN        = os.getenv("BOT_TOKEN",        "8828924570:AAHGKbAaPVrIVolwkpYbxp8E2emobgGmr0jM")
ALLOWED_USER_ID  = int(os.getenv("ALLOWED_USER_ID", "0"))
GEMINI_API_KEY   = os.getenv("GEMINI_API_KEY",   "AIzaSyAaYdp-IAIETPrW5Yd8DX9YVezfENdjx30")
SPREADSHEET_ID   = "1YlY7Mbmpfq64aLGNLmpk-6kaf2tLX5Tok5JvNiyPyM4"
SHEET_TAB        = os.getenv("SHEET_TAB",        "Database")
CREDS_FILE       = os.getenv("GOOGLE_CREDS",     "credentials.json")

CURRENCIES  = ["USD","EGP","TRY","SAR","Gold","Silver"]
BANKS       = ["Cash","Real Cash-Karim","Real Cash-Wessam","Qnb-Wessam","Qnb-Karim",
               "Kuveyt Turk","Kuveyt Turk Credit Card","Vodafone Cash-Wessam",
               "Vodafone Cash-Karim","Wise-Karim"]
PROJECTS    = ["Work","Family Home","Turkey Home","Transportation","Charity",
               "Qur2An","Karim","Wessam","Mother","Walaa"]
COMPANIES   = ["Karim","Wessam","Charity","Qur2An","Mother","Walaa"]
NAMES       = ["Karim","Wessam","Maria","Father","Manar","Turkey","Old Cairo",
               "New Cairo","Charity","IHH","Wael","Dina","Youtube"]
CATEGORIES  = ["Home","Transportation","Rent Expense","Charity","Household",
               "Utilities","Socialization","Food","Health","Education","Income","Other"]
SUB_CATS    = ["Groceries","Fuel","Taxi","Maintenance","Gebze emlak konut",
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

@dataclass
class Transaction:
    date:str=""; description:str=""; amount_in:float=0.0; amount_out:float=0.0
    currency:str=""; bank:str=""; category:str=""; sub_category:str=""
    name:str=""; project:str=""; company:str=""
    raw_amount:float=0.0; direction:str="OUT"; guessed:dict=field(default_factory=dict)

    def missing_required(self):
        m=[]
        if not self.description: m.append("description")
        if self.raw_amount==0.0: m.append("amount")
        if self.currency not in CURRENCIES: m.append("currency")
        if self.bank not in BANKS: m.append("bank")
        return m

    def to_sheet_row(self):
        return [self.date,self.description,self.amount_in or "",self.amount_out or "",
                self.currency,self.bank,self.category,self.sub_category,self.name,self.project,self.company]

    def summary(self):
        d=(f"IN + {self.amount_in} {self.currency}" if self.direction=="IN"
           else f"OUT - {self.amount_out} {self.currency}")
        lines=["*Transaction Preview*",f"Date: {self.date}",f"Description: {self.description}",
               f"Amount: {d}",f"Bank: {self.bank}",f"Category: {self.category} / {self.sub_category or '-'}"]
        if self.name: lines.append(f"Name: {self.name}")
        if self.project: lines.append(f"Project: {self.project}")
        if self.company: lines.append(f"Company: {self.company}")
        return "\n".join(lines)

pending:Dict[int,Transaction]={}
pending_step:Dict[int,str]={}

class SheetsClient:
    def __init__(self):
        scopes=["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
        creds_json=os.getenv("GOOGLE_CREDS_JSON")
        if creds_json:
            info=json.loads(creds_json)
            creds=Credentials.from_service_account_info(info,scopes=scopes)
        else:
            creds=Credentials.from_service_account_file(CREDS_FILE,scopes=scopes)
        gc=gspread.authorize(creds)
        sh=gc.open_by_key(SPREADSHEET_ID)
        self.ws=sh.worksheet(SHEET_TAB)

    def append(self,row):
        result=self.ws.append_row(row,value_input_option="USER_ENTERED",insert_data_option="INSERT_ROWS")
        try:
            rng=result["updates"]["updatedRange"]
            return int(re.search(r"(\d+):",rng).group(1))
        except: return -1

sheets=SheetsClient()

_SYSTEM=(
    "You are a precise financial data extraction engine.\n"
    "OUTPUT: valid JSON ONLY - no markdown, no explanation.\n"
    f"Schema: {{\"date\":\"DD-MMM-YY or empty\",\"description\":\"concise vendor/item/purpose\","
    f"\"amount\":number,\"direction\":\"OUT or IN\",\"currency\":\"one of {CURRENCIES} or empty\","
    f"\"bank\":\"one of {BANKS} or empty\",\"category\":\"one of {CATEGORIES} or empty\","
    f"\"sub_category\":\"string or empty\",\"name\":\"one of {NAMES} or empty\","
    f"\"project\":\"one of {PROJECTS} or empty\",\"company\":\"one of {COMPANIES} or empty\"}}\n"
    "Rules: direction defaults to OUT unless user says income/received/IN. "
    "Use today for missing dates. Match dropdowns strictly to the allowed lists."
)

class GeminiParser:
    def __init__(self):
        genai.configure(api_key=GEMINI_API_KEY)
        self.model=genai.GenerativeModel("gemini-1.5-flash")

    def _today(self): return datetime.now().strftime("%d-%b-%y")

    def _clean(self,text):
        text=text.strip()
        text=re.sub(r"^```[a-z]*\n?","",text)
        text=re.sub(r"\n?```$","",text)
        return json.loads(text)

    def _build(self,d):
        t=Transaction()
        t.date=d.get("date","") or self._today()
        t.description=(d.get("description","") or "").strip()
        t.raw_amount=float(d.get("amount",0) or 0)
        t.direction=(d.get("direction","OUT") or "OUT").upper()
        t.currency=d.get("currency","") if d.get("currency","") in CURRENCIES else ""
        t.bank=d.get("bank","") if d.get("bank","") in BANKS else ""
        t.category=d.get("category","") if d.get("category","") in CATEGORIES else ""
        t.sub_category=d.get("sub_category","") or ""
        t.name=d.get("name","") if d.get("name","") in NAMES else ""
        t.project=d.get("project","") if d.get("project","") in PROJECTS else ""
        t.company=d.get("company","") if d.get("company","") in COMPANIES else ""
        if t.direction=="IN": t.amount_in=t.raw_amount
        else: t.amount_out=t.raw_amount
        if not t.category:
            for kw,(cat,sub) in HINTS.items():
                if kw in t.description.lower():
                    t.category=cat; t.sub_category=t.sub_category or sub; t.guessed["category"]=True; break
        return t

    def _prompt(self,extra): return _SYSTEM+f"\n\nToday: {self._today()}\n\n"+extra

    def parse_text(self,text):
        return self._build(self._clean(self.model.generate_content(self._prompt(f"User input:\n{text}")).text))

    def parse_image(self,image_bytes):
        img=PIL.Image.open(io.BytesIO(image_bytes))
        return self._build(self._clean(self.model.generate_content([self._prompt("Extract from this receipt image."),img]).text))

    def parse_pdf(self,raw):
        pdf=fitz.open(stream=raw,filetype="pdf")
        text="\n".join(p.get_text() for p in pdf)[:4000]
        return self._build(self._clean(self.model.generate_content(self._prompt(f"PDF text:\n{text}")).text))

parser=GeminiParser()

def _grid(items,prefix,cols=2):
    btns=[InlineKeyboardButton(i,callback_data=f"{prefix}:{i}") for i in items]
    return InlineKeyboardMarkup([btns[i:i+cols] for i in range(0,len(btns),cols)])

def kb_direction(): return InlineKeyboardMarkup([[InlineKeyboardButton("Expense (OUT)",callback_data="dir:OUT"),InlineKeyboardButton("Income (IN)",callback_data="dir:IN")]])
def kb_confirm(): return InlineKeyboardMarkup([[InlineKeyboardButton("Save to sheet",callback_data="confirm:yes"),InlineKeyboardButton("Edit field",callback_data="confirm:edit"),InlineKeyboardButton("Discard",callback_data="confirm:no")]])
def kb_edit(): return InlineKeyboardMarkup([[InlineKeyboardButton("Description",callback_data="edit:description"),InlineKeyboardButton("IN/OUT",callback_data="edit:direction")],[InlineKeyboardButton("Currency",callback_data="edit:currency"),InlineKeyboardButton("Bank",callback_data="edit:bank")],[InlineKeyboardButton("Category",callback_data="edit:category"),InlineKeyboardButton("Sub-category",callback_data="edit:sub_category")],[InlineKeyboardButton("Name",callback_data="edit:name"),InlineKeyboardButton("Project",callback_data="edit:project")],[InlineKeyboardButton("Company",callback_data="edit:company")]])

def _allowed(uid): return ALLOWED_USER_ID==0 or uid==ALLOWED_USER_ID

async def advance(uid,ctx,chat_id):
    t=pending[uid]; m=t.missing_required()
    if "currency" in m:
        pending_step[uid]="currency"
        await ctx.bot.send_message(chat_id,"Which currency?",reply_markup=_grid(CURRENCIES,"currency",3)); return
    if "bank" in m:
        pending_step[uid]="bank"
        await ctx.bot.send_message(chat_id,f"Which bank account? ({t.description})",reply_markup=_grid(BANKS,"bank",2)); return
    if not t.category:
        pending_step[uid]="category"
        await ctx.bot.send_message(chat_id,"Category?",reply_markup=_grid(CATEGORIES,"category",2)); return
    pending_step[uid]="confirm"
    note="\n\n(Auto-guessed: "+(", ".join(t.guessed.keys()))+")" if t.guessed else ""
    await ctx.bot.send_message(chat_id,t.summary()+note,parse_mode="Markdown",reply_markup=kb_confirm())

async def cmd_start(u,c): await u.message.reply_text(f"Finance bot ready. Your ID: {u.effective_user.id}")
async def cmd_whoami(u,c): await u.message.reply_text(f"Your Telegram ID: {u.effective_user.id}")
async def cmd_cancel(u,c):
    uid=u.effective_user.id; pending.pop(uid,None); pending_step.pop(uid,None)
    await u.message.reply_text("Cancelled.")

async def handle_text(u,c):
    uid=u.effective_user.id
    if not _allowed(uid): return
    if pending_step.get(uid)=="edit_description" and uid in pending:
        pending[uid].description=u.message.text.strip(); pending_step[uid]=None
        await advance(uid,c,u.effective_chat.id); return
    await u.message.reply_text("Parsing...")
    try:
        t=parser.parse_text(u.message.text); pending[uid]=t
        await u.message.reply_text(f"Parsed: {t.direction} {t.raw_amount} — {t.description}")
        await advance(uid,c,u.effective_chat.id)
    except Exception as e: log.exception(e); await u.message.reply_text(f"Parse error: {e}")

async def handle_photo(u,c):
    uid=u.effective_user.id
    if not _allowed(uid): return
    await u.message.reply_text("Reading receipt...")
    try:
        f=await c.bot.get_file(u.message.photo[-1].file_id)
        t=parser.parse_image(bytes(await f.download_as_bytearray())); pending[uid]=t
        await u.message.reply_text(f"Found: {t.direction} {t.raw_amount} {t.currency} — {t.description}")
        await advance(uid,c,u.effective_chat.id)
    except Exception as e: log.exception(e); await u.message.reply_text(f"Image error: {e}")

async def handle_document(u,c):
    uid=u.effective_user.id
    if not _allowed(uid): return
    doc=u.message.document; mime=doc.mime_type or ""
    await u.message.reply_text("Processing document...")
    try:
        f=await c.bot.get_file(doc.file_id); raw=bytes(await f.download_as_bytearray())
        t=parser.parse_image(raw) if "image" in mime else parser.parse_pdf(raw) if "pdf" in mime else None
        if not t: await u.message.reply_text("Unsupported type."); return
        pending[uid]=t
        await u.message.reply_text(f"Found: {t.direction} {t.raw_amount} {t.currency} — {t.description}")
        await advance(uid,c,u.effective_chat.id)
    except Exception as e: log.exception(e); await u.message.reply_text(f"Document error: {e}")

async def handle_callback(u,c):
    q=u.callback_query; uid=q.from_user.id
    if not _allowed(uid): return
    await q.answer()
    if uid not in pending: await q.edit_message_text("No active transaction."); return
    prefix,value=q.data.split(":",1); t=pending[uid]
    if prefix=="currency": t.currency=value
    elif prefix=="bank": t.bank=value
    elif prefix=="category":
        t.category=value
        for kw,(cat,sub) in HINTS.items():
            if cat==value and not t.sub_category: t.sub_category=sub; break
    elif prefix=="sub_category": t.sub_category=value
    elif prefix=="name": t.name=value
    elif prefix=="project": t.project=value
    elif prefix=="company": t.company=value
    elif prefix=="dir":
        t.direction=value
        if value=="IN": t.amount_in=t.raw_amount; t.amount_out=0.0
        else: t.amount_out=t.raw_amount; t.amount_in=0.0
    elif prefix=="confirm":
        if value=="yes":
            try:
                n=sheets.append(t.to_sheet_row())
                await q.edit_message_text(f"Saved to row {n}!\n\n{t.summary()}",parse_mode="Markdown")
            except Exception as e: await q.edit_message_text(f"Sheet error: {e}")
            finally: pending.pop(uid,None); pending_step.pop(uid,None)
        elif value=="no":
            pending.pop(uid,None); pending_step.pop(uid,None)
            await q.edit_message_text("Discarded.")
        elif value=="edit":
            await q.edit_message_text("Which field to edit?",reply_markup=kb_edit())
        return
    elif prefix=="edit":
        if value=="description": pending_step[uid]="edit_description"; await q.edit_message_text("Type new description:")
        elif value=="direction": await q.edit_message_text("IN or OUT?",reply_markup=kb_direction())
        elif value=="currency": await q.edit_message_text("Pick currency:",reply_markup=_grid(CURRENCIES,"currency",3))
        elif value=="bank": await q.edit_message_text("Pick bank:",reply_markup=_grid(BANKS,"bank",2))
        elif value=="category": await q.edit_message_text("Pick category:",reply_markup=_grid(CATEGORIES,"category",2))
        elif value=="sub_category": await q.edit_message_text("Pick sub-category:",reply_markup=_grid(SUB_CATS,"sub_category",2))
        elif value=="name": await q.edit_message_text("Pick name:",reply_markup=_grid(NAMES,"name",2))
        elif value=="project": await q.edit_message_text("Pick project:",reply_markup=_grid(PROJECTS,"project",2))
        elif value=="company": await q.edit_message_text("Pick company:",reply_markup=_grid(COMPANIES,"company",2))
        return
    await advance(uid,c,q.message.chat_id)

def main():
    app=Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start",cmd_start))
    app.add_handler(CommandHandler("whoami",cmd_whoami))
    app.add_handler(CommandHandler("cancel",cmd_cancel))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND,handle_text))
    app.add_handler(MessageHandler(filters.PHOTO,handle_photo))
    app.add_handler(MessageHandler(filters.Document.ALL,handle_document))
    app.add_handler(CallbackQueryHandler(handle_callback))
    log.info("Bot starting...")
    app.run_polling(drop_pending_updates=True)

if __name__=="__main__":
    main()
