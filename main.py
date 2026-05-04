# main_pipeline.py
# Complete pipeline: YouTube → Filter → Gemini → VADER → Google Drive Excel
# Works in Google Colab AND GitHub Actions automatically

import os
import io
import re
import json
import time
import random
import hashlib
import warnings
import numpy as np
import pandas as pd
warnings.filterwarnings("ignore")

from googleapiclient.discovery import build
from google import genai
from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer

# ── KEYS ────────────────────────────────────────────────
# In Colab: paste your keys directly
# In GitHub Actions: reads from GitHub Secrets automatically
YOUTUBE_API_KEY   = os.getenv("YOUTUBE_API_KEY",   "PASTE_YOUR_YOUTUBE_KEY")
GEMINI_API_KEY    = os.getenv("GEMINI_API_KEY",    "PASTE_YOUR_GEMINI_KEY")
GDRIVE_CREDENTIALS = os.getenv("GDRIVE_CREDENTIALS", "")  # JSON string from GitHub Secret

LOCAL_FILE = "FINAL_STRUCTURED_COMPLAINTS.xlsx"

# ── SETUP APIs ───────────────────────────────────────────
youtube  = build("youtube", "v3", developerKey=YOUTUBE_API_KEY)
gemini_client = genai.Client(api_key=GEMINI_API_KEY)
analyzer = SentimentIntensityAnalyzer()

# ── AUTOMOTIVE VADER BOOSTERS ────────────────────────────
car_lexicon = {
    "breakdown":-3.0,"broke":-2.5,"failed":-2.5,"failure":-2.5,
    "defect":-2.5,"fault":-2.0,"malfunction":-2.5,"problem":-2.0,
    "issue":-1.8,"drain":-2.0,"stuck":-2.5,"rattle":-2.0,
    "recall":-3.0,"dangerous":-3.0,"frustrated":-2.5,"terrible":-3.0,
    "worst":-3.0,"avoid":-2.0,"degradation":-2.0,"stalled":-2.5,
    "disappointed":-2.5,"pathetic":-2.5,"useless":-2.5,
    "smooth":2.0,"excellent":2.5,"amazing":2.5,"reliable":2.0,
    "recommend":2.0,"efficient":1.8,"satisfied":2.0,"best":2.5,
    "love":2.0,"fantastic":2.5,"comfortable":1.5,"worth":1.5,
}
analyzer.lexicon.update(car_lexicon)

# ── KEYWORD GROUPS ───────────────────────────────────────
KEYWORDS = {
    "EV System": [
        "EV system","electric vehicle system","EV battery","battery system",
        "battery drain","battery degradation","battery management",
        "battery capacity","battery warning","battery dead","traction battery",
        "battery not charging","charge not working","EV not charging",
        "charging stopped","charging failed","charging port","charging system",
        "fast charging","DC charging","AC charging","onboard charger",
        "regenerative braking","electric motor failure","electric drivetrain",
        "electric range","range anxiety","range dropped","electric system fault",
        "EV failure","EV mode","state of charge","vehicle to load",
        "electric vehicle","BMS","OBC","SOC","V2L","e-motor",
        "Nexon EV","Tata EV","Punch EV","Tigor EV","MG ZS EV",
        "Windsor EV","Comet EV","Creta EV","Ioniq 5","Ioniq 6",
        "Mahindra EV","XEV 9e","BE 6e","XUV400 EV","BYD Atto","BYD Seal",
        "km on full charge","range on full charge","full charge range",
        "ev charger problem","charging station problem","ev breakdown",
        "charge nahi ho raha","battery khatam","range kam ho gayi",
    ],
    "Series Hybrid EV": [
        "series hybrid","series hybrid EV","hybrid electric vehicle",
        "parallel hybrid","mild hybrid","strong hybrid","full hybrid",
        "self charging hybrid","plug-in hybrid","range extender",
        "hybrid mode","hybrid system","hybrid drivetrain",
        "hybrid system fault","hybrid battery issue","hybrid not working",
        "hybrid fuel economy","hybrid mileage","electric only mode",
        "e-power","i-MMD","e-CVT","PHEV","MHEV","EREV",
        "Grand Vitara hybrid","Hyryder hybrid","Innova Hycross",
        "City hybrid","Honda hybrid","Invicto hybrid","Toyota hybrid",
        "Maruti hybrid","Camry hybrid","Tucson hybrid",
        "hybrid lag","atkinson cycle","e-CVT problem","hybrid jerky",
        "hybrid kitna deta hai","hybrid mode nahi aata",
    ],
    "Power Back Door": [
        "power back door","power tailgate","electric tailgate","auto tailgate",
        "hands free tailgate","power liftgate","electric boot","automatic boot",
        "smart tailgate","kick sensor tailgate","foot sensor boot","power trunk",
        "tailgate not opening","tailgate not closing","tailgate stuck",
        "tailgate sensor","tailgate malfunction","tailgate rattle",
        "tailgate noise","boot not closing","boot door problem",
        "rear door motor","liftgate motor","power rear door","hands free boot",
        "motorized tailgate","electric dicky","power dicky","boot sensor",
        "foot gesture boot","auto close boot","dicky problem",
        "boot khul nahi raha","dicky band nahi hoti",
    ]
}

OWNERS = {
    "EV System":        ("Nandini Sharma", "Nandinisharma8862@gmail.com"),
    "Series Hybrid EV": ("Nandini Singh",   "065038@fsm.ac.in"),
    "Power Back Door":  ("Tejas",  "Personaltej2909@gmail.com"),
    "Other":            ("General", "general@maruti.com"),
}

REQUIRES_CAR_CONTEXT = {"EV","BMS","OBC","SOC","V2L","PHEV","MHEV","EREV"}
CAR_CONTEXT_WORDS = [
    "car","vehicle","suv","sedan","maruti","tata","hyundai","kia","honda",
    "toyota","mg","mahindra","nexon","creta","brezza","hyryder","innova",
    "city","service","dealer","km","kmpl","mileage","range","charging",
    "battery","engine","drive","owner","review","problem","issue","fault",
]

SEARCH_QUERIES = [
    "Nexon EV battery problem India 2024",
    "Nexon EV charging issue India",
    "Tata EV breakdown India owner",
    "Punch EV range problem India",
    "MG Windsor EV issue India",
    "Creta EV battery review India",
    "EV charging problem India 2024",
    "electric vehicle battery drain India",
    "BMS error Nexon EV India",
    "Mahindra XUV400 EV problem India",
    "Grand Vitara hybrid problem India",
    "Hyryder hybrid mileage problem India",
    "Innova Hycross hybrid problem India",
    "Honda City hybrid review India",
    "strong hybrid mileage India",
    "power tailgate problem India SUV",
    "electric tailgate not working India",
    "tailgate sensor issue India car",
    "electric dicky problem India car",
    "SUV tailgate malfunction India",
]

LOG_FILE = "yt_last_run.json"

# ════════════════════════════════════════════════════════
# GOOGLE DRIVE FUNCTIONS
# ════════════════════════════════════════════════════════

def get_drive_service():
    """Builds Google Drive service from JSON credentials"""
    from google.oauth2 import service_account
    from googleapiclient.discovery import build as gdrive_build

    # In GitHub Actions — reads from GDRIVE_CREDENTIALS secret
    # In Colab — reads from uploaded JSON file
    if GDRIVE_CREDENTIALS:
        creds_info = json.loads(GDRIVE_CREDENTIALS)
    else:
        # Colab: load from uploaded JSON file
        json_files = [f for f in os.listdir(".")
                      if f.endswith(".json") and "pipeline" not in f
                      and "last_run" not in f]
        if not json_files:
            raise FileNotFoundError(
                "No credentials JSON found. "
                "Upload your service account JSON file to Colab."
            )
        with open(json_files[0]) as f:
            creds_info = json.load(f)
        print(f"Using credentials: {json_files[0]}")

    creds = service_account.Credentials.from_service_account_info(
        creds_info,
        scopes=["https://www.googleapis.com/auth/drive"]
    )
    print(f"Service account: {creds_info.get('client_email','')}")
    return gdrive_build("drive", "v3", credentials=creds,
                        cache_discovery=False)


def find_file_id(service, filename=LOCAL_FILE):
    results = service.files().list(
        q=f"name='{filename}' and trashed=false",
        fields="files(id, name, modifiedTime)",
        pageSize=10
    ).execute()

    files = results.get("files", [])

    if not files:
        print(f"❌ File '{filename}' NOT found in Drive")
        return None

    # pick latest modified file (important if duplicates exist)
    files = sorted(files, key=lambda x: x.get("modifiedTime", ""), reverse=True)
    f = files[0]

    print(f"✅ Found file: {f['name']} (ID: {f['id']})")
    return f["id"]


def download_from_drive(service):
    """Downloads Excel from Google Drive"""
    from googleapiclient.http import MediaIoBaseDownload
    print("Downloading Excel from Google Drive...")
    file_id = find_file_id(service)
    if not file_id:
        print("No existing file — will create new one")
        return False
    request    = service.files().get_media(fileId=file_id)
    buffer     = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    with open(LOCAL_FILE, "wb") as f:
        f.write(buffer.getvalue())
    size = os.path.getsize(LOCAL_FILE)
    print(f"Downloaded: {LOCAL_FILE} ({size:,} bytes)")
    return True


def upload_to_drive(service):
    """Uploads updated Excel back to Google Drive"""
    from googleapiclient.http import MediaFileUpload
    print("Uploading updated Excel to Google Drive...")
    file_id = find_file_id(service)
    media   = MediaFileUpload(
        LOCAL_FILE,
        mimetype=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        ),
        resumable=True
    )
    if file_id:
        service.files().update(
            fileId=file_id, media_body=media
        ).execute()
        print("Updated existing file on Drive")
    else:
        metadata = {"name": LOCAL_FILE}
        service.files().create(
            body=metadata, media_body=media, fields="id"
        ).execute()
        print("Created new file on Drive")
    print("Upload complete!")


# ════════════════════════════════════════════════════════
# HELPER FUNCTIONS
# ════════════════════════════════════════════════════════

def clean_text(text):
    """Normalizes text for duplicate detection"""
    text = str(text).lower()
    text = re.sub(r'[^a-z0-9 ]', '', text)
    return " ".join(text.split())


def has_car_context(text):
    t = text.lower()
    return any(w in t for w in CAR_CONTEXT_WORDS)


def comment_matches(text):
    """Returns (True, group_name) if comment matches any keyword"""
    if not text or len(text) < 25:
        return False, None
    text_lower = text.lower()
    for group, keywords in KEYWORDS.items():
        for kw in keywords:
            if kw.upper() in REQUIRES_CAR_CONTEXT:
                if re.search(r'\b' + re.escape(kw) + r'\b',
                             text, re.IGNORECASE):
                    if has_car_context(text):
                        return True, group
            else:
                if kw.lower() in text_lower:
                    return True, group
    return False, None


def classify(text):
    matched, group = comment_matches(text)
    return group if matched else "Other"


def get_sentiment(text):
    """VADER sentiment with automotive tuning"""
    if not text or len(str(text).strip()) < 5:
        return "Neutral", 0.0
    sentences = [s.strip() for s in
                 str(text).replace('\n', '. ').split('.')
                 if len(s.strip()) > 10]
    if len(sentences) > 1:
        scores   = [analyzer.polarity_scores(s) for s in sentences[:15]]
        compound = float(np.mean([s['compound'] for s in scores]))
    else:
        compound = analyzer.polarity_scores(text)['compound']

    if compound >= 0.15:
        label = "Positive"
    elif compound <= -0.10:
        label = "Negative"
    else:
        label = "Neutral"
    return label, round(compound, 4)


def generate_comment_id(video_id, text, date):
    return hashlib.md5(
        f"{video_id}_{text}_{date}".encode()
    ).hexdigest()


def get_next_id(df):
    if df.empty or "Complaint_ID" not in df.columns:
        return 1
    nums = df["Complaint_ID"].dropna().str.replace(
        "CMP", "", regex=False
    )
    nums = pd.to_numeric(nums, errors="coerce").dropna()
    return int(nums.max()) + 1 if not nums.empty else 1


def load_last_run():
    try:
        with open(LOG_FILE) as f:
            from datetime import datetime
            dt = datetime.fromisoformat(json.load(f)["last_run"])
            from datetime import timedelta
            if (datetime.now() - dt).total_seconds() < 3600:
                return datetime.now() - timedelta(days=7)
            return dt
    except:
        from datetime import datetime
        return datetime(2023, 1, 1)


def save_last_run():
    from datetime import datetime
    with open(LOG_FILE, "w") as f:
        json.dump({"last_run": datetime.now().isoformat()}, f)


# ════════════════════════════════════════════════════════
# YOUTUBE FUNCTIONS
# ════════════════════════════════════════════════════════

def search_videos(query, max_results=8):
    try:
        res = youtube.search().list(
            part="snippet",
            q=query,
            type="video",
            maxResults=max_results,
            regionCode="IN",
            relevanceLanguage="en",
            order="relevance",
            publishedAfter="2020-01-01T00:00:00Z"
        ).execute()
        return [
            {
                "video_id": item["id"]["videoId"],
                "title":    item["snippet"]["title"],
                "channel":  item["snippet"]["channelTitle"],
                "url":      f"https://youtube.com/watch?v={item['id']['videoId']}"
            }
            for item in res.get("items", [])
        ]
    except Exception as e:
        print(f"  Search error: {str(e)[:60]}")
        return []


def get_comments(video):
    comments = []
    try:
        req = youtube.commentThreads().list(
            part="snippet",
            videoId=video["video_id"],
            maxResults=100,
            textFormat="plainText",
            order="relevance"
        )
        pages = 0
        while req and pages < 3:
            res   = req.execute()
            pages += 1
            for item in res.get("items", []):
                c = item["snippet"]["topLevelComment"]["snippet"]
                text = c["textDisplay"].strip()
                matched, group = comment_matches(text)
                if matched:
                    comments.append({
                        "comment_id": generate_comment_id(
                            video["video_id"], text, c["publishedAt"]
                        ),
                            "video_id": video["video_id"],  
                        "text":        text,
                        "date":        c["publishedAt"][:10],
                        "likes":       int(c.get("likeCount", 0)),
                        "video_title": video["title"],
                        "video_url":   video["url"],
                        "keyword_group": group,
                    })
            req = youtube.commentThreads().list_next(req, res)
    except Exception as e:
        err = str(e)
        if "disabled" not in err.lower() and "403" not in err:
            print(f"  Comment error: {err[:60]}")
    return comments


# ════════════════════════════════════════════════════════
# GEMINI ANALYSIS
# ════════════════════════════════════════════════════════

GEMINI_PROMPT = '''
You are an automotive quality analyst in India.
Analyse this YouTube comment and extract structured information.
Be LENIENT — keep comment even if some fields missing.
Write exactly "Not specified" for unknown fields.
Reply ONLY with valid JSON. No markdown, no backticks.

Comment: "{comment}"
Video title: "{title}"
Category: "{group}"

{{
  "is_useful": true or false,
  "fn_type": "FN1 if broken/malfunctioning. FN2 if feedback/opinion.",
  "year": "Year e.g. 2024. Not specified if unclear.",
  "model": "Car name e.g. Tata Nexon EV. Not specified if unclear.",
  "defect_summary": "One sentence — what is owner reporting?",
  "cause": "Root cause if mentioned. Not specified if unknown.",
  "action": "What was done. Not specified if not mentioned.",
  "severity": "High if safety risk or breakdown. Medium if performance issue. Low if minor feedback."
}}
'''


def run_gemini(comment_data):
    prompt = GEMINI_PROMPT.format(
        comment=comment_data.get("text", "")[:600],
        title=comment_data.get("video_title", "")[:100],
        group=comment_data.get("keyword_group", "")
    )
    for attempt in range(3):
        try:
            resp = gemini_client.models.generate_content(
                model="gemini-2.0-flash", contents=prompt
            )
            raw = resp.text.strip()
            raw = raw.replace("```json","").replace("```","").strip()
            s   = raw.find("{")
            e   = raw.rfind("}") + 1
            if s != -1 and e > s:
                raw = raw[s:e]
            return json.loads(raw)
        except json.JSONDecodeError:
            time.sleep(2)
        except Exception as ex:
            if "429" in str(ex) or "quota" in str(ex).lower():
                print("  Rate limit — waiting 30s")
                time.sleep(30)
            else:
                time.sleep(3)
    return None


# ════════════════════════════════════════════════════════
# MAIN PIPELINE
# ════════════════════════════════════════════════════════

def main():
    print("=" * 60)
    print("PIPELINE STARTED")
    print("=" * 60)

    # ── Connect to Google Drive ──────────────────────────
    print("\n[1/7] Connecting to Google Drive...")
    drive_service = get_drive_service()

    # ── Download existing Excel ──────────────────────────
    print("\n[2/7] Downloading existing Excel...")
    file_exists = download_from_drive(drive_service)

    if file_exists:
        old_df = pd.read_excel(LOCAL_FILE)
        print(f"Existing rows loaded: {len(old_df)}")
    else:
        old_df = pd.DataFrame()
        print("Starting fresh — no existing file")

    # Load existing texts for deduplication
    if not old_df.empty and "Complaint_Text" in old_df.columns:
        old_texts = set(old_df["Complaint_Text"].apply(clean_text))
    else:
        old_texts = set()

    next_id = get_next_id(old_df)

    # ── Scrape YouTube ───────────────────────────────────
    print("\n[3/7] Scraping YouTube...")
    all_comments = []
    seen_videos  = set()

    for i, query in enumerate(SEARCH_QUERIES):
        print(f"  [{i+1}/{len(SEARCH_QUERIES)}] {query}")
        videos = search_videos(query)
        new_vids = [v for v in videos
                    if v["video_id"] not in seen_videos]

        for video in new_vids:
            seen_videos.add(video["video_id"])
            comments = get_comments(video)
            if comments:
                all_comments.extend(comments)
                print(f"    {video['title'][:50]} → {len(comments)} comments")

        time.sleep(random.uniform(0.5, 1.0))

    print(f"Total comments fetched: {len(all_comments)}")

    if not all_comments:
        print("No comments found. Saving unchanged file.")
        upload_to_drive(drive_service)
        save_last_run()
        return

    # ── Filter spam and duplicates ───────────────────────
    print("\n[4/7] Filtering...")
    spam_patterns = [
        r"^(nice|great|good|wow|amazing|superb|👍|❤️)\s*[!.]*$",
        r"first\s*(comment|view)",
        r"subscribe.*channel",
        r"^[\s\W\d]+$",
    ]
    useful_signals = [
        "problem","issue","fault","defect","not working","failed","broke",
        "error","complaint","battery","charging","tailgate","hybrid",
        "electric","motor","sensor","noise","rattle","mileage","range",
        "nahi","khatam","dikkat",
    ]

    filtered    = []
    seen_clean  = set(old_texts)  # include old texts to prevent duplicates

    for c in all_comments:
        text = c.get("text", "").strip()

        if len(text.split()) < 6:
            continue
        if any(re.search(p, text, re.IGNORECASE) for p in spam_patterns):
            continue

        clean = clean_text(text)
        if clean in seen_clean:
            continue
        seen_clean.add(clean)

        if sum(1 for s in useful_signals if s in text.lower()) < 1:
            continue

        filtered.append(c)

    print(f"After filtering: {len(all_comments)} → {len(filtered)} useful")

    if not filtered:
        print("Nothing passed filter. Uploading unchanged file.")
        upload_to_drive(drive_service)
        save_last_run()
        return

    # ── Gemini + VADER ───────────────────────────────────
    print(f"\n[5/7] Gemini + VADER analysis ({len(filtered)} comments)...")
    new_data = []

    for i, c in enumerate(filtered):
        print(f"  [{i+1}/{len(filtered)}] ", end="", flush=True)

        # Gemini analysis
        analysis = run_gemini(c)

        # If Gemini fails — use fallback values
        if analysis is None:
            analysis = {
                "is_useful":     True,
                "fn_type":       "Not specified",
                "year":          "Not specified",
                "model":         "Not specified",
                "defect_summary": c["text"][:150],
                "cause":         "Not specified",
                "action":        "Not specified",
                "severity":      "Medium",
            }

        if not analysis.get("is_useful", True):
            print("skipped by Gemini")
            continue

        # VADER sentiment
        sentiment_label, compound_score = get_sentiment(c["text"])

        # Owner mapping
        comp          = c.get("keyword_group", "Other")
        owner, email  = OWNERS.get(comp, OWNERS["Other"])

        complaint_id  = f"CMP{next_id:05d}"
        next_id      += 1

        new_data.append({
            "Complaint_ID":    complaint_id,
            "Date":            c["date"],
            "Source":          "YouTube",
            "Component":       comp,
            "Sub_Issue":       analysis.get("model", "Not specified"),
            "Complaint_Text":  c["text"],
            "Defect_Summary":  analysis.get("defect_summary", "Not specified"),
            "Cause":           analysis.get("cause", "Not specified"),
            "Action":          analysis.get("action", "Not specified"),
            "FN_Type":         analysis.get("fn_type", "Not specified"),
            "Year":            analysis.get("year", "Not specified"),
            "Severity":        analysis.get("severity", "Medium"),
            "Sentiment":       sentiment_label,
            "VADER_Score":     compound_score,
            "Likes":           c.get("likes", 0),
            "Owner":           owner,
            "Owner_Email":     email,
            "Status":          "Open",
            "Video_ID":        c.get("video_id", ""), 
            "Video_Title":     c.get("video_title", ""),
            "Video_URL":       c.get("video_url", ""),
            "Week_Added":      pd.Timestamp.now().strftime("Week %W %Y"),
        })

        print(f"{sentiment_label} | {comp} | "
              f"{analysis.get('model','?')[:25]}")

        time.sleep(random.uniform(1.2, 2.0))

    print(f"\nNew rows created: {len(new_data)}")

    if not new_data:
        print("No new data after Gemini. Uploading unchanged file.")
        upload_to_drive(drive_service)
        save_last_run()
        return

    # ── Merge old + new ─────────────────────────────────
    print("\n[6/7] Merging and saving Excel...")
    df_new   = pd.DataFrame(new_data)
    final_df = pd.concat([old_df, df_new], ignore_index=True)
    final_df["YouTube_Link"] = "https://www.youtube.com/watch?v=" + final_df["Video_ID"]

    # Final safety dedup
    final_df["_clean"] = final_df["Complaint_Text"].apply(clean_text)
    before             = len(final_df)
    final_df.drop_duplicates(subset=["_clean"], inplace=True)
    final_df.drop(columns=["_clean"], inplace=True)
    final_df.reset_index(drop=True, inplace=True)
    print(f"Dedup: {before} → {len(final_df)} rows "
          f"({before - len(final_df)} duplicates removed)")

    # Save with formatting
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    import openpyxl.utils

    wb = Workbook()
    ws = wb.active
    ws.title = "Complaints"

    cols = list(final_df.columns)

    # Header
    hfill = PatternFill("solid", fgColor="1F3864")
    hfont = Font(color="FFFFFF", bold=True, size=11)
    for c_idx, col in enumerate(cols, 1):
        cell           = ws.cell(row=1, column=c_idx, value=col)
        cell.fill      = hfill
        cell.font      = hfont
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"

    # Fills
    new_fill = PatternFill("solid", fgColor="D6EAF8")  # blue — new this week
    neg_fill = PatternFill("solid", fgColor="FFCCCC")
    pos_fill = PatternFill("solid", fgColor="CCFFCC")
    neu_fill = PatternFill("solid", fgColor="FFFACC")
    thin     = Side(style="thin", color="DDDDDD")
    border   = Border(left=thin, right=thin, top=thin, bottom=thin)

    new_start = len(old_df) 

    for c_idx, col in enumerate(cols, 1):
    val = row.get(col, "")

    if col == "YouTube_Link" and val:
        cell = ws.cell(row=row_num, column=c_idx)
        cell.value = "Watch Video"
        cell.hyperlink = val
        cell.style = "Hyperlink"
    else:
        cell = ws.cell(row=row_num, column=c_idx, value=val)

    cell.fill = fill
    cell.border = border
    cell.alignment = Alignment(wrap_text=True, vertical="top")
    for r_idx, row in final_df.iterrows():
        row_num   = r_idx + 2
        is_new    = r_idx >= new_start
        sentiment = str(row.get("Sentiment", "")).lower()

        fill = (new_fill if is_new else
                neg_fill if sentiment == "negative" else
                pos_fill if sentiment == "positive" else
                neu_fill)

        for c_idx, col in enumerate(cols, 1):
            cell           = ws.cell(row=row_num, column=c_idx,
                                     value=row.get(col, ""))
            cell.fill      = fill
            cell.border    = border
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    # Column widths
    for c_idx in range(1, len(cols) + 1):
        col_letter = openpyxl.utils.get_column_letter(c_idx)
        ws.column_dimensions[col_letter].width = 20
    if "Complaint_Text" in cols:
        ws.column_dimensions[
            openpyxl.utils.get_column_letter(
                cols.index("Complaint_Text") + 1
            )
        ].width = 60
    if "Defect_Summary" in cols:
        ws.column_dimensions[
            openpyxl.utils.get_column_letter(
                cols.index("Defect_Summary") + 1
            )
        ].width = 50

    wb.save(LOCAL_FILE)
    print(f"Excel saved: {len(final_df)} total rows "
          f"({len(new_data)} new highlighted in blue)")

    # ── Upload to Google Drive ───────────────────────────
    print("\n[7/7] Uploading to Google Drive...")
    upload_to_drive(drive_service)

    save_last_run()

    print("\n" + "=" * 60)
    print("PIPELINE COMPLETE")
    print(f"Total rows in Excel : {len(final_df)}")
    print(f"New rows this run   : {len(new_data)}")
    print(f"File on Drive       : {LOCAL_FILE}")
    print("=" * 60)


if __name__ == "__main__":
    main()
