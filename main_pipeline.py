# main_pipeline.py — 50 most relevant, diverse, non-duplicate comments per run

import os, io, re, json, time, random, hashlib, warnings
import numpy as np
import pandas as pd
warnings.filterwarnings("ignore")

from googleapiclient.discovery import build
from google import genai
from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer
from gdrive_sync import download_from_onedrive, upload_to_onedrive

# ── KEYS ─────────────────────────────────────────────────────────
YOUTUBE_API_KEY = os.getenv("YOUTUBE_API_KEY", "")
GEMINI_API_KEY  = os.getenv("GEMINI_API_KEY",  "")
LOCAL_FILE      = "FINAL_STRUCTURED_COMPLAINTS.xlsx"
LOG_FILE        = "yt_last_run.json"

# ── CAP ──────────────────────────────────────────────────────────
MAX_COMMENTS_PER_RUN      = 50   # hard cap — never more than 50 new rows
MAX_PER_KEYWORD_GROUP     = 20   # max 20 per group (EV/Hybrid/PBD)
MAX_PER_PROBLEM_TYPE      = 5    # max 5 comments on same problem type
MIN_COMMENT_WORDS         = 10   # minimum words to be considered
MAX_COMMENT_WORDS         = 200  # maximum words — avoids essays

# ── APIS ─────────────────────────────────────────────────────────
youtube       = build("youtube", "v3", developerKey=YOUTUBE_API_KEY)
gemini_client = genai.Client(api_key=GEMINI_API_KEY)
analyzer      = SentimentIntensityAnalyzer()

analyzer.lexicon.update({
    "breakdown":-3.0,"broke":-2.5,"failed":-2.5,"failure":-2.5,
    "defect":-2.5,"fault":-2.0,"malfunction":-2.5,"problem":-2.0,
    "issue":-1.8,"drain":-2.0,"stuck":-2.5,"rattle":-2.0,
    "recall":-3.0,"dangerous":-3.0,"frustrated":-2.5,
    "terrible":-3.0,"worst":-3.0,"avoid":-2.0,
    "smooth":2.0,"excellent":2.5,"amazing":2.5,"reliable":2.0,
    "recommend":2.0,"efficient":1.8,"satisfied":2.0,"best":2.5,
})

# ════════════════════════════════════════════════════════════════
# KEYWORDS
# ════════════════════════════════════════════════════════════════
KEYWORD_GROUPS = {
    "EV System": [
        "EV system","EV battery","battery drain","battery degradation",
        "battery dead","battery not charging","charge not working",
        "EV not charging","charging stopped","charging failed",
        "charging port","fast charging","DC charging","AC charging",
        "regenerative braking","electric motor failure","electric range",
        "range anxiety","range dropped","EV failure","state of charge",
        "electric vehicle","BMS","OBC","SOC","V2L",
        "Nexon EV","Tata EV","Punch EV","Tigor EV","MG ZS EV",
        "Windsor EV","Comet EV","Creta EV","Ioniq 5","Ioniq 6",
        "Mahindra EV","XUV400 EV","BYD Atto","BYD Seal","XEV 9e",
        "km on full charge","range on full charge","ev breakdown",
        "charging station problem","charge nahi ho raha",
        "battery khatam","range kam ho gayi",
    ],
    "Series Hybrid EV": [
        "series hybrid","hybrid electric vehicle","mild hybrid",
        "strong hybrid","full hybrid","self charging hybrid",
        "plug-in hybrid","hybrid mode","hybrid system",
        "hybrid system fault","hybrid battery issue",
        "hybrid not working","hybrid mileage","e-CVT","PHEV","MHEV",
        "Grand Vitara hybrid","Hyryder hybrid","Innova Hycross",
        "City hybrid","Invicto hybrid","Toyota hybrid","Maruti hybrid",
        "hybrid lag","e-CVT problem","hybrid jerky",
        "hybrid kitna deta hai","hybrid mode nahi aata",
    ],
    "Power Back Door": [
        "power back door","power tailgate","electric tailgate",
        "auto tailgate","hands free tailgate","power liftgate",
        "electric boot","automatic boot","smart tailgate",
        "kick sensor tailgate","foot sensor boot",
        "tailgate not opening","tailgate not closing","tailgate stuck",
        "tailgate sensor","tailgate malfunction","tailgate rattle",
        "tailgate noise","boot not closing","rear door motor",
        "liftgate motor","hands free boot","motorized tailgate",
        "electric dicky","power dicky","boot sensor",
        "foot gesture boot","dicky problem","dicky not opening",
        "boot khul nahi raha","dicky band nahi hoti",
    ]
}

OWNERS = {
    "EV System":        ("Nandini Sharma", "Nandinisharma8862@gmail.com"),
    "Series Hybrid EV": ("Nandini Singh",  "065038@fsm.ac.in"),
    "Power Back Door":  ("Tejas",          "Personaltej2909@gmail.com"),
    "Other":            ("General",        "general@maruti.com"),
}

# Problem type buckets — for diversity (max 5 per bucket)
PROBLEM_BUCKETS = {
    "battery_drain":    ["battery drain","battery dead","battery khatam",
                         "overnight drain","parasitic drain"],
    "charging_failure": ["charge not working","not charging","charging stopped",
                         "charging failed","charger not working"],
    "range_issue":      ["range anxiety","range dropped","range problem",
                         "km on full charge","range kam"],
    "motor_failure":    ["motor failure","motor noise","electric motor",
                         "motor problem"],
    "hybrid_mileage":   ["hybrid mileage","mileage problem","fuel efficiency",
                         "mileage kam","kitna deta"],
    "hybrid_system":    ["hybrid system fault","hybrid not working",
                         "hybrid lag","hybrid jerky","e-cvt problem"],
    "tailgate_sensor":  ["tailgate sensor","sensor fail","kick sensor",
                         "foot sensor","sensor not working"],
    "tailgate_stuck":   ["tailgate stuck","not opening","not closing",
                         "boot stuck","dicky nahi"],
    "tailgate_noise":   ["tailgate rattle","tailgate noise","boot rattle",
                         "dicky rattle"],
    "ev_software":      ["software update","software bug","ota update",
                         "firmware","software issue"],
    "ev_breakdown":     ["ev breakdown","breakdown","stalled","highway",
                         "towed","band ho gaya"],
}

SEARCH_QUERIES = [
    "Nexon EV battery problem India owner review",
    "Nexon EV charging issue India 2024",
    "Tata Punch EV range problem India",
    "MG Windsor EV problem India owner",
    "Creta EV battery drain India",
    "Mahindra XUV400 EV problem India",
    "BMS error Nexon EV India fix",
    "EV battery degradation India owner complaint",
    "Grand Vitara strong hybrid problem India",
    "Hyryder hybrid mileage real world India",
    "Innova Hycross hybrid fault India owner",
    "Honda City hybrid review problem India",
    "Invicto hybrid issue India 2024",
    "e-CVT problem India hybrid car",
    "power tailgate problem India SUV 2024",
    "electric tailgate not working India car",
    "tailgate sensor fail India owner review",
    "electric dicky problem India car owner",
    "power boot door malfunction India SUV",
    "hands free tailgate issue India review",
]

# Filter config
COMPLAINT_WORDS = [
    "problem","issue","fault","defect","not working","failed","failure",
    "error","broken","complaint","bad","poor","worst","terrible",
    "stopped","not opening","not closing","drain","noise","rattle",
    "stuck","malfunction","breakdown","recall","repair","replace",
    "warning","disappointed","frustrated","pathetic","useless","avoid",
    "nahi chal raha","band ho gaya","kharab","dikkat",
    "nahi ho raha","nahi deta","nahi khul raha",
]

INDIA_CONTEXT = [
    "india","indian","delhi","mumbai","bangalore","bengaluru","chennai",
    "hyderabad","pune","kolkata","maruti","tata","mahindra","nexon",
    "hyryder","vitara","hycross","innova","city hybrid","creta ev",
    "mg zs","windsor ev","xuv400","rupee","lakh","kmpl",
    "service centre","service center","dealer","showroom","emi",
    "punch ev","tigor","brezza","baleno",
]

COMPONENT_MUST_HAVE = {
    "EV System": [
        "battery","charging","charge","range","ev","electric",
        "bms","soc","motor","kwh","plug","charger",
    ],
    "Series Hybrid EV": [
        "hybrid","mileage","fuel","petrol","generator","e-cvt",
        "self charging","strong hybrid","electric mode",
    ],
    "Power Back Door": [
        "tailgate","boot","dicky","door","liftgate","trunk",
        "sensor","kick","hands free","motorized",
    ]
}

SPAM_PATTERNS = [
    r"^(nice|great|good|wow|amazing|superb|👍|❤️|🔥|lol)\s*[!.]*$",
    r"first\s*(comment|view)",
    r"(subscribe|like).{0,20}(channel|karo|please)",
    r"check.{0,10}(my|out).{0,20}channel",
    r"^\d+\s*(likes?|views?)",
    r"^[\s\W\d]{0,5}$",
]

REQUIRES_BOUNDARY = {"EV","BMS","OBC","SOC","V2L","PHEV","MHEV"}

# ════════════════════════════════════════════════════════════════
# HELPERS
# ════════════════════════════════════════════════════════════════

def clean_text(text):
    t = str(text).lower()
    t = re.sub(r'[^a-z0-9 ]', '', t)
    return " ".join(t.split())


def comment_matches_keyword(text):
    if not text or len(text) < 20:
        return False, None
    tl = text.lower()
    for group, keywords in KEYWORD_GROUPS.items():
        for kw in keywords:
            if kw.upper() in REQUIRES_BOUNDARY:
                if re.search(r'\b' + re.escape(kw) + r'\b',
                             text, re.IGNORECASE):
                    return True, group
            else:
                if kw.lower() in tl:
                    return True, group
    return False, None


def get_problem_bucket(text):
    """Identifies which problem type this comment falls into"""
    tl = text.lower()
    for bucket, signals in PROBLEM_BUCKETS.items():
        if any(s in tl for s in signals):
            return bucket
    return "other"


def score_comment(c):
    """
    Scores a comment for relevance quality.
    Higher score = more useful, more specific, more likely to be a real complaint.
    Max possible score = 100
    """
    text  = c.get("text","")
    likes = int(c.get("likes", 0))
    words = len(text.split())
    tl    = text.lower()
    score = 0

    # Length sweet spot — 15 to 100 words is ideal for a complaint
    if 15 <= words <= 100:
        score += 25
    elif 10 <= words <= 150:
        score += 15
    else:
        score += 5

    # Likes — real people upvoted it
    if likes >= 10:
        score += 20
    elif likes >= 5:
        score += 15
    elif likes >= 1:
        score += 8

    # Contains specific model name — more professional
    indian_models = [
        "nexon ev","punch ev","tigor ev","windsor ev","comet ev",
        "creta ev","xuv400","xev 9e","grand vitara","hyryder",
        "innova hycross","city hybrid","invicto",
    ]
    if any(m in tl for m in indian_models):
        score += 20

    # Contains specific technical detail — higher quality
    technical_terms = [
        "bms","soc","kwh","kmpl","km","service centre","warranty",
        "dealer","ota","firmware","sensor","motor","e-cvt","range",
    ]
    score += min(15, sum(3 for t in technical_terms if t in tl))

    # Contains specific complaint signal — not vague
    specific_complaints = [
        "not working","failed","breakdown","stuck","malfunction",
        "fault","defect","error","recall","replaced","repair",
    ]
    if any(sc in tl for sc in specific_complaints):
        score += 10

    # Penalty for very vague
    vague = ["nice","love it","great car","awesome","superb","excellent"]
    if all(v not in tl for v in vague):
        score += 5
    else:
        score -= 10

    # Penalty for non-India context
    if not any(ic in tl for ic in INDIA_CONTEXT):
        score -= 15

    return max(0, score)


def passes_filter(c, seen_clean, old_texts):
    text  = c.get("text","").strip()
    group = c.get("keyword_group","")
    words = text.split()

    if len(words) < MIN_COMMENT_WORDS:
        return False, "too_short"
    if len(words) > MAX_COMMENT_WORDS:
        return False, "too_long"

    for p in SPAM_PATTERNS:
        if re.search(p, text, re.IGNORECASE):
            return False, "spam"

    clean = clean_text(text)
    if clean in seen_clean or clean in old_texts:
        return False, "duplicate"

    tl = text.lower()

    if not any(cw in tl for cw in COMPLAINT_WORDS):
        return False, "no_complaint"

    if not any(ic in tl for ic in INDIA_CONTEXT):
        return False, "no_india_context"

    must = COMPONENT_MUST_HAVE.get(group, [])
    if must and not any(kw in tl for kw in must):
        return False, "no_component_word"

    return True, "passed"


def select_top_50(filtered_comments):
    """
    From all filtered comments, selects top 50 that are:
    1. Highest relevance score
    2. Diverse — max MAX_PER_PROBLEM_TYPE per problem bucket
    3. Balanced — max MAX_PER_KEYWORD_GROUP per keyword group
    """
    # Score every comment
    for c in filtered_comments:
        c["_score"]   = score_comment(c)
        c["_bucket"]  = get_problem_bucket(c.get("text",""))

    # Sort by score descending
    sorted_comments = sorted(
        filtered_comments,
        key=lambda x: x["_score"],
        reverse=True
    )

    selected        = []
    bucket_counts   = {}
    group_counts    = {}

    for c in sorted_comments:
        if len(selected) >= MAX_COMMENTS_PER_RUN:
            break

        group  = c.get("keyword_group","Other")
        bucket = c.get("_bucket","other")

        # Check group cap
        if group_counts.get(group, 0) >= MAX_PER_KEYWORD_GROUP:
            continue

        # Check bucket cap (diversity)
        if bucket_counts.get(bucket, 0) >= MAX_PER_PROBLEM_TYPE:
            continue

        selected.append(c)
        group_counts[group]   = group_counts.get(group, 0) + 1
        bucket_counts[bucket] = bucket_counts.get(bucket, 0) + 1

    print(f"\nTop {len(selected)} selected from {len(filtered_comments)} filtered:")
    print(f"  By group:")
    for g, cnt in sorted(group_counts.items()):
        print(f"    {g:<22}: {cnt}")
    print(f"  By problem type:")
    for b, cnt in sorted(bucket_counts.items(), key=lambda x: -x[1]):
        if cnt > 0:
            print(f"    {b:<22}: {cnt}")
    print(f"  Score range: "
          f"{selected[-1]['_score'] if selected else 0}"
          f" to {selected[0]['_score'] if selected else 0}")

    return selected


def get_sentiment(text):
    if not text or len(str(text).strip()) < 5:
        return "Neutral", 0.0
    sentences = [s.strip() for s in
                 str(text).replace('\n','. ').split('.')
                 if len(s.strip()) > 10]
    if len(sentences) > 1:
        scores   = [analyzer.polarity_scores(s) for s in sentences[:15]]
        compound = float(np.mean([s['compound'] for s in scores]))
    else:
        compound = analyzer.polarity_scores(text)['compound']
    label = ("Positive" if compound >=  0.15 else
             "Negative" if compound <= -0.10 else "Neutral")
    return label, round(compound, 4)


def get_next_sno(df):
    if df.empty or "S.No" not in df.columns:
        return 1
    nums = pd.to_numeric(df["S.No"], errors="coerce").dropna()
    return int(nums.max()) + 1 if not nums.empty else 1


def save_last_run():
    from datetime import datetime
    with open(LOG_FILE, "w") as f:
        json.dump({"last_run": datetime.now().isoformat()}, f)


# ════════════════════════════════════════════════════════════════
# YOUTUBE
# ════════════════════════════════════════════════════════════════

def search_videos(query, max_results=5):
    try:
        res = youtube.search().list(
            part="snippet", q=query, type="video",
            maxResults=max_results, regionCode="IN",
            relevanceLanguage="en", order="relevance",
            publishedAfter="2020-01-01T00:00:00Z"
        ).execute()
        return [
            {
                "video_id": item["id"]["videoId"],
                "title":    item["snippet"]["title"],
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
        req   = youtube.commentThreads().list(
            part="snippet", videoId=video["video_id"],
            maxResults=100, textFormat="plainText",
            order="relevance"
        )
        pages = 0
        while req and pages < 2:
            res   = req.execute()
            pages += 1
            for item in res.get("items", []):
                c    = item["snippet"]["topLevelComment"]["snippet"]
                text = c["textDisplay"].strip()
                matched, group = comment_matches_keyword(text)
                if matched:
                    comments.append({
                        "comment_id": hashlib.md5(
                            f"{video['video_id']}_{text[:50]}".encode()
                        ).hexdigest(),
                        "video_id":      video["video_id"],
                        "text":          text,
                        "date":          c["publishedAt"][:10],
                        "likes":         int(c.get("likeCount", 0)),
                        "video_title":   video["title"],
                        "video_url":     video["url"],
                        "keyword_group": group,
                    })
            req = youtube.commentThreads().list_next(req, res)
    except Exception as e:
        err = str(e)
        if "disabled" not in err.lower() and "403" not in err:
            print(f"  Comment error: {err[:60]}")
    return comments


# ════════════════════════════════════════════════════════════════
# GEMINI
# ════════════════════════════════════════════════════════════════

GEMINI_PROMPT = '''
You are an automotive quality analyst at Maruti Suzuki India.
Read this YouTube comment from an Indian car owner.
Extract structured complaint data. This must be a REAL complaint from a REAL owner.
Write "Not specified" for fields you cannot determine.
Reply ONLY with valid JSON — no markdown, no backticks.

Comment: "{comment}"
Video: "{title}"
Feature Category: "{group}"

{{
  "is_useful": true or false — false if not a genuine owner complaint,
  "system_technology": "Specific subcategory e.g. EV System (Battery Drain) or EV System (Charging Failure) or Series Hybrid EV (Mileage Drop) or Power Back Door (Sensor Failure)",
  "fn_type": "FN1 if something is broken or malfunctioning. FN2 if owner feedback or feature suggestion.",
  "month": "Month if mentioned e.g. March. Not specified if not mentioned.",
  "year": "Year if mentioned e.g. 2024. Not specified if not mentioned.",
  "model": "Full car model e.g. Tata Nexon EV or Maruti Grand Vitara Hybrid. Not specified if unclear.",
  "defect_summary": "One professional sentence describing exactly what the owner experienced.",
  "cause": "Root cause if mentioned e.g. BMS software bug, faulty sensor, water ingress. Not specified if unknown.",
  "action": "What was done e.g. visited service centre, software update resolved it. Not specified if not mentioned.",
  "sentiment": "Negative, Positive, or Neutral"
}}
'''


def run_gemini(c):
    prompt = GEMINI_PROMPT.format(
        comment=c.get("text","")[:600],
        title=c.get("video_title","")[:100],
        group=c.get("keyword_group","")
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
            err = str(ex)
            if "429" in err or "quota" in err.lower():
                print("  Rate limit — waiting 30s")
                time.sleep(30)
            else:
                time.sleep(3)
    return None


def fallback(c):
    return {
        "is_useful":        True,
        "system_technology": c.get("keyword_group","Not specified"),
        "fn_type":          "Not specified",
        "month":            "Not specified",
        "year":             "Not specified",
        "model":            "Not specified",
        "defect_summary":   c.get("text","")[:200],
        "cause":            "Not specified",
        "action":           "Not specified",
        "sentiment":        "Neutral",
    }


# ════════════════════════════════════════════════════════════════
# EXCEL — your exact column format
# ════════════════════════════════════════════════════════════════

EXCEL_COLUMNS = [
    "S.No",
    "System / Technology",
    "FN1 / FN2",
    "Month",
    "Year",
    "Model",
    "Defect / Feedback Summary",
    "Cause",
    "Action",
    "Sentiment",
    "VADER Score",
    "Owner Name",
    "Owner Email",
    "Status",
    "Source",
    "Video URL",
    "Date",
    "Original Comment",
]

COL_WIDTHS = [6,28,10,12,8,28,55,40,35,12,12,18,30,10,10,45,12,60]


def save_excel(final_df, new_start):
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    import openpyxl.utils

    wb  = Workbook()
    ws  = wb.active
    ws.title = "Complaints"

    hfill = PatternFill("solid", fgColor="1F3864")
    hfont = Font(color="FFFFFF", bold=True, size=11)
    for ci, col in enumerate(EXCEL_COLUMNS, 1):
        cell = ws.cell(row=1, column=ci, value=col)
        cell.fill      = hfill
        cell.font      = hfont
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
    ws.row_dimensions[1].height = 35
    ws.freeze_panes = "A2"
    for ci, w in enumerate(COL_WIDTHS, 1):
        ws.column_dimensions[
            openpyxl.utils.get_column_letter(ci)
        ].width = w

    new_fill = PatternFill("solid", fgColor="D6EAF8")
    neg_fill = PatternFill("solid", fgColor="FFCCCC")
    pos_fill = PatternFill("solid", fgColor="CCFFCC")
    neu_fill = PatternFill("solid", fgColor="FFFACC")
    thin     = Side(style="thin", color="DDDDDD")
    border   = Border(left=thin, right=thin, top=thin, bottom=thin)

    for ri, row in final_df.iterrows():
        rn        = ri + 2
        is_new    = ri >= new_start
        sentiment = str(row.get("Sentiment","")).lower()
        fill      = (new_fill if is_new else
                     neg_fill if sentiment == "negative" else
                     pos_fill if sentiment == "positive" else
                     neu_fill)

        values = [
            row.get("S.No",""),
            row.get("System / Technology",""),
            row.get("FN1 / FN2",""),
            row.get("Month",""),
            row.get("Year",""),
            row.get("Model",""),
            row.get("Defect / Feedback Summary",""),
            row.get("Cause",""),
            row.get("Action",""),
            row.get("Sentiment",""),
            row.get("VADER Score",""),
            row.get("Owner Name",""),
            row.get("Owner Email",""),
            row.get("Status","Open"),
            row.get("Source","YouTube"),
            row.get("Video URL",""),
            row.get("Date",""),
            row.get("Original Comment",""),
        ]

        for ci, val in enumerate(values, 1):
            cell = ws.cell(row=rn, column=ci, value=val)
            cell.fill      = fill
            cell.border    = border
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(LOCAL_FILE)


# ════════════════════════════════════════════════════════════════
# MAIN
# ════════════════════════════════════════════════════════════════

def main():
    print("=" * 60)
    print("PIPELINE STARTED")
    print(f"Cap: {MAX_COMMENTS_PER_RUN} most relevant comments per run")
    print("=" * 60)

    # 1 — Download
    print("\n[1/7] Downloading from Google Drive...")
    file_exists = download_from_onedrive()

    if file_exists:
        try:
            old_df = pd.read_excel(LOCAL_FILE)
            for col_check in ["Original Comment","Complaint_Text","text"]:
                if col_check in old_df.columns:
                    old_texts = set(
                        old_df[col_check].dropna().apply(clean_text)
                    )
                    break
            else:
                old_texts = set()
            print(f"Existing rows: {len(old_df)}")
            print(f"Existing comment fingerprints: {len(old_texts)}")
        except Exception as e:
            print(f"Could not read file ({e}) — starting fresh")
            old_df    = pd.DataFrame()
            old_texts = set()
    else:
        old_df    = pd.DataFrame()
        old_texts = set()
        print("No existing file — starting fresh")

    next_sno = get_next_sno(old_df)

    # 2 — Scrape
    print("\n[2/7] Scraping YouTube...")
    all_comments = []
    seen_videos  = set()

    for i, query in enumerate(SEARCH_QUERIES):
        print(f"  [{i+1}/{len(SEARCH_QUERIES)}] {query}")
        videos   = search_videos(query, max_results=5)
        new_vids = [v for v in videos
                    if v["video_id"] not in seen_videos]
        for video in new_vids:
            seen_videos.add(video["video_id"])
            comments = get_comments(video)
            if comments:
                all_comments.extend(comments)
                print(f"    → {video['title'][:40]} "
                      f"| {len(comments)} keyword-matched")
        time.sleep(random.uniform(0.5, 1.0))

    print(f"\nTotal keyword-matched: {len(all_comments)} "
          f"from {len(seen_videos)} videos")

    if not all_comments:
        print("No comments found.")
        save_last_run()
        return

    # 3 — Filter
    print("\n[3/7] Relevance filtering...")
    filtered   = []
    seen_clean = set()
    rejected   = {}

    for c in all_comments:
        passed, reason = passes_filter(c, seen_clean, old_texts)
        if passed:
            seen_clean.add(clean_text(c.get("text","")))
            filtered.append(c)
        else:
            rejected[reason] = rejected.get(reason, 0) + 1

    print(f"Before filter: {len(all_comments)}")
    print(f"After filter : {len(filtered)} relevant")
    for reason, count in sorted(rejected.items(), key=lambda x: -x[1]):
        print(f"  {reason:<22}: {count} rejected")

    if not filtered:
        print("Nothing passed filter.")
        save_last_run()
        return

    # 4 — Select top 50 most relevant + diverse
    print(f"\n[4/7] Selecting top {MAX_COMMENTS_PER_RUN} most relevant...")
    top_comments = select_top_50(filtered)

    # 5 — Gemini + VADER
    print(f"\n[5/7] Gemini + VADER ({len(top_comments)} comments)...")
    new_rows = []

    for i, c in enumerate(top_comments):
        print(f"  [{i+1}/{len(top_comments)}] "
              f"[score={c.get('_score',0)}] ", end="", flush=True)

        analysis = run_gemini(c)
        if analysis is None:
            print("Gemini failed → fallback")
            analysis = fallback(c)

        if not analysis.get("is_useful", True):
            print("Gemini: not useful → skipped")
            continue

        sentiment_label, vader_score = get_sentiment(c["text"])
        group        = c.get("keyword_group","Other")
        owner, email = OWNERS.get(group, OWNERS["Other"])

        new_rows.append({
            "S.No":                      next_sno,
            "System / Technology":       analysis.get("system_technology", group),
            "FN1 / FN2":                 analysis.get("fn_type","Not specified"),
            "Month":                     analysis.get("month","Not specified"),
            "Year":                      analysis.get("year","Not specified"),
            "Model":                     analysis.get("model","Not specified"),
            "Defect / Feedback Summary": analysis.get("defect_summary",
                                                      c["text"][:150]),
            "Cause":                     analysis.get("cause","Not specified"),
            "Action":                    analysis.get("action","Not specified"),
            "Sentiment":                 sentiment_label,
            "VADER Score":               vader_score,
            "Owner Name":                owner,
            "Owner Email":               email,
            "Status":                    "Open",
            "Source":                    "YouTube",
            "Video URL":                 c.get("video_url",""),
            "Date":                      c.get("date",""),
            "Original Comment":          c.get("text",""),
        })

        next_sno += 1
        print(f"{sentiment_label} | "
              f"{analysis.get('fn_type','?')} | "
              f"{analysis.get('model','?')[:18]} | "
              f"{analysis.get('defect_summary','?')[:35]}")

        time.sleep(random.uniform(1.2, 2.0))

    print(f"\nNew rows created: {len(new_rows)} "
          f"(max was {MAX_COMMENTS_PER_RUN})")

    if not new_rows:
        print("No new rows.")
        save_last_run()
        return

    # 6 — Merge
    print("\n[6/7] Merging...")
    df_new    = pd.DataFrame(new_rows)
    new_start = len(old_df)

    if not old_df.empty:
        for col in EXCEL_COLUMNS:
            if col not in old_df.columns:
                old_df[col] = ""
        final_df = pd.concat(
            [old_df[EXCEL_COLUMNS], df_new[EXCEL_COLUMNS]],
            ignore_index=True
        )
    else:
        final_df = df_new[EXCEL_COLUMNS].copy()

    # Final dedup
    final_df["_d"] = final_df["Original Comment"].apply(clean_text)
    before         = len(final_df)
    final_df.drop_duplicates(subset=["_d"], inplace=True)
    final_df.drop(columns=["_d"], inplace=True)
    final_df.reset_index(drop=True, inplace=True)
    if before > len(final_df):
        print(f"Safety dedup: removed {before - len(final_df)} rows")

    # Recalculate S.No sequentially
    final_df["S.No"] = range(1, len(final_df) + 1)

    print(f"Total: {len(final_df)} rows "
          f"(existing: {new_start} + new: {len(new_rows)})")

    # 7 — Save + Upload
    print("\n[7/7] Saving Excel and uploading...")
    save_excel(final_df, new_start)
    upload_to_onedrive()
    save_last_run()

    print("\n" + "=" * 60)
    print("PIPELINE COMPLETE")
    print(f"Total rows in Excel  : {len(final_df)}")
    print(f"New rows this run    : {len(new_rows)}")
    print(f"Cap used             : {len(new_rows)}/{MAX_COMMENTS_PER_RUN}")
    print("=" * 60)


if __name__ == "__main__":
    main()
