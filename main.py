import requests
import pandas as pd
import hashlib
import re
import os
from googleapiclient.discovery import build
from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer

# =========================
# ENV VARIABLES (GitHub Secrets)
# =========================
YOUTUBE_API_KEY = os.getenv("YOUTUBE_API_KEY")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
EXCEL_FILE_PATH = os.getenv("EXCEL_FILE_PATH")

# =========================
# MICROSOFT GRAPH AUTH
# =========================
def get_access_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"

    data = {
        "client_id": CLIENT_ID,
        "scope": "https://graph.microsoft.com/.default",
        "client_secret": CLIENT_SECRET,
        "grant_type": "client_credentials"
    }

    response = requests.post(url, data=data)
    return response.json().get("access_token")


def download_excel(token):
    url = f"https://graph.microsoft.com/v1.0/me/drive/root:{EXCEL_FILE_PATH}:/content"
    headers = {"Authorization": f"Bearer {token}"}

    r = requests.get(url, headers=headers)

    if r.status_code == 200:
        with open("data.xlsx", "wb") as f:
            f.write(r.content)
        return True
    return False


def upload_excel(token):
    url = f"https://graph.microsoft.com/v1.0/me/drive/root:{EXCEL_FILE_PATH}:/content"
    headers = {"Authorization": f"Bearer {token}"}

    with open("data.xlsx", "rb") as f:
        requests.put(url, headers=headers, data=f)


# =========================
# YOUTUBE SETUP
# =========================
youtube = build('youtube', 'v3', developerKey=YOUTUBE_API_KEY)

def search_videos(query, max_results=5):
    req = youtube.search().list(
        part="snippet",
        q=query,
        type="video",
        maxResults=max_results
    )
    res = req.execute()
    return [item['id']['videoId'] for item in res['items']]


def get_video_title(video_id):
    try:
        res = youtube.videos().list(part="snippet", id=video_id).execute()
        return res['items'][0]['snippet']['title']
    except:
        return ""


def generate_comment_id(video_id, text, date):
    return hashlib.md5(f"{video_id}_{text}_{date}".encode()).hexdigest()


def get_comments(video_id):
    comments = []
    title = get_video_title(video_id)

    req = youtube.commentThreads().list(
        part="snippet",
        videoId=video_id,
        maxResults=100,
        textFormat="plainText"
    )

    while req:
        res = req.execute()

        for item in res.get('items', []):
            c = item['snippet']['topLevelComment']['snippet']

            comments.append({
                "comment_id": generate_comment_id(video_id, c['textDisplay'], c['publishedAt']),
                "text": c['textDisplay'],
                "date": c['publishedAt'],
                "likes": c.get('likeCount', 0),
                "video_title": title
            })

        req = youtube.commentThreads().list_next(req, res)

    return comments


# =========================
# TEXT CLEANING (FOR DUPLICATES)
# =========================
def clean_text(text):
    text = str(text).lower()
    text = re.sub(r'[^a-z0-9 ]', '', text)
    return text.strip()


# =========================
# CLASSIFICATION + OWNER
# =========================
KEYWORDS = {
    "EV System": ["battery", "charging", "range"],
    "Series Hybrid EV": ["hybrid", "mileage"],
    "Power Back Door": ["boot", "tailgate", "door"]
}

OWNERS = {
    "EV System": ("EV Team", "ev@company.com"),
    "Series Hybrid EV": ("Hybrid Team", "hybrid@company.com"),
    "Power Back Door": ("Body Team", "body@company.com"),
    "Other": ("General Team", "general@company.com")
}

def classify(text):
    t = text.lower()
    for comp, words in KEYWORDS.items():
        if any(w in t for w in words):
            return comp
    return "Other"


# =========================
# SENTIMENT
# =========================
analyzer = SentimentIntensityAnalyzer()

def get_sentiment(text):
    score = analyzer.polarity_scores(text)['compound']
    if score >= 0.05:
        return "Positive"
    elif score <= -0.05:
        return "Negative"
    return "Neutral"


# =========================
# GENERATE SEQUENTIAL ID
# =========================
def get_next_id(df):
    if df.empty or "Complaint_ID" not in df.columns:
        return 1

    nums = df["Complaint_ID"].str.replace("CMP", "").astype(int)
    return nums.max() + 1


# =========================
# MAIN
# =========================
def main():
    print("🚀 Starting process...")

    token = get_access_token()

    # STEP 1: Download existing file
    file_exists = download_excel(token)

    if file_exists:
        old_df = pd.read_excel("data.xlsx")
        print("Old file loaded:", len(old_df))
    else:
        old_df = pd.DataFrame()
        print("No existing file found")

    # STEP 2: Load old texts
    if not old_df.empty and "Complaint_Text" in old_df.columns:
        old_texts = set(old_df["Complaint_Text"].apply(clean_text))
    else:
        old_texts = set()

    # STEP 3: Prepare ID counter
    next_id = get_next_id(old_df)

    # STEP 4: Fetch YouTube data
    queries = [
        "EV battery problem India",
        "electric vehicle complaints",
        "Nexon EV issues review",
        "hybrid car problems India",
        "boot door not working car"
    ]

    all_comments = []

    for q in queries:
        print("Searching:", q)
        vids = search_videos(q)

        for v in vids:
            all_comments.extend(get_comments(v))

    print("Total comments fetched:", len(all_comments))

    # STEP 5: Process new data
    new_data = []

    for c in all_comments:
        text_clean = clean_text(c["text"])

        if text_clean in old_texts:
            continue

        comp = classify(c["text"])
        owner, email = OWNERS.get(comp, OWNERS["Other"])

        complaint_id = f"CMP{next_id:05d}"
        next_id += 1

        new_data.append({
            "Complaint_ID": complaint_id,
            "Date": c["date"][:10],
            "Component": comp,
            "Complaint_Text": c["text"],
            "Likes": c["likes"],
            "Sentiment": get_sentiment(c["text"]),
            "Owner": owner,
            "Owner_Email": email,
            "Status": "Open"
        })

    df_new = pd.DataFrame(new_data)

    # STEP 6: Merge
    final_df = pd.concat([old_df, df_new], ignore_index=True)

    # STEP 7: Remove duplicates again (safety)
    final_df["clean"] = final_df["Complaint_Text"].apply(clean_text)
    final_df.drop_duplicates(subset=["clean"], inplace=True)
    final_df.drop(columns=["clean"], inplace=True)

    # STEP 8: Save
    final_df.to_excel("data.xlsx", index=False)

    # STEP 9: Upload back to OneDrive
    upload_excel(token)

    print("✅ Done! Total records:", len(final_df))


if __name__ == "__main__":
    main()
