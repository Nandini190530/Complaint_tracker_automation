import os
import pandas as pd
from googleapiclient.discovery import build
from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer

# =========================
# CONFIG
# =========================
API_KEY = os.environ.get("YOUTUBE_API_KEY")
youtube = build('youtube', 'v3', developerKey=API_KEY)

# =========================
# SEARCH VIDEOS
# =========================
def search_videos(query, max_results=5):
    request = youtube.search().list(
        part="snippet",
        q=query,
        type="video",
        maxResults=max_results
    )
    response = request.execute()
    return [item['id']['videoId'] for item in response['items']]

# =========================
# GET COMMENTS
# =========================
def get_youtube_comments(video_id, max_comments=100):
    comments = []

    request = youtube.commentThreads().list(
        part="snippet",
        videoId=video_id,
        maxResults=100,
        textFormat="plainText"
    )

    count = 0

    while request and count < max_comments:
        response = request.execute()

        for item in response.get('items', []):
            c = item['snippet']['topLevelComment']['snippet']

            text = c['textDisplay']
            date = c['publishedAt']
            likes = c.get('likeCount', 0)

            comment_id = f"{video_id}_{date}_{text[:20]}"

            comments.append({
                "comment_id": comment_id,
                "text": text,
                "date": date,
                "likes": likes,
                "video_id": video_id
            })

            count += 1
            if count >= max_comments:
                break

        request = youtube.commentThreads().list_next(request, response)

    return comments

# =========================
# FILTER
# =========================
def filter_comments(all_comments):
    filtered = []

    for c in all_comments:
        text = str(c.get("text", "")).lower()
        likes = c.get("likes", 0)

        if any(word in text for word in ["issue", "problem", "not working", "failure"]):
            score = 5
            score += min(len(text)//50, 3)
            score += min(likes//5, 3)

            c["score"] = score
            filtered.append(c)

    return sorted(filtered, key=lambda x: x["score"], reverse=True)[:100]

# =========================
# MODEL
# =========================
def extract_model(text):
    text = text.lower()
    if "nexon" in text:
        return "Nexon EV"
    elif "tigor" in text:
        return "Tigor EV"
    elif "xuv400" in text:
        return "XUV400"
    else:
        return "Unknown"

# =========================
# SENTIMENT
# =========================
analyzer = SentimentIntensityAnalyzer()

def get_sentiment(text):
    score = analyzer.polarity_scores(text)["compound"]
    if score >= 0.05:
        return "Positive"
    elif score <= -0.05:
        return "Negative"
    else:
        return "Neutral"

# =========================
# MAIN PIPELINE
# =========================
def run():

    queries = [
        "EV battery problem India",
        "electric vehicle complaints",
        "hybrid car problems India",
        "tailgate issue car"
    ]

    all_comments = []

    # STEP 1 — FETCH
    for q in queries:
        vids = search_videos(q)
        for v in vids:
            all_comments.extend(get_youtube_comments(v))

    print("Fetched:", len(all_comments))

    # STEP 2 — FILTER
    filtered = filter_comments(all_comments)
    print("Filtered:", len(filtered))

    df = pd.DataFrame(filtered)

    # STEP 3 — PROCESS
    df["Model"] = df["text"].apply(extract_model)
    df["Sentiment"] = df["text"].apply(get_sentiment)
    df["Severity"] = df["text"].apply(
        lambda x: "High" if "not working" in x.lower() else "Medium"
    )
    df["Status"] = "Open"
    df["Source"] = "YouTube"

    # STEP 4 — REMOVE DUPLICATES
    df.drop_duplicates(subset=["comment_id"], inplace=True)

    # STEP 5 — SAVE
    df.to_excel("FINAL_STRUCTURED_COMPLAINTS.xlsx", index=False)

    print("✅ File created")

# =========================
# RUN
# =========================
if __name__ == "__main__":
    run()
