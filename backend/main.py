
import os
import requests
from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException, Query
from collections import Counter
from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer
load_dotenv()
from fastapi.middleware.cors import CORSMiddleware
app = FastAPI(title="TwitterSense AI API")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
API_KEY = os.getenv("TWEETAPI_API_KEY")
BASE_URL = "https://api.tweetapi.com/tw-v2"


def tweetapi_get(endpoint: str, params: dict):
    if not API_KEY:
        raise HTTPException(
            status_code=500,
            detail="TweetAPI key not configured"
        )

    try:
        response = requests.get(
            f"{BASE_URL}{endpoint}",
            params=params,
            headers={"X-API-Key": API_KEY},
            timeout=25
        )
        response.raise_for_status()
        return response.json()

    except requests.exceptions.HTTPError:
        raise HTTPException(
            status_code=response.status_code,
            detail=f"TweetAPI error: {response.text[:500]}"
        )
    except requests.exceptions.RequestException as e:
        raise HTTPException(
            status_code=502,
            detail=f"TweetAPI connection failed: {str(e)}"
        )


@app.get("/")
def home():
    return {"message": "TwitterSense AI Backend Running"}


@app.get("/health")
def health():
    return {"status": "ok"}


@app.get("/profile/{username}")
def get_profile(username: str):
    result = tweetapi_get(
        "/user/by-username",
        {"username": username.lstrip("@")}
    )
    return result


@app.get("/profile/{username}/tweets")
def get_profile_tweets(
    username: str,
    limit: int = Query(default=100, ge=1, le=100)
):
    # First, retrieve the user's numeric ID.
    profile_response = tweetapi_get(
        "/user/by-username",
        {"username": username.lstrip("@")}
    )

    profile = profile_response.get("data", {})
    user_id = profile.get("id")

    if not user_id:
        raise HTTPException(
            status_code=502,
            detail="User ID missing from TweetAPI response"
        )

    posts = []
    seen_ids = set()
    cursor = None
    pages_fetched = 0
    max_pages = 10

    # Fetch pages until we reach the requested limit.
    while len(posts) < limit and pages_fetched < max_pages:

        params = {"userId": user_id}

        if cursor:
            params["cursor"] = cursor

        timeline = tweetapi_get("/user/tweets", params)

        page_posts = timeline.get("data", [])

        if not isinstance(page_posts, list):
            raise HTTPException(
                status_code=502,
                detail="Unexpected tweet timeline response"
            )

        for post in page_posts:
            post_id = post.get("id")

            if not post_id or post_id in seen_ids:
                continue

            seen_ids.add(post_id)
            posts.append(post)

            if len(posts) >= limit:
                break

        pages_fetched += 1

        cursor = (
            timeline.get("pagination", {}).get("nextCursor")
        )

        if not cursor or not page_posts:
            break

    return {
        "username": profile.get("username", username),
        "name": profile.get("name"),
        "followers": profile.get("followerCount"),
        "requested_count": limit,
        "fetched_count": len(posts),
        "pages_fetched": pages_fetched,
        "has_more": bool(cursor),
        "tweets": posts
    }
sentiment_analyzer = SentimentIntensityAnalyzer()

CATEGORY_KEYWORDS = {
    "Technology & AI": [
        "ai", "artificial intelligence", "machine learning",
        "software", "technology", "robot", "computer", "coding"
    ],
    "Business & Finance": [
        "business", "market", "stock", "finance",
        "economy", "investment", "company", "startup"
    ],
    "Politics": [
        "government", "election", "politics", "president",
        "minister", "policy", "parliament"
    ],
    "Sports": [
        "cricket", "football", "basketball", "tennis",
        "match", "player", "tournament", "olympics"
    ],
    "Entertainment": [
        "movie", "film", "music", "actor",
        "actress", "celebrity", "song", "cinema"
    ],
    "Science & Space": [
        "science", "nasa", "space", "planet",
        "research", "satellite", "mars", "earth"
    ],
    "Health": [
        "health", "fitness", "medical", "doctor",
        "hospital", "disease", "medicine"
    ],
    "Environment": [
        "climate", "environment", "pollution",
        "forest", "wildlife", "sustainability"
    ],
    "Education": [
        "education", "student", "school",
        "university", "college", "learning"
    ],
    "Fashion & Lifestyle": [
    "fashion", "clothing", "apparel", "outfit",
    "dress", "style", "trendy", "wear",
    "shoes", "sneakers", "beauty", "skincare",
    "makeup", "accessories", "collection",
    "shopping", "sale", "discount"
],

"E-commerce & Retail": [
    "ecommerce", "e-commerce", "retail",
    "online shopping", "delivery", "order",
    "product", "customer", "marketplace",
    "shopping app", "cart"
],

"Food & Beverage": [
    "food", "restaurant", "recipe", "cooking",
    "beverage", "coffee", "drink", "cuisine"
],

"Travel & Tourism": [
    "travel", "tourism", "hotel", "flight",
    "vacation", "destination", "holiday"
]
}


def get_tweet_text(tweet):
    """Extract text from common TweetAPI response formats."""

    if not isinstance(tweet, dict):
        return ""

    for key in ["text", "fullText", "full_text"]:
        value = tweet.get(key)
        if isinstance(value, str) and value.strip():
            return value.strip()

    legacy = tweet.get("legacy", {})
    if isinstance(legacy, dict):
        return (
            legacy.get("full_text")
            or legacy.get("text")
            or ""
        ).strip()

    return ""



def classify_category(text):
    text_lower = text.lower()

    matched_categories = []

    for category, keywords in CATEGORY_KEYWORDS.items():
        matches = [
            keyword for keyword in keywords
            if keyword in text_lower
        ]

        if matches:
            matched_categories.append({
                "category": category,
                "matched_keywords": matches,
                "match_count": len(matches)
            })

    # Sort categories by number of matching keywords.
    matched_categories.sort(
        key=lambda x: x["match_count"],
        reverse=True
    )

    return [
        item["category"]
        for item in matched_categories
    ]


def analyze_post(tweet):
    text = get_tweet_text(tweet)

    if not text:
        return {
            "text": "",
            "category": [],
            "sentiment": "Unknown",
            "sentiment_score": 0
        }

    sentiment_result = sentiment_analyzer.polarity_scores(text)
    compound = sentiment_result["compound"]

    if compound >= 0.05:
        sentiment = "Positive"
    elif compound <= -0.05:
        sentiment = "Negative"
    else:
        sentiment = "Neutral"

    return {
        "text": text,
        "categories": classify_category(text),
        "sentiment": sentiment,
        "sentiment_score": round(compound, 3)
    }


@app.get("/analyze/{username}")
def analyze_profile(
    username: str,
    limit: int = Query(default=100, ge=1, le=100)
):
    result = get_profile_tweets(username, limit)

    tweets = result["tweets"]
    analyzed_posts = [analyze_post(tweet) for tweet in tweets]

    # Exclude posts whose text could not be extracted.
    valid_posts = [
        post for post in analyzed_posts
        if post["text"]
    ]

    sentiment_counts = Counter(
        post["sentiment"] for post in valid_posts
    )

    category_counts = Counter()

    for post in valid_posts:
        category_counts.update(post["categories"])

    total = len(valid_posts)

    return {
        "username": username,
        "fetched_count": len(tweets),
        "analyzed_count": total,
        "unclassified_count": len(tweets) - total,
        "sentiment_distribution": dict(sentiment_counts),
        "category_distribution": dict(category_counts),
        "analyzed_posts": valid_posts
    }

def extract_author(tweet):
    """Handle common TweetAPI/X response formats."""

    author = tweet.get("author") or tweet.get("user") or {}

    if not author and isinstance(tweet.get("core"), dict):
        user_result = tweet["core"].get("user_results", {})
        author = user_result.get("result", {})

        if author.get("legacy"):
            legacy = author["legacy"]
            return {
                "username": legacy.get("screen_name", ""),
                "name": legacy.get("name", ""),
                "followers": legacy.get("followers_count", 0)
            }

    return {
        "username": (
            author.get("username")
            or author.get("screen_name")
            or ""
        ),
        "name": author.get("name", ""),
        "followers": (
            author.get("followerCount")
            or author.get("followers_count")
            or 0
        )
    }


@app.get("/topic-search/{topic}")
def topic_search(
    topic: str,
    pages: int = Query(default=3, ge=1, le=5)
):
    if not topic.strip():
        raise HTTPException(status_code=400, detail="Enter a topic")

    accounts = {}
    matching_posts = []
    total_posts = 0
    cursor = None
    pages_fetched = 0

    for _ in range(pages):
        params = {
            "query": topic.strip(),
            "type": "Latest"
        }

        if cursor:
            params["cursor"] = cursor

        result = tweetapi_get("/search", params)

        posts = result.get("data", [])

        if not isinstance(posts, list):
            raise HTTPException(
                status_code=502,
                detail="Unexpected TweetAPI search response"
            )

        pages_fetched += 1

        for post in posts:
            author = extract_author(post)
            username = author["username"]

            if not username:
                continue

            username = username.lstrip("@")
            total_posts += 1

            if username not in accounts:
                accounts[username] = {
                    "username": username,
                    "name": author["name"] or username,
                    "followers": author["followers"],
                    "post_count": 0,
                    "positive": 0,
                    "negative": 0,
                    "neutral": 0
                }

            text = get_tweet_text(post)
            analysis = analyze_post(post)
            matching_posts.append({
            "text": text,
            "username": username,
            "name": author["name"] or username,
            "categories": analysis["categories"],
            "sentiment": analysis["sentiment"],
            "sentiment_score": analysis["sentiment_score"],
            "created_at": post.get("createdAt") or post.get("created_at"),
            "url": post.get("url")
        })

            account = accounts[username]
            account["post_count"] += 1

            if analysis["sentiment"] == "Positive":
                account["positive"] += 1
            elif analysis["sentiment"] == "Negative":
                account["negative"] += 1
            elif analysis["sentiment"] == "Neutral":
                account["neutral"] += 1

        cursor = result.get("pagination", {}).get("nextCursor")

        if not cursor or not posts:
            break

    ranked = sorted(
        accounts.values(),
        key=lambda x: (x["post_count"], x["followers"]),
        reverse=True
    )[:10]

    return {
        "topic": topic,
        "total_posts_retrieved": total_posts,
        "unique_accounts": len(accounts),
        "pages_fetched": pages_fetched,
        "ranking_basis": "Number of matching posts, then followers",
        "top_accounts": ranked,
        "matching_posts": matching_posts,
        "multi_label": True,
    }
