from flask import Flask
import pyodbc
import praw
from datetime import datetime, timezone
from typing import List, Dict, Any
import logging

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s %(levelname)s: %(message)s",
    datefmt="%d-%m-%Y %H:%M:%S"
)

def launch_flask() -> Flask:
    logging.debug("launching flask")
    try:
        app = Flask(__name__)
        logging.info("flask launched successfully")
    except Exception:
        logging.exception("flask failed to start")
    return app

def praw_setup() -> praw.Reddit:
    logging.debug("setting up praw")
    try:
        reddit = praw.Reddit(
            client_id='',
            client_secret='',
            user_agent=''
            )
        logging.info("praw setup completed")
    except Exception:
        logging.exception(f"praw setup failed")
    return reddit

def connect_string() -> str:
    logging.debug("configuring connection string")
    try:
        connection_string = ('')
        logging.info("connection string configured")
    except Exception:
        logging.exception("connection string configuration failed")
    return connection_string

def subreddit_filters() -> List[str]:
        logging.debug("setting up subreddit list")
        try:
            subreddit_list = [
            "programming", "learnprogramming", "datascience", "machinelearning",
            "python", "SQL", "coding", "opensource", "webdev", "dataisbeautiful"
            ]
            logging.info(f"subreddit filters: {subreddit_list}")
        except Exception:
            logging.exception(f"subreddit list failed. {subreddit_list}")
        return subreddit_list

def fetch_subreddit_data(subreddit_list: List[str], reddit: praw.Reddit) -> List[Dict[str, Any]]:
    logging.debug("fetching subreddit data")
    try:
        cutoff_timestamp = int(datetime.now(timezone.utc).timestamp() - 3600)
        return [
            {
                "id": post.id,
                "subreddit": sub,
                "title": post.title,
                "upvotes": post.ups,
                "upvote_ratio": post.upvote_ratio,
                "comments": post.num_comments,
                "created_utc": datetime.fromtimestamp(post.created_utc, timezone.utc)
            }
            for sub in subreddit_list
            for post in reddit.subreddit(sub).new()
            if post.created_utc <= cutoff_timestamp
        ]
    except Exception:
        logging.exception("subreddit post failed")

def post_to_database(posts: List[Dict[str, Any]], cursor: pyodbc):
    logging.debug("posting data to db")
    try:
        for post in posts:
            logging.info(f"posting: {post}")
            cursor.execute(
                """
                MERGE reddit_posts AS target
                USING (SELECT ? AS id, ? AS subreddit, ? AS title, ? AS upvotes, ? AS upvote_ratio, ? AS comments, ? AS created_utc) AS source
                ON target.id = source.id
                WHEN NOT MATCHED THEN
                    INSERT (id, subreddit, title, upvotes, upvote_ratio, comments, created_utc)
                    VALUES (source.id, source.subreddit, source.title, source.upvotes, source.upvote_ratio, source.comments, source.created_utc);
                """,
                (post['id'], post['subreddit'], post['title'], post['upvotes'], post['upvote_ratio'], post['comments'], post['created_utc'])
            )
        logging.info(f"posted {len(posts)} posts to db successfully")
    except Exception:
        logging.exception("failed to post data to db")

def register_routes(app: Flask) -> None:
    logging.debug("registering routes")
    @app.route('/', methods=['GET'])
    def run_etl() -> str:
        try:
            logging.debug("running ETL pipeline")
            reddit = praw_setup()
            connection_string = connect_string()
            subreddit_list = subreddit_filters()
            
            posts = fetch_subreddit_data(subreddit_list, reddit)

            with pyodbc.connect(connection_string) as conn:
                with conn.cursor() as cursor:
                    post_to_database(posts, cursor)
                    conn.commit()
            return "ETL process complete"
        except Exception:
            logging.exception("ETL pipeline failed")
            return "ETL process failed", 500

if __name__ == '__main__':
    logging.debug("starting app")
    app = launch_flask()
    register_routes(app)
    app.run(host='0.0.0.0', port=5000)
