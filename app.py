from flask import Flask, request, abort

from linebot.v3 import (
    WebhookHandler
)
from linebot.v3.exceptions import (
    InvalidSignatureError
)
from linebot.v3.messaging import (
    Configuration,
    ApiClient,
    MessagingApi,
    ReplyMessageRequest,
    PushMessageRequest,
)
from linebot.v3.webhooks import (
    MessageEvent,
    TextMessageContent
)
import pickle
import os
from state import *

from pymongo.mongo_client import MongoClient
from pymongo.server_api import ServerApi

USERS_DATA_FILE="users_data.pkl"

app = Flask(__name__)
configuration = Configuration(access_token=os.getenv("LINE_CHANNEL_ACCESS_TOKEN"))
line_handler = WebhookHandler(os.getenv("LINE_CHANNEL_SECRET"))
mongo_client = MongoClient(os.getenv("MONGO_URI"), server_api=ServerApi('1'))
db = mongo_client['absence-record']
users_col = db['user']

users = dict()
# users_data = dict()

# if os.path.isfile(USERS_DATA_FILE):
#     with open(USERS_DATA_FILE, "rb") as f:
#         users_data = pickle.load(f)

# for k, v in users_data.items():
#     users[k] = {
#         "state": Normal(),
#         "user_info": v
#     }

@app.route("/callback", methods=['POST'])
def callback():
    # get X-Line-Signature header value
    signature = request.headers['X-Line-Signature']

    # get request body as text
    body = request.get_data(as_text=True)
    app.logger.info("Request body: " + body)

    # handle webhook body
    try:
        line_handler.handle(body, signature)
    except InvalidSignatureError:
        app.logger.info("Invalid signature. Please check your channel access token/channel secret.")
        abort(400)

    return 'OK'


@line_handler.add(MessageEvent, message=TextMessageContent)
def handle_message(event):
    # print(event)
    with ApiClient(configuration) as api_client:
        if event.source.type == "user":
            user_id = event.source.user_id
            line_bot_api = MessagingApi(api_client)
            reply = event.message.text.strip()
            messages = []
            if users.get(user_id) == None:
                user_info_from_db = users_col.find_one({"_id" : user_id})
                if user_info_from_db == None:
                    users[user_id] = {
                        "state": DataCollect(),
                        "user_info": {},
                    }
                else:
                    users[user_id] = {
                        "state": Normal(),
                        "user_info": {
                            "name": user_info_from_db["name"],
                            "session": user_info_from_db["session"],
                            "unit": user_info_from_db["unit"]
                        },
                    }
            elif reply in KEYWORD:
                users[user_id]["state"] = Normal()

            # print("before state: ", users[user_id]["state"])
            users[user_id]["state"] = users[user_id]["state"].next(
                reply,
                users[user_id]['user_info']
            )()
            messages = users[user_id]["state"].generate_message(users[user_id]["user_info"])
            
            if isinstance(users[user_id]["state"], DataFinish):
                # users_data[user_id] = users[user_id]['user_info']
                users_col.insert_one({
                    "_id": user_id,
                    "name": users[user_id]['user_info']["name"],
                    "session": users[user_id]['user_info']["session"],
                    "unit": users[user_id]['user_info']["unit"],
                })
                # with open(USERS_DATA_FILE, "wb") as f:
                #     pickle.dump(users_data, f)

            if not users[user_id]["state"].block_for_next_message():
                users[user_id]["state"] = users[user_id]["state"].next(
                    reply,
                    users[user_id]['user_info']
                )()
            # print("After state: ", users[user_id]["state"])
            if messages["user"]:     
                r = line_bot_api.reply_message_with_http_info(
                    ReplyMessageRequest(
                        reply_token=event.reply_token,
                        messages=messages["user"]
                    )
                )
                print("User response status code: ", r.status_code)
            if messages["group"]:
                r = line_bot_api.push_message_with_http_info(
                    PushMessageRequest(
                        to="Cbfeffdfa0e89cf85eafc61560e923fef",
                        messages=messages["group"]
                    )
                )
                print("Group response status code: ", r.status_code)
        else:
            print(event)

if __name__ == "__main__":
    app.run()