from flask import Flask, request, abort

from linebot.v3 import (WebhookHandler)
from linebot.v3.exceptions import (InvalidSignatureError)
from linebot.v3.messaging import (
    Configuration,
    ApiClient,
    MessagingApi,
    ReplyMessageRequest,
    PushMessageRequest,
)
from linebot.v3.webhooks import (MessageEvent, TextMessageContent, UnsendEvent)
import os
from state import *
import re

USERS_DATA_FILE = "users_data.pkl"

app = Flask(__name__)
configuration = Configuration(
    access_token=os.getenv("LINE_CHANNEL_ACCESS_TOKEN"))
line_handler = WebhookHandler(os.getenv("LINE_CHANNEL_SECRET"))
group_chat_id = os.getenv("GROUP_CHAT_ID")

users = dict()


@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']

    body = request.get_data(as_text=True)
    app.logger.info("Request body: " + body)

    try:
        line_handler.handle(body, signature)
    except InvalidSignatureError:
        app.logger.info(
            "Invalid signature. Please check your channel access token/channel secret."
        )
        abort(400)

    return 'OK'


@line_handler.add(MessageEvent, message=TextMessageContent)
def handle_message(event):
    app.logger.info(event)
    with ApiClient(configuration) as api_client:
        user_id = event.source.user_id
        if event.source.type == "user":
            line_bot_api = MessagingApi(api_client)
            reply = event.message.text.strip()
            messages = []
            if users.get(user_id) == None:
                app.logger.info(f"User {user_id} not found")
                user_info_from_db = users_col.find_one({"_id": user_id})
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
                reply, users[user_id]['user_info'])()
            try:
                messages = users[user_id]["state"].generate_message(
                    users[user_id]["user_info"])
            except Exception as e:
                app.logger.error(e)
                messages = {
                    "user": [TextMessage(text="系統錯誤，請稍後再試", )],
                    "group": None
                }
                users[user_id]["state"] = Normal()

            if isinstance(users[user_id]["state"], DataFinish):
                users_col.update_one({"_id": user_id}, {
                    "$set": {
                        "_id": user_id,
                        "name": users[user_id]['user_info']["name"],
                        "session": users[user_id]['user_info']["session"],
                        "unit": users[user_id]['user_info']["unit"],
                    }
                })

            if not users[user_id]["state"].block_for_next_message():
                users[user_id]["state"] = users[user_id]["state"].next(
                    reply, users[user_id]['user_info'])()
            # print("After state: ", users[user_id]["state"])
            if messages["user"]:
                r = line_bot_api.reply_message_with_http_info(
                    ReplyMessageRequest(reply_token=event.reply_token,
                                        messages=messages["user"]))
                print("User response status code: ", r.status_code)
            if messages["group"]:
                r = line_bot_api.push_message_with_http_info(
                    PushMessageRequest(to=group_chat_id,
                                       messages=messages["group"]))
                print("Group response status code: ", r.status_code)
        elif event.source.type == "group":
            reply = event.message.text.strip()
            splitted_reply = re.split(' |，|[|]|［|］', reply)
            splitted_reply = [s.strip() for s in splitted_reply if s.strip()]
            absence_date = splitted_reply[0]
            absence_month, absence_day = absence_date.split("/")
            session = re.findall(r'\d+', splitted_reply[1])[0]
            name = splitted_reply[2]
            unit = splitted_reply[3]
            absence_type = splitted_reply[4]
            if re.match(r"隔.{1}補(休|修)", absence_type):
                absence_type = "隔天補休"
            user_info = {
                "name":
                name,
                "session":
                session,
                "unit":
                unit,
                "absence_type":
                absence_type,
                "absence_date":
                format_datetime(int(absence_month), int(absence_day)),
                "id":
                user_id
            }
            if absence_type == "夜假":
                state = NightTimeoff()
            else:
                state = OtherTimeoff()
            state.generate_message(user_info)

            user_info_from_db = users_col.find_one({"_id": user_id})
            if user_info_from_db == None:
                users_col.insert_one({
                    "_id": user_id,
                    "name": name,
                    "session": session,
                    "unit": unit,
                })


@line_handler.add(UnsendEvent)
def handle_unseen(event):
    if event.source.type == "group":
        user_id = event.source.user_id
        user_info_from_db = users_col.find_one({"_id": user_id})
        user_info = {
            "name": user_info_from_db["name"],
            "session": user_info_from_db["session"],
            "unit": user_info_from_db["unit"]
        }

        newest_absence_record = get_absence_records(
            records_col,
            user_id=user_info_to_id(user_info['session'], user_info['unit'],
                                    user_info['name'])).sort("date",
                                                             -1).limit(1)
        user_info['absence_date'] = pytz.utc.localize(
            newest_absence_record["date"]).astimezone(taipei_timezone)
        user_info['absence_type'] = newest_absence_record["type"]
        state = FinishCancelTimeoff()
        state.generate_message(user_info)


if __name__ == "__main__":
    app.run(port=5002)
