from linebot.v3.messaging import (
    TextMessage,
    FlexBox,
    FlexText,
)
from linebot.v3.messaging import (
    QuickReply,
    QuickReplyItem,
    MessageAction,
    FlexMessage,
)
from template import night_timeoff_template
from datetime import datetime, timedelta
import pytz
import pandas as pd
import gspread
import copy
import json
import os

SHEET_KEY="10o1RavT1RGKFccEdukG1HsEgD3FPOBOPMB6fQqTc_wI"
service_account_info = json.loads(os.getenv('GOOGLE_SERVICE_ACCOUNT_JSON'))
service_account_info['private_key'] = service_account_info['private_key'].replace("\\n", "\n")
gc = gspread.service_account_from_dict(service_account_info)

local_tz = 'Asia/Taipei'

def format_datetime(month, day):
    today = pytz.timezone(local_tz).localize(datetime.today())
    absence_day = today.replace(month=month, day=day)
    if absence_day < today:
        absence_day = absence_day.replace(year=absence_day.year+1)
    return absence_day

class State:
    def __init__(self):
        self.block = True
    
    def generate_message(self):
        pass

    def next(self, user_input, user_info):
        pass

    def block_for_next_message(self):
        return self.block

class DataCollect(State):
    def __init__(self):
        super(DataCollect, self).__init__()

    def generate_message(self, user_info):
        return {
            "user": None,
            "group": None
        }
    
    def next(self, user_input, user_info):
        splitted_input = user_input.split()
        if len(splitted_input) == 3 and len(splitted_input[0]) > 1 and splitted_input[1].isdigit():
            user_info["name"] = splitted_input[0]
            user_info["session"] = splitted_input[1]
            user_info["unit"] = splitted_input[2]
            return DataConfirm
        else:
            return DataError

class DataConfirm(State):
    def __init__(self):
        super(DataConfirm, self).__init__()

    def generate_message(self, user_info):
        option_items = []
        for option in ['個資正確', '個資錯誤']:
            option_items.append(QuickReplyItem(action=MessageAction(label=option, text=f"{option}")))
        message = [TextMessage(
            text=f"請確認個人資料\n姓名: {user_info['name']}\n梯次: {user_info['session']}\n服勤單位: {user_info['unit']}",
            quick_reply=QuickReply(items=option_items)
        )]
        return {
            "user": message,
            "group": None
        }
    
    def next(self, user_input, user_info):
        if user_input == "個資正確":
            return DataFinish
        else:
            return DataError

class DataError(DataCollect):
    def __init__(self):
        super(DataError, self).__init__()

    def generate_message(self, user_info):
        message = [TextMessage(
            text=f"請重新輸入您的中文姓名、梯次與服勤單位，中間用空白隔開 (Ex. 王小明 261 社家署)",
        )]
        return {
            "user": message,
            "group": None
        }
    
class DataFinish(State):
    def __init__(self):
        super(DataFinish, self).__init__()
        self.block = False

    def generate_message(self, user_info):
        message = [TextMessage(
            text=f"已記錄您的資料，現在您可以開啟選單來使用本系統",
        )]
        return {
            "user": message,
            "group": None
        }
    
    def next(self, user_input, user_info):
        return Normal

class Normal(State):
    def __init__(self):
        super(Normal, self).__init__()
        # self.block = False

    def generate_message(self, user_info):
        return {
            "user": None,
            "group": None
        }
    
    def next(self, user_input, user_info):
        if user_input == "== 請假 ==":
            return Absence
        elif user_input == "== 查看夜假 ==":
            return CheckNightTimeoff
        elif user_input == "== 管理員功能 ==":
            return Administration
        else:
            return Normal

class Absence(State):
    def __init__(self):
        super(Absence, self).__init__()

    def generate_message(self, user_info):
        option_items = []
        for option in ['夜假', '補休', '公差', '返回']:
            option_items.append(QuickReplyItem(action=MessageAction(label=option, text=f"{option}")))
        message = [TextMessage(
            text="請選擇請假種類",
            quick_reply=QuickReply(items=option_items)
        )]
        return {
            "user": message,
            "group": None
        }
    
    def next(self, user_input, user_info):
        if user_input == "夜假" or user_input == "補休" or user_input == "公差":
            user_info["absence_type"] = user_input
            return AbsenceDate
        elif user_input == "== 查看夜假 ==":
            return CheckNightTimeoff
        else:
            return Normal
        
class AbsenceDate(State):
    def __init__(self):
        super(AbsenceDate, self).__init__()

    def generate_message(self, user_info):
        option_items = []
        for option in ['今天', '明天', '取消']:
            option_items.append(QuickReplyItem(action=MessageAction(label=option, text=f"{option}")))
        message = [TextMessage(
            text="請輸入請假日期 (Ex. 1/3)，或按'今天'、'明天'，若要取消請按'取消'",
            quick_reply=QuickReply(items=option_items)
        )]
        return {
            "user": message,
            "group": None
        }

    def next(self, user_input, user_info):
        if "取消" in user_input:
            return Normal
        elif user_input == "今天":
            today = pytz.timezone(local_tz).localize(datetime.today())
            user_info["absence_date"] = today
            return AbsenceConfirm
        elif user_input == "明天":
            today = pytz.timezone(local_tz).localize(datetime.today())
            user_info["absence_date"] = today + timedelta(days=1)
            return AbsenceConfirm
        elif "/" in user_input and len(user_input.split("/")) == 2:
            month, day = user_input.split("/")
            month = month.lstrip("0")
            day = day.lstrip("0")
            if month.isdigit() and day.isdigit():
                user_info["absence_date"] = format_datetime(int(month), int(day))
                return AbsenceConfirm
            else:
                return AbsenceDateFormatError
        else:
            return AbsenceDateFormatError

class AbsenceDateFormatError(AbsenceDate):
    def __init__(self):
        super(AbsenceDateFormatError, self).__init__()
    
    def generate_message(self, user_info):
        option_items = []
        for option in ['今天', '明天', '取消']:
            option_items.append(QuickReplyItem(action=MessageAction(label=option, text=f"{option}")))
        message = [TextMessage(
            text=f"格式錯誤，請輸入請假日期 (Ex. 1/3)，或按'今天'、'明天'，若要取消請按'取消'",
            quick_reply=QuickReply(items=option_items)
        )]
        return {
            "user": message,
            "group": None
        }

class AbsenceConfirm(State):
    def __init__(self):
        super(AbsenceConfirm, self).__init__()

    def generate_message(self, user_info):
        option_items = []
        for option in ['確定', '返回']:
            option_items.append(QuickReplyItem(action=MessageAction(label=option, text=f"{option}")))
        message = [TextMessage(
            text=f"確定於 {user_info['absence_date'].strftime('%Y/%m/%d')} 請{user_info['absence_type']}",
            quick_reply=QuickReply(items=option_items)
        )]
        return {
            "user": message,
            "group": None
        }

    def next(self, user_input, user_info):
        if user_input == "確定":
            if user_info["absence_type"] == "夜假":
                return NightTimeoff
            else:
                return OtherTimeoff
        else:
            return Normal

class NightTimeoff(State):
    def __init__(self):
        super(NightTimeoff, self).__init__()
        self.block = False

    def read_sheet(self, name, session):
        sheet = gc.open_by_key(SHEET_KEY)
        df = pd.DataFrame(sheet.worksheet(f"{session}T{name}").get_all_records())
        available_night_timeoff = []
        for index, row in df.iterrows():
            if len(row["請假日期"]) == 0:
                available_night_timeoff.append(row["有效期限"])
        return available_night_timeoff

    def update_sheet(self, name, session, date):
        sheet = gc.open_by_key(SHEET_KEY)
        worksheet = sheet.worksheet(f"{session}T{name}")
        df = pd.DataFrame(worksheet.get_all_records())
        target_row_id = -1
        deadline = datetime.max
        for row_id, row in df.iterrows():
            if len(row["請假日期"]) == 0:
                year, month, day = [int(x) for x in row["有效期限"].split("/")]
                row_deadline = datetime(year=year, month=month, day=day)
                if row_deadline < deadline:
                    deadline = row_deadline
                    target_row_id = row_id
        worksheet.update_cell(row=target_row_id+2, col=5, value=date.strftime('%Y/%m/%d'))

    def generate_message(self, user_info):
        avaiable_night_timeoff = self.read_sheet(user_info["name"], user_info["session"])
        if len(avaiable_night_timeoff) == 0:
            message = [TextMessage(
                text=f"您目前沒有可用夜假",
            )]
            return {
                "user": message,
                "group": None
            }
        else:
            self.update_sheet(user_info["name"], user_info["session"], user_info["absence_date"])
            user_message = [TextMessage(
                text=f"已登記您的請假申請，請至群組確認請假資訊",
            )]
            group_message = [TextMessage(
                text=f"{user_info['absence_date'].strftime('%Y/%m/%d')} [{user_info['session']}梯次］{user_info['name']} {user_info['unit']} 夜假",
            )]
            return {
                "user": user_message,
                "group": group_message
            }

    def next(self, user_input, user_info):
        return Normal

class CheckNightTimeoff(NightTimeoff):
    def __init__(self):
        super(CheckNightTimeoff, self).__init__()
        self.block = False

    def generate_night_timeoff_box(self, id, deadline):
        year, month, day = [int(x) for x in deadline.split("/")]
        return FlexBox(
            layout="baseline",
            spacing="sm",
            contents=[
                FlexText(
                    text=f"{id}.",
                    flex=1,
                    size="sm",
                    color="#aaaaaa"
                ),
                FlexText(
                    text=f"{year}/{month:02d}/{day:02d}",
                    flex=5,
                    size="sm",
                    color="#666666",
                    wrap=True
                )
            ]
        )


    def generate_message(self, user_info):
        try:
            available_night_timeoff = self.read_sheet(user_info["name"], user_info["session"])
            flex_message = copy.deepcopy(night_timeoff_template)
            flex_message.body.contents[0].text += str(len(available_night_timeoff))
            for i, nigth_timeoff in enumerate(available_night_timeoff):
                flex_message.body.contents[1].contents.append(self.generate_night_timeoff_box(i+1, nigth_timeoff))
            message = [FlexMessage(
                alt_text="夜假",
                contents=flex_message
            )]
        except KeyError:
            message = [TextMessage(
                text="您的夜假資料尚未登入，請稍後再試",
            )]
        return {
            "user": message,
            "group": None
        }

    def next(self, user_input, user_info):
        return Normal

class OtherTimeoff(State):
    def __init__(self):
        super(OtherTimeoff, self).__init__()
        self.block = False
    
    def generate_message(self, user_info):
        user_message = [TextMessage(
            text=f"已登記您的請假申請，請至群組確認請假資訊",
        )]
        group_message = [TextMessage(
            text=f"{user_info['absence_date'].strftime('%Y/%m/%d')} [{user_info['session']}梯次］{user_info['name']} {user_info['unit']} {user_info['absence_type']}",
        )]
        return {
            "user": user_message,
            "group": group_message
        }

    def next(self, user_input, user_info):
        return Normal

class Administration(State):
    def __init__(self):
        super(Administration, self).__init__()
        self.block = False
    
    def generate_message(self, user_info):
        message = [TextMessage(
            text=f"尚未實作管理員功能",
        )]
        return {
            "user": message,
            "group": None
        }

    def next(self, user_input, user_info):
        return Normal