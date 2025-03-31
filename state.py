from linebot.v3.messaging import (TextMessage, FlexBox, FlexText, QuickReply,
                                  QuickReplyItem, MessageAction, FlexMessage)
from template import night_timeoff_template, absence_record_template, today_absence_template, all_absence_record_template
from datetime import datetime, timedelta
import pytz
import gspread
import copy
import json
import os

from pymongo.mongo_client import MongoClient
from pymongo.server_api import ServerApi
from mongo_util import *

NIGHT_TIMEOFF_SHEET_KEY = "10o1RavT1RGKFccEdukG1HsEgD3FPOBOPMB6fQqTc_wI"
# ABSENCE_RECORD_SHEET_KEY = "1TxClL3L0pDQAIoIidgJh7SP-BF4GaBD6KKfVKw0CLZQ"
service_account_info = json.loads(os.getenv('GOOGLE_SERVICE_ACCOUNT_JSON'))
service_account_info['private_key'] = service_account_info[
    'private_key'].replace("\\n", "\n")
gc = gspread.service_account_from_dict(service_account_info)
taipei_timezone = pytz.timezone('Asia/Taipei')

mongo_client = MongoClient(os.getenv("MONGO_URI"), server_api=ServerApi('1'))
db = mongo_client['absence-record']
users_col = db['user']
folders_col = db['folder']
records_col = db['record']

COMMAND_REQUEST_ABSENCE = "== 其他請假 =="
COMMAND_CANCEL_ABSENCE = "== 取消請假 =="
COMMAND_CHECK_NIGHT_TIMEOFF = "== 查看剩餘夜假 =="
COMMAND_CHECK_ABSENCE_RECORD = "== 請假紀錄 =="
COMMAND_CHECK_FULL_ABSENCE_RECORD = "== 完整請假紀錄 =="
COMMAND_CHECK_TODAY_ABSENCE = "== 今日請假役男 =="
COMMAND_REQUEST_TODAY_NIGHT_TIMEOFF = "== 請今晚夜假 =="
COMMAND_REQUEST_TOMORROW_TIMEOFF = "== 請隔天補休 =="
COMMAND_REQUEST_OFFICIAL_LEAVE = "== 請公假 =="
COMMAND_UPLOAD_PROOF = "== 上傳請假證明 =="
COMMAND_CHECK_SELF_INFO = "== 查看個人資料 =="
COMMAND_UPDATE_SELF_INFO = "== 更新個人資料 =="

KEYWORD = {
    COMMAND_REQUEST_ABSENCE, COMMAND_CANCEL_ABSENCE,
    COMMAND_CHECK_NIGHT_TIMEOFF, COMMAND_CHECK_ABSENCE_RECORD,
    COMMAND_CHECK_FULL_ABSENCE_RECORD, COMMAND_CHECK_TODAY_ABSENCE,
    COMMAND_REQUEST_TODAY_NIGHT_TIMEOFF, COMMAND_REQUEST_TOMORROW_TIMEOFF, COMMAND_REQUEST_OFFICIAL_LEAVE,
    COMMAND_CHECK_SELF_INFO, COMMAND_UPDATE_SELF_INFO
}

SUCCESS = "SUCCESS"

night_timeoff_headers = ["核發原因", "核發日期", "有效期限", "使用日期"]
absence_headers = ["請假日期", "假別"]


def user_id_to_info(user_id):
    session, unit, name = user_id.split("_")
    return session.strip("T"), unit, name


def user_info_to_id(session, unit, name):
    return f"{session}T_{unit}_{name}"


def format_datetime(month, day):
    today = get_today_date()
    absence_day = today.replace(month=month, day=day)
    if absence_day < today:
        absence_day = absence_day.replace(year=absence_day.year + 1)
    return absence_day


def valid_date(absence_date, absence_type):
    now = datetime.now(taipei_timezone)
    overtime = False
    if now.weekday() == 6:
        overtime = now.hour >= 22
    else:
        overtime = now.hour >= 20
    today = now.replace(hour=0, minute=0, second=0, microsecond=0)
    if (absence_type == "夜假" or absence_type == "隔天補休"
        ) and absence_date >= today + timedelta(days=1) * (overtime):
        return True
    elif absence_type == "公假" and absence_date >= today:
        return True
    else:
        return False


def get_valid_date():
    today = datetime.now(taipei_timezone)
    overtime = False
    if today.weekday() == 6:
        overtime = today.hour > 22
    else:
        overtime = (today.hour > 22) or (today.hour > 21 and today.minute > 30)
    today = today.replace(hour=0, minute=0, second=0, microsecond=0)

    return today + timedelta(days=1) * (overtime)


def get_today_date():
    today = datetime.now(taipei_timezone)
    return today.replace(hour=0, minute=0, second=0, microsecond=0)


def is_date_format(date_str, date_format="%Y/%m/%d"):
    try:
        datetime.strptime(date_str, date_format)
        return True
    except ValueError:
        return False


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
        return {"user": None, "group": None}

    def next(self, user_input, user_info):
        splitted_input = user_input.split()
        if len(splitted_input) == 3 and len(
                splitted_input[0]) > 1 and splitted_input[1].isdigit():
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
            option_items.append(
                QuickReplyItem(
                    action=MessageAction(label=option, text=f"{option}")))
        message = [
            TextMessage(
                text=
                f"請確認個人資料\n姓名: {user_info['name']}\n梯次: {user_info['session']}\n服勤單位: {user_info['unit']}",
                quick_reply=QuickReply(items=option_items))
        ]
        return {"user": message, "group": None}

    def next(self, user_input, user_info):
        if user_input == "個資正確":
            user_id = user_info_to_id(user_info['session'], user_info['unit'], user_info['name'])
            result = check_user_exists(folders_col, user_id)
            if result:
                return DataFinish
            else:
                return DataNotFound
        else:
            return DataError


class DataError(DataCollect):

    def __init__(self):
        super(DataError, self).__init__()

    def generate_message(self, user_info):
        message = [
            TextMessage(
                text=f"請重新輸入您的中文姓名、梯次與服勤單位，中間用空白隔開 (Ex. 王小明 261 社家署)", )
        ]
        return {"user": message, "group": None}


class DataNotFound(DataCollect):

    def __init__(self):
        super(DataNotFound, self).__init__()

    def generate_message(self, user_info):
        message = [
            TextMessage(
                text=f"您輸入的個資有誤，或輔導員尚未將您加入系統，請洽輔導員詢問，或重新輸入您的中文姓名、梯次與服勤單位，中間用空白隔開 (Ex. 王小明 261 社家署)", )
        ]
        return {"user": message, "group": None}


class DataFinish(State):

    def __init__(self):
        super(DataFinish, self).__init__()
        self.block = False

    def generate_message(self, user_info):
        message = [TextMessage(text=f"已記錄您的資料，現在您可以開啟選單來使用本系統", )]
        return {"user": message, "group": None}

    def next(self, user_input, user_info):
        return Normal


class DataCheck(State):

    def __init__(self):
        super(DataCheck, self).__init__()
        self.block = False

    def generate_message(self, user_info):
        message = [
            TextMessage(
                text=
                f"個人資料\n姓名: {user_info['name']}\n梯次: {user_info['session']}\n服勤單位: {user_info['unit']}",
            )
        ]
        return {"user": message, "group": None}

    def next(self, user_input, user_info):
        return Normal


class Normal(State):

    def __init__(self):
        super(Normal, self).__init__()
        # self.block = False

    def generate_message(self, user_info):
        return {"user": None, "group": None}

    def next(self, user_input, user_info):
        if user_input == COMMAND_REQUEST_ABSENCE:
            return Absence
        elif user_input == COMMAND_CANCEL_ABSENCE:
            return CancelTimeoff
        elif user_input == COMMAND_CHECK_NIGHT_TIMEOFF:
            return CheckNightTimeoff
        elif user_input == COMMAND_CHECK_ABSENCE_RECORD:
            return CheckAbsenceRecord
        elif user_input == COMMAND_CHECK_FULL_ABSENCE_RECORD:
            return CheckAllAbsenceRecord
        elif user_input == COMMAND_CHECK_TODAY_ABSENCE:
            return Administration
        elif user_input == COMMAND_REQUEST_TODAY_NIGHT_TIMEOFF:
            today = get_today_date()
            user_info["absence_date"] = today
            user_info["absence_type"] = "夜假"
            if valid_date(user_info["absence_date"],
                          user_info["absence_type"]):
                return NightTimeoff
            else:
                return AbsenceLate
        elif user_input == COMMAND_REQUEST_TOMORROW_TIMEOFF:
            today = get_today_date()
            user_info["absence_date"] = today
            user_info["absence_type"] = "隔天補休"
            if valid_date(user_info["absence_date"],
                          user_info["absence_type"]):
                return OtherTimeoff
            else:
                return AbsenceLate
        elif user_input == COMMAND_REQUEST_OFFICIAL_LEAVE:
            today = get_today_date()
            user_info["absence_date"] = today
            user_info["absence_type"] = "公假"
            if valid_date(user_info["absence_date"],
                          user_info["absence_type"]):
                return OtherTimeoff
            else:
                return AbsenceLate
        elif user_input == COMMAND_CHECK_SELF_INFO:
            return DataCheck
        elif user_input == COMMAND_UPDATE_SELF_INFO:
            return DataError
        else:
            return OutOfScope


class OutOfScope(State):

    def __init__(self):
        super(OutOfScope, self).__init__()
        self.block = False

    def generate_message(self, user_info):
        message = [TextMessage(text="非有效指令，請重新從選單選擇您要執行的操作", )]
        return {"user": message, "group": None}

    def next(self, user_input, user_info):
        return Normal


class Absence(State):

    def __init__(self):
        super(Absence, self).__init__()

    def generate_message(self, user_info):
        option_items = []
        for option in ['夜假', '隔天補休', '公假', '返回']:
            option_items.append(
                QuickReplyItem(
                    action=MessageAction(label=option, text=f"{option}")))
        message = [
            TextMessage(text="請選擇請假種類",
                        quick_reply=QuickReply(items=option_items))
        ]
        return {"user": message, "group": None}

    def next(self, user_input, user_info):
        if user_input == "夜假" or user_input == "隔天補休" or user_input == "公假":
            user_info["absence_type"] = user_input
            return AbsenceDate
        elif user_input == "返回":
            return Normal
        else:
            return OutOfScope


class AbsenceDate(State):

    def __init__(self):
        super(AbsenceDate, self).__init__()

    def generate_message(self, user_info):
        option_items = []
        options = ['取消']
        text = "請輸入請假日期 (Ex. 1/3)，若要取消請按'取消'。請注意，若您是要請'隔天補休'，請輸入您沒有要回來住替中的日期 (Ex. 如果您在2/5補休，您是2/4晚上不用回來住替中，因此請輸入2/4)"
        for option in options:
            option_items.append(
                QuickReplyItem(
                    action=MessageAction(label=option, text=f"{option}")))
        message = [
            TextMessage(text=text, quick_reply=QuickReply(items=option_items))
        ]
        return {"user": message, "group": None}

    def next(self, user_input, user_info):
        if "取消" in user_input:
            return Normal
        elif user_input == "今天":
            today = get_today_date()
            user_info["absence_date"] = today
            return AbsenceConfirm
        elif user_input == "明天":
            today = get_today_date()
            user_info["absence_date"] = today + timedelta(days=1)
            return AbsenceConfirm
        elif "/" in user_input and len(user_input.split("/")) == 2:
            month, day = user_input.split("/")
            month = month.lstrip("0")
            day = day.lstrip("0")
            if month.isdigit() and day.isdigit():
                try:
                    user_info["absence_date"] = format_datetime(
                        int(month), int(day))
                    return AbsenceConfirm
                except ValueError:
                    return AbsenceDateFormatError
            else:
                return AbsenceDateFormatError
        else:
            return AbsenceDateFormatError


class AbsenceDateFormatError(AbsenceDate):

    def __init__(self):
        super(AbsenceDateFormatError, self).__init__()

    def generate_message(self, user_info):
        option_items = []
        options = ['取消']
        text = "請重新輸入請假日期 (Ex. 1/3)，若要取消請按'取消'。請注意，若您是要請'隔天補休'，請輸入您沒有要回來住替中的日期 (Ex. 如果您在2/5補休，您是2/4晚上不用回來住替中，因此請輸入2/4)"

        for option in options:
            option_items.append(
                QuickReplyItem(
                    action=MessageAction(label=option, text=f"{option}")))
        message = [
            TextMessage(text=text, quick_reply=QuickReply(items=option_items))
        ]
        return {"user": message, "group": None}


class AbsenceConfirm(State):

    def __init__(self):
        super(AbsenceConfirm, self).__init__()

    def generate_message(self, user_info):
        option_items = []
        for option in ['確定', '返回']:
            option_items.append(
                QuickReplyItem(
                    action=MessageAction(label=option, text=f"{option}")))
        message = [
            TextMessage(
                text=
                f"確定於 {user_info['absence_date'].strftime('%Y/%-m/%-d')} 請{user_info['absence_type']}",
                quick_reply=QuickReply(items=option_items))
        ]
        return {"user": message, "group": None}

    def next(self, user_input, user_info):
        if user_input == "確定":
            if valid_date(user_info["absence_date"],
                          user_info["absence_type"]):
                if user_info["absence_type"] == "夜假":
                    return NightTimeoff
                else:
                    return OtherTimeoff
            else:
                return AbsenceLate
        elif user_input == "返回":
            return Normal
        else:
            return OutOfScope


class AbsenceLate(State):

    def __init__(self):
        super(AbsenceLate, self).__init__()
        self.block = False

    def generate_message(self, user_info):
        message = [TextMessage(text=f"此日期已超過可請假時間", )]
        return {"user": message, "group": None}

    def next(self, user_input, user_info):
        return Normal


class OtherTimeoff(State):

    def __init__(self):
        super(OtherTimeoff, self).__init__()
        self.block = False

    def generate_message(self, user_info):
        try:
            absence_record = get_absence_records(
                records_col,
                absence_date=user_info['absence_date'].astimezone(pytz.utc),
                user_id=user_info_to_id(user_info['session'],
                                        user_info['unit'], user_info['name']))
            absence_record = list(absence_record)
            if len(absence_record) > 0:
                user_message = [
                    TextMessage(
                        text=
                        f"您之前已在所選日期請了{absence_record[0]['type']}，若要更改假別，請先將舊的假取消",
                    )
                ]
                group_message = None
            else:
                add_absence_record(
                    records_col,
                    user_info['absence_date'].astimezone(pytz.utc),
                    user_info['absence_type'],
                    user_info_to_id(user_info['session'], user_info['unit'],
                                    user_info['name']))
                user_message = [
                    TextMessage(text=f"已登記您的請假申請，可透過選單查看請假紀錄，記得補休假/公假證明給輔導員", )
                ]
                group_message = [
                    TextMessage(
                        text=
                        f"{user_info['absence_date'].strftime('%Y/%-m/%-d')} [{user_info['session']}梯次］{user_info['name']} {user_info['unit']} {user_info['absence_type']}",
                    )
                ]
        except gspread.exceptions.WorksheetNotFound:
            user_message = [TextMessage(text="您的請假資料尚未登入，請稍後再試", )]
            group_message = None
        return {"user": user_message, "group": group_message}

    def next(self, user_input, user_info):
        return Normal


class NightTimeoff(OtherTimeoff):

    def __init__(self):
        super(NightTimeoff, self).__init__()
        self.block = False

    def get_night_timeoff_amount(self, records):
        available_night_timeoff = []
        for row in records:
            if len(row["使用日期"]) == 0:
                available_night_timeoff.append(row)
        return available_night_timeoff

    def update_nigth_timeoff_sheet(self, worksheet, records, user_info):
        data = []
        for record in records:
            if len(record["使用日期"]) != 0:
                data.append([record["使用日期"]])
            else:
                break
        data.append([user_info['absence_date'].strftime('%Y/%-m/%-d')])
        sorted_data = []
        indexes = []
        for i, date in enumerate(data):
            if is_date_format(date[0]):
                sorted_data.append(date[0])
                indexes.append(i)

        sorted_data.sort(key=lambda row: datetime.strptime(row, '%Y/%m/%d'))

        for i, date in enumerate(sorted_data):
            data[indexes[i]] = [date]
        worksheet.update(f"D2:D{len(data)+2}", data)

    def generate_message(self, user_info):
        try:
            absence_record = get_absence_records(
                records_col,
                absence_date=user_info['absence_date'].astimezone(pytz.utc),
                user_id=user_info_to_id(user_info['session'],
                                        user_info['unit'], user_info['name']))
            absence_record = list(absence_record)
            if len(absence_record) > 0:
                user_message = [
                    TextMessage(
                        text=
                        f"您之前已在所選日期請了{absence_record[0]['type']}，若要更改假別，請先將舊的假取消",
                    )
                ]
                group_message = None
            else:
                night_timeoff_sheet = gc.open_by_key(NIGHT_TIMEOFF_SHEET_KEY)
                night_timeoff_worksheet = night_timeoff_sheet.worksheet(
                    user_info_to_id(user_info['session'], user_info['unit'],
                                    user_info['name']))
                night_timeoff_records = night_timeoff_worksheet.get_all_records(
                )
                available_night_timeoff = self.get_night_timeoff_amount(
                    night_timeoff_records)

                if len(available_night_timeoff) == 0:
                    user_message = [TextMessage(text=f"您目前沒有可用夜假", )]
                    group_message = None
                else:
                    self.update_nigth_timeoff_sheet(night_timeoff_worksheet,
                                                    night_timeoff_records,
                                                    user_info)
                    add_absence_record(
                        records_col,
                        user_info['absence_date'].astimezone(pytz.utc),
                        user_info['absence_type'],
                        user_info_to_id(user_info['session'],
                                        user_info['unit'], user_info['name']))
                    user_message = [
                        TextMessage(text=f"已登記您的請假申請，可透過選單查看請假紀錄", )
                    ]
                    group_message = [
                        TextMessage(
                            text=
                            f"{user_info['absence_date'].strftime('%Y/%-m/%-d')} [{user_info['session']}梯次］{user_info['name']} {user_info['unit']} 夜假",
                        )
                    ]
            return {"user": user_message, "group": group_message}
        except gspread.exceptions.WorksheetNotFound:
            message = [TextMessage(text="您的夜假資料尚未登入，請稍後再試", )]
            return {"user": message, "group": None}

    def next(self, user_input, user_info):
        return Normal


class CheckNightTimeoff(NightTimeoff):

    def __init__(self):
        super(CheckNightTimeoff, self).__init__()
        self.block = False

    def generate_night_timeoff_box(self, night_timeoff):
        deadline = night_timeoff["有效期限"]
        if is_date_format(deadline):
            deadline = datetime.strptime(deadline, '%Y/%m/%d').strftime('%m/%d')
        reason = night_timeoff["核發原因"]
        use_date = night_timeoff["使用日期"]
        if is_date_format(use_date):
            use_date = datetime.strptime(use_date, '%Y/%m/%d').strftime('%m/%d')
        elif len(use_date) == 0:
            use_date = " "
        return FlexBox(layout="baseline",
                       spacing="sm",
                       contents=[
                           FlexText(text=reason,
                                    flex=6,
                                    size="sm",
                                    wrap=True),
                           FlexText(text=deadline,
                                    flex=2,
                                    size="sm",
                                    wrap=True),
                            FlexText(text=use_date,
                                    flex=2,
                                    size="sm",
                                    wrap=True)
                       ])

    def generate_message(self, user_info):
        try:
            night_timeoff_sheet = gc.open_by_key(NIGHT_TIMEOFF_SHEET_KEY)
            worksheet = night_timeoff_sheet.worksheet(
                user_info_to_id(user_info['session'], user_info['unit'],
                                user_info['name']))
            night_timeoff_records = worksheet.get_all_records()
            available_night_timeoff = self.get_night_timeoff_amount(
                night_timeoff_records)
            flex_message = copy.deepcopy(night_timeoff_template)
            flex_message.body.contents[0].text += str(
                len(available_night_timeoff))
            for record in night_timeoff_records:
                flex_message.body.contents[1].contents.append(
                    self.generate_night_timeoff_box(record))
            flex_message.footer.contents[0].action.uri += str(worksheet.id)
            message = [FlexMessage(alt_text="夜假", contents=flex_message)]
        except gspread.exceptions.WorksheetNotFound:
            message = [TextMessage(text="您的夜假資料尚未登入，請稍後再試", )]
        return {"user": message, "group": None}

    def next(self, user_input, user_info):
        return Normal


class CheckAbsenceRecord(State):

    def __init__(self):
        super(CheckAbsenceRecord, self).__init__()
        self.block = False

    def generate_absence_record_box(self, date, absence_type):
        return FlexBox(layout="baseline",
                       spacing="sm",
                       contents=[
                           FlexText(text=date,
                                    flex=3,
                                    size="sm"),
                           FlexText(text=absence_type,
                                    flex=3,
                                    size="sm",
                                    wrap=True)
                       ])

    def generate_message(self, user_info):
        records = get_absence_records(
            records_col,
            user_id=user_info_to_id(user_info['session'], user_info['unit'],
                                    user_info['name'])).sort("date",
                                                             -1).limit(5)
        flex_message = copy.deepcopy(absence_record_template)
        for record in list(records)[::-1]:
            date = pytz.utc.localize(
                record["date"]).astimezone(taipei_timezone)
            flex_message.body.contents[1].contents.append(
                self.generate_absence_record_box(date.strftime('%Y/%-m/%-d'),
                                                 record["type"]))

        message = [FlexMessage(alt_text="請假紀錄", contents=flex_message)]
        return {"user": message, "group": None}

    def next(self, user_input, user_info):
        return Normal


class CheckAllAbsenceRecord(CheckAbsenceRecord):

    def __init__(self):
        super(CheckAllAbsenceRecord, self).__init__()
        self.block = False

    def generate_message(self, user_info):
        records = get_absence_records(
            records_col,
            user_id=user_info_to_id(user_info['session'], user_info['unit'],
                                    user_info['name'])).sort("date", -1)
        flex_message = copy.deepcopy(all_absence_record_template)
        for record in list(records)[::-1]:
            date = pytz.utc.localize(
                record["date"]).astimezone(taipei_timezone)
            flex_message.body.contents[1].contents.append(
                self.generate_absence_record_box(date.strftime('%Y/%-m/%-d'),
                                                 record["type"]))
        message = [FlexMessage(alt_text="完整請假紀錄", contents=flex_message)]
        return {"user": message, "group": None}


class Administration(State):

    def __init__(self):
        super(Administration, self).__init__()
        self.block = False

    def generate_today_absence_box(self, session, unit, name, absence_type):
        return FlexBox(layout="baseline",
                       spacing="sm",
                       contents=[
                           FlexText(text=session,
                                    flex=3,
                                    size="sm",
                                    color="#666666"),
                           FlexText(text=unit,
                                    flex=3,
                                    size="sm",
                                    color="#666666"),
                           FlexText(
                               text=name,
                               flex=3,
                               size="sm",
                               color="#666666",
                           ),
                           FlexText(text=absence_type,
                                    flex=3,
                                    size="sm",
                                    color="#666666",
                                    wrap=True),
                       ])

    def generate_message(self, user_info):
        records = get_absence_records(records_col,
                                      absence_date=get_today_date().astimezone(
                                          pytz.utc))
        records = list(records)
        records.sort(key=lambda r: r["userId"])
        flex_message = copy.deepcopy(today_absence_template)
        for record in records:
            session, unit, name = user_id_to_info(record["userId"])

            flex_message.body.contents[1].contents.append(
                self.generate_today_absence_box(session, unit, name,
                                                record["type"]))
        flex_message.body.contents[0].text += str(len(records)) + "人"
        message = [FlexMessage(alt_text="今晚請假役男", contents=flex_message)]

        return {"user": message, "group": None}

    def next(self, user_input, user_info):
        return Normal


class CancelTimeoff(State):

    def __init__(self):
        super(CancelTimeoff, self).__init__()

    def generate_message(self, user_info):
        date_condition = {
            "$gte": get_valid_date().astimezone(pytz.utc),
        }
        records = get_absence_records(
            records_col,
            absence_date=date_condition,
            user_id=user_info_to_id(user_info['session'], user_info['unit'],
                                    user_info['name']))
        timeoff = []
        records = list(records)
        records.sort(key=lambda x: x["date"])
        for record in records:
            date = pytz.utc.localize(
                record["date"]).astimezone(taipei_timezone)
            timeoff.append(f"{date.strftime('%Y/%-m/%-d')} {record['type']}")
        timeoff.append("返回")
        option_items = []
        for option in timeoff:
            option_items.append(
                QuickReplyItem(
                    action=MessageAction(label=option, text=f"{option}")))
        message = [
            TextMessage(text=f"請選擇您要取消的假，若要返回請按返回",
                        quick_reply=QuickReply(items=option_items))
        ]
        return {"user": message, "group": None}

    def next(self, user_input, user_info):
        if "返回" in user_input:
            return Normal
        else:
            try:
                date, absence_type = user_input.split()
                user_info['absence_date'] = taipei_timezone.localize(
                    datetime.strptime(date, '%Y/%m/%d'))
                user_info['absence_type'] = absence_type
                return FinishCancelTimeoff
            except:
                return OutOfScope


class FinishCancelTimeoff(State):

    def __init__(self):
        super(FinishCancelTimeoff, self).__init__()
        self.block = False

    def generate_message(self, user_info):
        fail = False
        delete_absence_record(
            records_col,
            absence_date=user_info["absence_date"].astimezone(pytz.utc),
            absence_type=user_info["absence_type"],
            user_id=user_info_to_id(user_info['session'], user_info['unit'],
                                    user_info['name']))

        date = user_info['absence_date'].strftime('%Y/%-m/%-d')
        if user_info["absence_type"] == "夜假":
            night_timeoff_sheet = gc.open_by_key(NIGHT_TIMEOFF_SHEET_KEY)
            night_timeoff_worksheet = night_timeoff_sheet.worksheet(
                user_info_to_id(user_info['session'], user_info['unit'],
                                user_info['name']))
            night_timeoff_records = night_timeoff_worksheet.get_all_records()

            data = []
            for record in night_timeoff_records:
                if len(record["使用日期"]) != 0:
                    data.append([record["使用日期"]])
                else:
                    break

            previous = -1
            for i in range(len(data)):
                if previous == -1 and data[i][0] == date:
                    data[i] = [""]
                    previous = i
                elif previous != -1 and is_date_format(data[i][0]):
                    data[previous] = data[i]
                    data[i] = [""]
                    previous = i
            night_timeoff_worksheet.update(f"D2:D{len(data)+1}", data)

        if fail == False:
            user_message = [TextMessage(text=f"已幫您取消該假，可透過選單查看請假紀錄", )]
            group_message = [
                TextMessage(
                    text=
                    f"== 取消請假 == {user_info['absence_date'].strftime('%Y/%-m/%-d')} [{user_info['session']}梯次］{user_info['name']} {user_info['unit']} 夜假",
                )
            ]
            return {"user": user_message, "group": group_message}
        else:
            message = [
                TextMessage(text="取消失敗，請重新操作 (請注意，若超過晚上8點，則不能取消今日之夜假或明日之補休)", )
            ]
            return {"user": message, "group": None}

    def next(self, user_input, user_info):
        return Normal


def get_folder_id(folder_name):
    folder_info_from_db = folders_col.find_one({"_id": folder_name})
    if folder_info_from_db:
        return folder_info_from_db["folder_id"]
    else:
        return None


class UploadProof(State):

    def __init__(self):
        super(UploadProof, self).__init__()

    def generate_message(self, user_info):
        folder_name = user_info_to_id(user_info['session'], user_info['unit'],
                                      user_info['name'])
        folder_id = get_folder_id(folder_name)
        if folder_id:
            message = [
                TextMessage(
                    text=
                    f"請將證明上傳至 https://drive.google.com/drive/folders/{folder_id}?usp=sharing\n\n檔名請標明日期與假別 (Ex. 2025-02-03-補休.jpg)",
                )
            ]
        else:
            message = [TextMessage(text=f"您的上傳資料夾尚未建立，請稍後再試", )]
        return {"user": message, "group": None}

    def next(self, user_input, user_info):
        return Normal()
