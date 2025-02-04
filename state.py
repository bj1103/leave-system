from linebot.v3.messaging import (
    TextMessage,
    FlexBox,
    FlexText,
)
from linebot.v3.messaging import (QuickReply, QuickReplyItem, MessageAction,
                                  FlexMessage)
from template import night_timeoff_template, absence_record_template, today_absence_template
from datetime import datetime, timedelta
import pytz
import pandas as pd
import gspread
import copy
import json
import os

NIGHT_TIMEOFF_SHEET_KEY = "10o1RavT1RGKFccEdukG1HsEgD3FPOBOPMB6fQqTc_wI"
ABSENCE_RECORD_SHEET_KEY = "1TxClL3L0pDQAIoIidgJh7SP-BF4GaBD6KKfVKw0CLZQ"
service_account_info = json.loads(os.getenv('GOOGLE_SERVICE_ACCOUNT_JSON'))
service_account_info['private_key'] = service_account_info[
    'private_key'].replace("\\n", "\n")
gc = gspread.service_account_from_dict(service_account_info)
taipei_timezone = pytz.timezone('Asia/Taipei')

COMMAND_REQUEST_ABSENCE = "== 請假 =="
COMMAND_CANCEL_ABSENCE = "== 取消請假 =="
COMMAND_CHECK_NIGHT_TIMEOFF = "== 查看夜假 =="
COMMAND_CHECK_ABSENCE_RECORD = "== 查看請假紀錄 =="
COMMAND_CHECK_TODAY_ABSENCE = "== 今日請假役男 =="
KEYWORD = {
    COMMAND_REQUEST_ABSENCE, COMMAND_CANCEL_ABSENCE,
    COMMAND_CHECK_NIGHT_TIMEOFF, COMMAND_CHECK_ABSENCE_RECORD,
    COMMAND_CHECK_TODAY_ABSENCE
}


def format_datetime(month, day):
    today = datetime.now(taipei_timezone)
    absence_day = today.replace(month=month, day=day)
    if absence_day < today:
        absence_day = absence_day.replace(year=absence_day.year + 1)
    return absence_day


def valid_date(absence_date, absence_type):
    today = datetime.now(taipei_timezone)
    over_8_pm = today.hour > 20
    today = today.replace(hour=0, minute=0, second=0, microsecond=0)
    if absence_type == "補休" and absence_date > today + timedelta(
            days=1) * (over_8_pm):
        return True
    elif absence_type == "夜假" and absence_date >= today + timedelta(
            days=1) * (over_8_pm):
        return True
    elif absence_type == "公差" and absence_date >= today:
        return True
    else:
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
            return DataFinish
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


class DataFinish(State):

    def __init__(self):
        super(DataFinish, self).__init__()
        self.block = False

    def generate_message(self, user_info):
        message = [TextMessage(text=f"已記錄您的資料，現在您可以開啟選單來使用本系統", )]
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
        elif user_input == COMMAND_CHECK_TODAY_ABSENCE:
            return Administration
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
        for option in ['夜假', '補休', '公差', '返回']:
            option_items.append(
                QuickReplyItem(
                    action=MessageAction(label=option, text=f"{option}")))
        message = [
            TextMessage(text="請選擇請假種類",
                        quick_reply=QuickReply(items=option_items))
        ]
        return {"user": message, "group": None}

    def next(self, user_input, user_info):
        if user_input == "夜假" or user_input == "補休" or user_input == "公差":
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
        if user_info['absence_type'] == '補休':
            options = ['明天', '取消']
            text = "請按'明天'或輸入請假日期 (Ex. 1/3)，若要取消請按'取消'"
        else:
            options = ['今天', '取消']
            text = "請按'今天'或輸入請假日期 (Ex. 1/3)，若要取消請按'取消'"

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
            today = datetime.now(taipei_timezone).replace(hour=0,
                                                          minute=0,
                                                          second=0,
                                                          microsecond=0)
            user_info["absence_date"] = today
            return AbsenceConfirm
        elif user_input == "明天":
            today = datetime.now(taipei_timezone).replace(hour=0,
                                                          minute=0,
                                                          second=0,
                                                          microsecond=0)
            user_info["absence_date"] = today + timedelta(days=1)
            return AbsenceConfirm
        elif "/" in user_input and len(user_input.split("/")) == 2:
            month, day = user_input.split("/")
            month = month.lstrip("0")
            day = day.lstrip("0")
            if month.isdigit() and day.isdigit():
                user_info["absence_date"] = format_datetime(
                    int(month), int(day))
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
        if user_info['absence_type'] == '補休':
            options = ['明天', '取消']
            text = "請重新輸入請假日期 (Ex. 1/3)，或按'明天'，若要取消請按'取消'"
        else:
            options = ['今天', '明天', '取消']
            text = "請重新輸入請假日期 (Ex. 1/3)，或按'今天'、'明天'，若要取消請按'取消'"

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

    def update_absence_record(self, df, worksheet, user_info):
        idx = len(df)
        cells = worksheet.range(f"A{idx+2}:B{idx+2}")
        cells[0].value = user_info['absence_date'].strftime('%Y/%-m/%-d')
        cells[1].value = user_info['absence_type']
        worksheet.update_cells(cells)

    def generate_message(self, user_info):
        try:
            absence_record_sheet = gc.open_by_key(ABSENCE_RECORD_SHEET_KEY)
            worksheet = absence_record_sheet.worksheet(
                f"{user_info['session']}T{user_info['name']}")
            df = pd.DataFrame(worksheet.get_all_records())
            self.update_absence_record(df, worksheet, user_info)

            user_message = [TextMessage(text=f"已登記您的請假申請，請至群組確認請假資訊", )]
            group_message = [
                TextMessage(
                    text=
                    f"{user_info['absence_date'].strftime('%Y/%-m/%-d')} [{user_info['session']}梯次］{user_info['name']} {user_info['unit']} {user_info['absence_type']}",
                )
            ]
        except KeyError:
            user_message = [TextMessage(text="您的請假資料尚未登入，請稍後再試", )]
            group_message = None
        return {"user": user_message, "group": group_message}

    def next(self, user_input, user_info):
        return Normal


class NightTimeoff(OtherTimeoff):

    def __init__(self):
        super(NightTimeoff, self).__init__()
        self.block = False

    def get_night_timeoff_amount(self, df):
        available_night_timeoff = []
        for index, row in df.iterrows():
            if len(row["請假日期"]) == 0 and len(row["核發日期"]) != 0:
                available_night_timeoff.append(row["有效期限"])
        return available_night_timeoff

    def update_nigth_timeoff_sheet(self, df, worksheet, user_info):
        target_row_id = -1
        deadline = datetime.max
        for row_id, row in df.iterrows():
            if len(row["請假日期"]) == 0:
                year, month, day = [int(x) for x in row["有效期限"].split("/")]
                row_deadline = datetime(year=year, month=month, day=day)
                if row_deadline < deadline:
                    deadline = row_deadline
                    target_row_id = row_id
        worksheet.update_cell(
            row=target_row_id + 2,
            col=5,
            value=user_info["absence_date"].strftime('%Y/%-m/%-d'))

    def generate_message(self, user_info):
        try:
            night_timeoff_sheet = gc.open_by_key(NIGHT_TIMEOFF_SHEET_KEY)
            worksheet = night_timeoff_sheet.worksheet(
                f"{user_info['session']}T{user_info['name']}")
            df = pd.DataFrame(worksheet.get_all_records())
            available_night_timeoff = self.get_night_timeoff_amount(df)

            if len(available_night_timeoff) == 0:
                message = [TextMessage(text=f"您目前沒有可用夜假", )]
                return {"user": message, "group": None}
            else:
                self.update_nigth_timeoff_sheet(df, worksheet, user_info)

                absence_record_sheet = gc.open_by_key(ABSENCE_RECORD_SHEET_KEY)
                worksheet = absence_record_sheet.worksheet(
                    f"{user_info['session']}T{user_info['name']}")
                df = pd.DataFrame(worksheet.get_all_records())
                self.update_absence_record(df, worksheet, user_info)

                user_message = [TextMessage(text=f"已登記您的請假申請，請至群組確認請假資訊", )]
                group_message = [
                    TextMessage(
                        text=
                        f"{user_info['absence_date'].strftime('%Y/%-m/%-d')} [{user_info['session']}梯次］{user_info['name']} {user_info['unit']} 夜假",
                    )
                ]
                return {"user": user_message, "group": group_message}
        except KeyError:
            message = [TextMessage(text="您的夜假資料尚未登入，請稍後再試", )]
            return {"user": message, "group": None}

    def next(self, user_input, user_info):
        return Normal


class CheckNightTimeoff(NightTimeoff):

    def __init__(self):
        super(CheckNightTimeoff, self).__init__()
        self.block = False

    def generate_night_timeoff_box(self, id, deadline):
        year, month, day = [int(x) for x in deadline.split("/")]
        return FlexBox(layout="baseline",
                       spacing="sm",
                       contents=[
                           FlexText(text=f"{id}.",
                                    flex=1,
                                    size="sm",
                                    color="#aaaaaa"),
                           FlexText(text=f"{year}/{month:02d}/{day:02d}",
                                    flex=3,
                                    size="sm",
                                    color="#666666",
                                    wrap=True)
                       ])

    def generate_message(self, user_info):
        try:
            night_timeoff_sheet = gc.open_by_key(NIGHT_TIMEOFF_SHEET_KEY)
            worksheet = night_timeoff_sheet.worksheet(
                f"{user_info['session']}T{user_info['name']}")
            df = pd.DataFrame(worksheet.get_all_records())
            available_night_timeoff = self.get_night_timeoff_amount(df)
            flex_message = copy.deepcopy(night_timeoff_template)
            flex_message.body.contents[0].text += str(
                len(available_night_timeoff))
            for i, nigth_timeoff in enumerate(available_night_timeoff):
                flex_message.body.contents[1].contents.append(
                    self.generate_night_timeoff_box(i + 1, nigth_timeoff))
            message = [FlexMessage(alt_text="夜假", contents=flex_message)]
        except KeyError:
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
                                    size="sm",
                                    color="#aaaaaa"),
                           FlexText(text=absence_type,
                                    flex=3,
                                    size="sm",
                                    color="#666666",
                                    wrap=True)
                       ])

    def generate_message(self, user_info):
        try:
            absence_record_sheet = gc.open_by_key(ABSENCE_RECORD_SHEET_KEY)
            worksheet = absence_record_sheet.worksheet(
                f"{user_info['session']}T{user_info['name']}")
            df = pd.DataFrame(worksheet.get_all_records())

            flex_message = copy.deepcopy(absence_record_template)
            for i, row in df.tail(5).iterrows():
                flex_message.body.contents[1].contents.append(
                    self.generate_absence_record_box(row["日期"], row["假別"]))
            message = [FlexMessage(alt_text="夜假", contents=flex_message)]
        except KeyError:
            message = [TextMessage(text="您的請假資料尚未登入，請稍後再試", )]

        return {"user": message, "group": None}

    def next(self, user_input, user_info):
        return Normal


class Administration(State):

    def __init__(self):
        super(Administration, self).__init__()
        self.block = False

    def generate_today_absence_box(self, user, absence_type):
        session, name = user.split("T")
        return FlexBox(layout="baseline",
                       spacing="sm",
                       contents=[
                           FlexText(text=session,
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
        absence_record_sheet = gc.open_by_key(ABSENCE_RECORD_SHEET_KEY)
        worksheet = absence_record_sheet.worksheet("今日請假")
        df = pd.DataFrame(worksheet.get_all_records())

        flex_message = copy.deepcopy(today_absence_template)
        for i, row in df.iterrows():
            flex_message.body.contents[1].contents.append(
                self.generate_today_absence_box(row["請假人"], row["假別"]))
        flex_message.body.contents[0].text += str(len(df)) + "人"
        message = [FlexMessage(alt_text="夜假", contents=flex_message)]

        return {"user": message, "group": None}

    def next(self, user_input, user_info):
        return Normal


class CancelTimeoff(State):

    def __init__(self):
        super(CancelTimeoff, self).__init__()

    def get_future_timeoff(self, df):
        out = []
        today = datetime.now(taipei_timezone)
        today = today.replace(hour=0, minute=0, second=0, microsecond=0)
        for i, row in df.iterrows():
            year, month, day = [int(x) for x in row["日期"].split("/")]
            row_date = taipei_timezone.localize(
                datetime(year=year, month=month, day=day))
            if valid_date(row_date, row["假別"]):
                out.append(f"{row['日期']} {row['假別']}")
        return out

    def generate_message(self, user_info):
        absence_record_sheet = gc.open_by_key(ABSENCE_RECORD_SHEET_KEY)
        worksheet = absence_record_sheet.worksheet(
            f"{user_info['session']}T{user_info['name']}")
        df = pd.DataFrame(worksheet.get_all_records())
        timeoff = self.get_future_timeoff(df) + ["返回"]
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
                year, month, day = [int(x) for x in date.split("/")]
                user_info['absence_date'] = datetime(year=year,
                                                     month=month,
                                                     day=day)
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
        absence_record_sheet = gc.open_by_key(ABSENCE_RECORD_SHEET_KEY)
        worksheet = absence_record_sheet.worksheet(
            f"{user_info['session']}T{user_info['name']}")
        df = pd.DataFrame(worksheet.get_all_records())
        date = user_info['absence_date'].strftime('%Y/%-m/%-d')
        idx = df[(df['日期'] == date)
                 & (df['假別'] == user_info['absence_type'])].index
        if len(idx):
            worksheet.delete_rows(int(idx[-1]) + 2)
        else:
            fail = True

        if user_info["absence_type"] == "夜假":
            night_timeoff_sheet = gc.open_by_key(NIGHT_TIMEOFF_SHEET_KEY)
            worksheet = night_timeoff_sheet.worksheet(
                f"{user_info['session']}T{user_info['name']}")
            df = pd.DataFrame(worksheet.get_all_records())
            idx = df[(df['請假日期'] == date)].index
            if len(idx):
                idx_int = int(idx[-1]) + 2
                worksheet.update_cell(
                    row=idx_int,
                    col=5,
                    value=
                    f'=IF(D{idx_int}="", "", IF(D{idx_int}<TODAY(), "expired", ""))'
                )
            else:
                fail = True

        if fail == False:
            user_message = [TextMessage(text=f"已幫您取消申請，請至群組確認取消資訊", )]
            group_message = [
                TextMessage(
                    text=
                    f"== 取消請假 == {user_info['absence_date'].strftime('%Y/%-m/%-d')} [{user_info['session']}梯次］{user_info['name']} {user_info['unit']} 夜假",
                )
            ]
            return {"user": user_message, "group": group_message}
        else:
            message = user_message = [
                TextMessage(text="取消失敗，請重新操作 (請注意，若超過晚上8點，則不能取消今日之夜假或明日之補休)", )
            ]
            return {"user": message, "group": None}

    def next(self, user_input, user_info):
        return Normal
