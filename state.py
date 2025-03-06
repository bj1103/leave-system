from linebot.v3.messaging import (TextMessage, FlexBox, FlexText, QuickReply,
                                  QuickReplyItem, MessageAction, FlexMessage)
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

COMMAND_REQUEST_ABSENCE = "== 其他請假 =="
COMMAND_CANCEL_ABSENCE = "== 取消請假 =="
COMMAND_CHECK_NIGHT_TIMEOFF = "== 查看剩餘夜假 =="
COMMAND_CHECK_ABSENCE_RECORD = "== 請假紀錄 =="
COMMAND_CHECK_TODAY_ABSENCE = "== 今日請假役男 =="
COMMAND_REQUEST_TODAY_NIGHT_TIMEOFF = "== 請今晚夜假 =="
COMMAND_REQUEST_TOMORROW_TIMEOFF = "== 請隔天補休 =="

KEYWORD = {
    COMMAND_REQUEST_ABSENCE, COMMAND_CANCEL_ABSENCE,
    COMMAND_CHECK_NIGHT_TIMEOFF, COMMAND_CHECK_ABSENCE_RECORD,
    COMMAND_CHECK_TODAY_ABSENCE, COMMAND_REQUEST_TODAY_NIGHT_TIMEOFF,
    COMMAND_REQUEST_TOMORROW_TIMEOFF
}

night_timeoff_headers = ["核發原因", "核發日期", "有效期限", "使用日期"]
absence_headers = ["請假日期", "假別"]

def format_datetime(month, day):
    today = datetime.now(taipei_timezone)
    absence_day = today.replace(month=month, day=day)
    if absence_day < today:
        absence_day = absence_day.replace(year=absence_day.year + 1)
    return absence_day


def valid_date(absence_date, absence_type):
    today = datetime.now(taipei_timezone)
    overtime = False
    if today.weekday() == 6:
        overtime = today.hour > 22
    else:
        overtime = today.hour > 20

    today = today.replace(hour=0, minute=0, second=0, microsecond=0)
    if (absence_type == "夜假" or absence_type == "隔天補休"
        ) and absence_date >= today + timedelta(days=1) * (overtime):
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
        elif user_input == COMMAND_REQUEST_TODAY_NIGHT_TIMEOFF:
            today = datetime.now(taipei_timezone).replace(hour=0,
                                                          minute=0,
                                                          second=0,
                                                          microsecond=0)
            user_info["absence_date"] = today
            user_info["absence_type"] = "夜假"
            if valid_date(user_info["absence_date"],
                          user_info["absence_type"]):
                return NightTimeoff
            else:
                return AbsenceLate
        elif user_input == COMMAND_REQUEST_TOMORROW_TIMEOFF:
            today = datetime.now(taipei_timezone).replace(hour=0,
                                                          minute=0,
                                                          second=0,
                                                          microsecond=0)
            user_info["absence_date"] = today
            user_info["absence_type"] = "隔天補休"
            if valid_date(user_info["absence_date"],
                          user_info["absence_type"]):
                return OtherTimeoff
            else:
                return AbsenceLate
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
        for option in ['夜假', '隔天補休', '公差', '返回']:
            option_items.append(
                QuickReplyItem(
                    action=MessageAction(label=option, text=f"{option}")))
        message = [
            TextMessage(text="請選擇請假種類",
                        quick_reply=QuickReply(items=option_items))
        ]
        return {"user": message, "group": None}

    def next(self, user_input, user_info):
        if user_input == "夜假" or user_input == "隔天補休" or user_input == "公差":
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

    def update_absence_record(self, worksheet, user_info):
        length = 0
        for record in worksheet.get_all_records(expected_headers=absence_headers):
            if len(record["請假日期"]) != 0:
                length += 1
            else:
                break
        if length == 0:
            data = []
        else:
            data = worksheet.get(f"A2:B{length+1}")
        data.append([
            user_info['absence_date'].strftime('%Y/%-m/%-d'),
            user_info['absence_type']  
        ])
        sorted_data = sorted(
            data, key=lambda row: datetime.strptime(row[0], '%Y/%m/%d'))
        worksheet.update(f"A2:B{length+2}", sorted_data)

    def generate_message(self, user_info):
        try:
            absence_record_sheet = gc.open_by_key(ABSENCE_RECORD_SHEET_KEY)
            worksheet = absence_record_sheet.worksheet(
                f"{user_info['session']}T_{user_info['unit']}_{user_info['name']}")
            self.update_absence_record(worksheet, user_info)

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

    def get_night_timeoff_amount(self, worksheet):
        available_night_timeoff = []
        for row in worksheet.get_all_records(expected_headers=night_timeoff_headers):
            if len(row["使用日期"]) == 0 and len(row["核發日期"]) != 0:
                available_night_timeoff.append(row["有效期限"])
        return available_night_timeoff

    def update_nigth_timeoff_sheet(self, worksheet, user_info):
        length = 0
        for record in worksheet.get_all_records(expected_headers=night_timeoff_headers):
            if len(record["使用日期"]) != 0:
                length += 1
            else:
                break
        if length == 0:
            data = []
        else:
            data = worksheet.get(f"D2:D{length+1}")
        data.append([
            user_info['absence_date'].strftime('%Y/%-m/%-d')
        ])
        sorted_data = []
        indexes = []
        for i, date in enumerate(data):
            if date[0] != "已過期":
                sorted_data.append(date[0])
                indexes.append(i)

        sorted_data.sort(key=lambda row: datetime.strptime(row, '%Y/%m/%d'))
        
        for i, date in enumerate(sorted_data):
            data[indexes[i]] = [date]
        worksheet.update(f"D2:D{length+2}", data)



        # target_row_id = -1
        # deadline = datetime.max
        # for row_id, row in enumerate(worksheet.get_all_records(expected_headers=night_timeoff_headers)):
        #     if len(row["使用日期"]) == 0 and len(row["核發日期"]) != 0:
        #         year, month, day = [int(x) for x in row["有效期限"].split("/")]
        #         row_deadline = datetime(year=year, month=month, day=day)
        #         if row_deadline < deadline:
        #             deadline = row_deadline
        #             target_row_id = row_id
        # worksheet.update_cell(
        #     row=target_row_id + 2,
        #     col=4,
        #     value=user_info["absence_date"].strftime('%Y/%-m/%-d'))

    def generate_message(self, user_info):
        try:
            night_timeoff_sheet = gc.open_by_key(NIGHT_TIMEOFF_SHEET_KEY)
            worksheet = night_timeoff_sheet.worksheet(
                f"{user_info['session']}T_{user_info['unit']}_{user_info['name']}")
            available_night_timeoff = self.get_night_timeoff_amount(worksheet)

            if len(available_night_timeoff) == 0:
                message = [TextMessage(text=f"您目前沒有可用夜假", )]
                return {"user": message, "group": None}
            else:
                self.update_nigth_timeoff_sheet(worksheet, user_info)
                
                absence_record_sheet = gc.open_by_key(ABSENCE_RECORD_SHEET_KEY)
                worksheet = absence_record_sheet.worksheet(
                    f"{user_info['session']}T_{user_info['unit']}_{user_info['name']}")
                self.update_absence_record(worksheet, user_info)

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
        # try:
        night_timeoff_sheet = gc.open_by_key(NIGHT_TIMEOFF_SHEET_KEY)
        worksheet = night_timeoff_sheet.worksheet(
            f"{user_info['session']}T_{user_info['unit']}_{user_info['name']}")
        available_night_timeoff = self.get_night_timeoff_amount(worksheet)
        flex_message = copy.deepcopy(night_timeoff_template)
        flex_message.body.contents[0].text += str(
            len(available_night_timeoff))
        for i, nigth_timeoff in enumerate(available_night_timeoff):
            flex_message.body.contents[1].contents.append(
                self.generate_night_timeoff_box(i + 1, nigth_timeoff))
        flex_message.footer.contents[0].action.uri += str(worksheet.id)
        message = [FlexMessage(alt_text="夜假", contents=flex_message)]
        # except KeyError:
        #     message = [TextMessage(text="您的夜假資料尚未登入，請稍後再試", )]
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
                f"{user_info['session']}T_{user_info['unit']}_{user_info['name']}")

            flex_message = copy.deepcopy(absence_record_template)
            count = 0
            records = worksheet.get_all_records(expected_headers=absence_headers)
            for record in records[::-1]:
                if len(record["請假日期"]) == 0:
                    continue
                else:
                    flex_message.body.contents[1].contents.insert(
                        0,
                        self.generate_absence_record_box(
                            record["請假日期"], record["假別"]))
                    count += 1
                if count == 5:
                    break
            flex_message.footer.contents[0].action.uri += str(worksheet.id)
            message = [FlexMessage(alt_text="請假紀錄", contents=flex_message)]
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
        session, unit, name = user.split("_")
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
        absence_record_sheet = gc.open_by_key(ABSENCE_RECORD_SHEET_KEY)
        worksheet = absence_record_sheet.worksheet("今日請假")
        df = pd.DataFrame(worksheet.get_all_records())

        flex_message = copy.deepcopy(today_absence_template)
        for i, row in df.iterrows():
            flex_message.body.contents[1].contents.append(
                self.generate_today_absence_box(row["請假人"], row["假別"]))
        flex_message.body.contents[0].text += str(len(df)) + "人"
        message = [FlexMessage(alt_text="今晚請假役男", contents=flex_message)]

        return {"user": message, "group": None}

    def next(self, user_input, user_info):
        return Normal


class CancelTimeoff(State):

    def __init__(self):
        super(CancelTimeoff, self).__init__()

    def get_future_timeoff(self, worksheet):
        out = []
        today = datetime.now(taipei_timezone)
        today = today.replace(hour=0, minute=0, second=0, microsecond=0)
        for row in worksheet.get_all_records(expected_headers=absence_headers):
            if len(row["請假日期"]):
                year, month, day = [int(x) for x in row["請假日期"].split("/")]
                row_date = taipei_timezone.localize(
                    datetime(year=year, month=month, day=day))
                if valid_date(row_date, row["假別"]):
                    out.append(f"{row['請假日期']} {row['假別']}")
        return out

    def generate_message(self, user_info):
        absence_record_sheet = gc.open_by_key(ABSENCE_RECORD_SHEET_KEY)
        worksheet = absence_record_sheet.worksheet(
            f"{user_info['session']}T_{user_info['unit']}_{user_info['name']}")
        timeoff = self.get_future_timeoff(worksheet) + ["返回"]
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
        absence_worksheet = absence_record_sheet.worksheet(
            f"{user_info['session']}T_{user_info['unit']}_{user_info['name']}")
        df = pd.DataFrame(absence_worksheet.get_all_records(expected_headers=absence_headers))
        length = 0
        for _, row in df.iterrows():
            if len(row["請假日期"]) != 0:
                length += 1
            else:
                break
        
        date = user_info['absence_date'].strftime('%Y/%-m/%-d')
        idxs = df[(df['請假日期'] == date)
                  & (df['假別'] == user_info['absence_type'])].index
        if len(idxs):
            idx = int(idxs[-1])
            if idx == length - 1:
                cells = absence_worksheet.range(f"A{idx + 2}:B{idx + 2}")
                cells[0].value = ""
                cells[1].value = ""
                absence_worksheet.update_cells(cells)
            else:
                data = absence_worksheet.get(f"A{idx+3}:B{length+1}")
                absence_worksheet.update(f"A{idx+2}:B{length}", data)
                absence_worksheet.batch_clear([f"A{length+1}:B{length+1}"])
        else:
            fail = True

        if user_info["absence_type"] == "夜假":
            length = 0
            night_timeoff_sheet = gc.open_by_key(NIGHT_TIMEOFF_SHEET_KEY)
            night_timeoff_worksheet = night_timeoff_sheet.worksheet(
                f"{user_info['session']}T_{user_info['unit']}_{user_info['name']}")
            for record in night_timeoff_worksheet.get_all_records(expected_headers=night_timeoff_headers):
                if len(record["使用日期"]) != 0:
                    length += 1
                else:
                    break

            data = night_timeoff_worksheet.get(f"D2:D{length+1}")

            previous = -1
            for i in range(len(data)):
                if previous == -1 and data[i][0] == date:
                    data[i] = [""]
                    previous = i
                elif previous != -1 and data[i][0] != "已過期":
                    data[previous] = data[i]
                    data[i] = [""]
                    previous = i
            night_timeoff_worksheet.update(f"D2:D{length+1}", data)
            
            # df = pd.DataFrame(worksheet.get_all_records(expected_headers=night_timeoff_headers))
            # idxs = df[(df['使用日期'] == date)].index
            # if len(idxs):
            #     idx = int(idxs[-1]) + 2
            #     worksheet.update_cell(
            #         row=idx,
            #         col=4,
            #         value=""
            #     )
            # else:
            #     fail = True

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
