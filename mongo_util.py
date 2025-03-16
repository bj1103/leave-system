def count_absence_record(record_col,
                         absence_date=None,
                         absence_type=None,
                         user_id=None):
    query = {}
    if absence_date:
        query["date"] = absence_date
    if absence_type:
        query["type"] = absence_type
    if user_id:
        query["userId"] = user_id
    return record_col.count_documents(query)


def get_absence_records(record_col,
                        absence_date=None,
                        absence_type=None,
                        user_id=None):
    query = {}
    if absence_date:
        query["date"] = absence_date
    if absence_type:
        query["type"] = absence_type
    if user_id:
        query["userId"] = user_id
    return record_col.find(query)


def add_absence_record(record_col, absence_date, absence_type, user_id):
    data = {"userId": user_id, "type": absence_type, "date": absence_date}
    return record_col.insert_one(data)


def delete_absence_record(record_col, absence_date, absence_type, user_id):
    data = {"userId": user_id, "type": absence_type, "date": absence_date}
    return record_col.delete_one(data)


def get_absence_users(record_col, absence_date=None, absence_type=None):
    match_condition = {}
    if absence_date:
        match_condition["date"] = absence_date
    if absence_type:
        match_condition["type"] = absence_type

    pipeline = [
        {
            "$match": match_condition
        },
        {
            "$lookup": {
                "from": "user",  # 要 JOIN 的 collection
                "localField": "userId",  # collection_b 的 user_id
                "foreignField": "_id",  # collection_a 的 user_id
                "as": "user_info"  # 輸出結果存放在 user_info
            }
        },
        {
            "$unwind": "$user_info"  # 展開 user_info 陣列
        },
        {
            "$project": {  # 選擇要顯示的欄位
                "_id": 0,
                "user_id": 1,
                "name": "$user_info.name",
                "unit": "$user_info.unit",
                "session": "$user_info.session",
                "date": 1,
                "type": 1
            }
        }
    ]
    return record_col.aggregate(pipeline)
