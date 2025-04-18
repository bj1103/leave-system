from linebot.v3.messaging import (FlexBubble, FlexBox, FlexText, FlexButton,
                                  URIAction, MessageAction)

night_timeoff_template = FlexBubble(
    body=FlexBox(layout="vertical",
                 contents=[
                     FlexText(text="可用夜假總數: ",
                              weight="bold",
                              size="xl",),
                     FlexBox(layout="vertical",
                             margin="lg",
                             spacing="lg",
                             contents=[
                                 FlexBox(layout="baseline",
                                         spacing="sm",
                                         contents=[
                                             FlexText(text="核發原因",
                                                      flex=6,
                                                      size="sm",
                                                      color="#666666",
                                                      weight="bold"),
                                             FlexText(text="到期日",
                                                      flex=2,
                                                      size="sm",
                                                      color="#666666",
                                                      weight="bold"),
                                            FlexText(text="使用日",
                                                      flex=2,
                                                      size="sm",
                                                      color="#666666",
                                                      weight="bold")
                                         ])
                             ])
                 ]),
    footer=FlexBox(
        layout="vertical",
        contents=[
            FlexButton(action=URIAction(
                label="詳細資料",
                uri=
                "https://docs.google.com/spreadsheets/d/10o1RavT1RGKFccEdukG1HsEgD3FPOBOPMB6fQqTc_wI/edit?usp=sharing#gid="
            ))
        ]))

absence_record_template = FlexBubble(
    body=FlexBox(layout="vertical",
                 contents=[
                     FlexText(text="近期5筆請假紀錄",
                              weight="bold",
                              size="xl"),
                     FlexBox(layout="vertical",
                             margin="lg",
                             spacing="sm",
                             contents=[])
                 ]),
    footer=FlexBox(
        layout="vertical",
        contents=[
            FlexButton(
                action=MessageAction(label="完整請假紀錄", text="== 完整請假紀錄 =="))
        ])
    # footer=FlexBox(
    #     layout="vertical",
    #     contents=[
    #         FlexButton(action=URIAction(
    #             label="詳細資料",
    #             uri=
    #             "https://docs.google.com/spreadsheets/d/1TxClL3L0pDQAIoIidgJh7SP-BF4GaBD6KKfVKw0CLZQ/edit?usp=sharing#gid="
    #         ))
    #     ]
    # )
)

all_absence_record_template = FlexBubble(body=FlexBox(
    layout="vertical",
    contents=[
        FlexText(text="所有請假紀錄", weight="bold", size="xl"),
        FlexBox(layout="vertical", margin="lg", spacing="sm", contents=[])
    ]))

today_absence_template = FlexBubble(
    body=FlexBox(layout="vertical",
                 contents=[
                     FlexText(text="今日請假役男: ",
                              weight="bold",
                              size="xl"),
                     FlexBox(layout="vertical",
                             margin="lg",
                             spacing="sm",
                             contents=[
                                 FlexBox(layout="baseline",
                                         spacing="sm",
                                         contents=[
                                             FlexText(text="梯次",
                                                      flex=3,
                                                      size="md",
                                                      weight="bold"),
                                             FlexText(text="單位",
                                                      flex=3,
                                                      size="md",
                                                      weight="bold"),
                                             FlexText(text="姓名",
                                                      flex=3,
                                                      size="md",
                                                      weight="bold"),
                                             FlexText(text="假別",
                                                      flex=3,
                                                      size="md",
                                                      weight="bold")
                                         ])
                             ])
                 ]),
    # footer=FlexBox(
    #     layout="vertical",
    #     contents=[
    #         FlexButton(action=URIAction(
    #             label="詳細資料",
    #             uri=
    #             "https://docs.google.com/spreadsheets/d/1TxClL3L0pDQAIoIidgJh7SP-BF4GaBD6KKfVKw0CLZQ/edit?usp=sharing"
    #         ))
    #     ]
    # )
)
