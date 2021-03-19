import pandas as pd

if __name__ == '__main__':
    data_file = "input.xlsx"
    
    df = pd.read_excel(data_file)
    print(df.head())
    names = df.iloc[:, 0]
    phones = df.iloc[:, 1]
    output_file = "output.xlsx"
    content = "【电院科协面试通知】亲爱的{}同学，感谢您报名电院科协的招新活动，您的面试时间为{}，面试地点在xxxxxxxx，包含1分钟的自我介绍以及4分钟的提问环节。\
        请及时加入QQ群xxxxxxxxxx（面试安排如有变化将在群内通知）。\
        收到请回复【姓名+收到】，如需调整时间请回复【姓名+调整+调整时间】(可在3月24日下午2：00-4:00间调整)，电院科协期待您的加入。"
    recruit_time = "3月24日下午{}: {:02}"
    current_time = 2
    text_content = []
    n = len(names)
    for i in range(n):
        name = names[i]
        current_hour = int(current_time)
        current_minute = int(100 * (current_time - int(current_time)))
        current_time += 0.05
        if current_time - int(current_time) >= 0.6:
            current_time += (0.4)
        current_content = content.format(name, recruit_time.format(current_hour, current_minute))
        text_content.append(current_content)
    output_df = pd.DataFrame(
        {
            "姓名": names,
            "手机号": phones,
            "文本": text_content
        }
    )
    print(output_df)
    output_df.to_excel(output_file, index=False)