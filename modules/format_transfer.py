def number_to_chinese(number):
    chinese_dict = {
        0: '零',
        1: '一',
        2: '二',
        3: '三',
        4: '四',
        5: '五',
        6: '六',
        7: '七',
        8: '八',
        9: '九',
        10: '十',
        100: '百',
        1000: '千',
        10000: '万'
    }

    if number in chinese_dict:
        return chinese_dict[number]
    elif 10 <= number < 100:
        tens = number // 10
        units = number % 10
        if units == 0:
            return chinese_dict[tens] + chinese_dict[10]
        else:
            return chinese_dict[tens] + chinese_dict[10] + chinese_dict[units]
    elif 100 <= number < 1000:
        hundreds = number // 100
        tens_units = number % 100
        if tens_units == 0:
            return chinese_dict[hundreds] + chinese_dict[100]
        else:
            return chinese_dict[hundreds] + chinese_dict[100] + number_to_chinese(tens_units)
    else:
        return str(number)  # 返回原始阿拉伯数字表示