import pandas as pd
import random
import os, sys, stat




file_path = ""

# os.chmod(file_path, stat.S_IRWXU and stat.S_IRWXG and stat.S_IRWXO)

df = pd.read_excel(file_path)



data = {}
for i in range(len(df)):
    a_code = df.iloc[i, 0]
    a_word = df.iloc[i, 1]
    a_pron = df.iloc[i, 2]
    if pd.notna(a_code):  
        data[a_code] = (a_word, a_pron)

    d_code = df.iloc[i, 3]
    d_word = df.iloc[i, 4]
    d_pron = df.iloc[i, 5]
    if pd.notna(d_code):  
        data[d_code] = (d_word, d_pron)


used_codes = set()

def conduct_quiz(selected_data, mode):
        global used_codes 
        while True:

            
            code = random.choice(list(selected_data.keys()))

            while code in used_codes:
                code = random.choice(list(selected_data.keys()))

            word, pron = selected_data[code]

            if mode == '1':
                print(f"请输入编码 {code} 的联想词：")
                user_input = input("联想词: ")
                if user_input == 'q':
                    break
                if user_input == word:
                    print("正确")
                else:
                    print(f"错误。正确的联想词是：{word}")

            elif mode == '2':
                print(f"请输入编码 {code} 的读音：")
                user_input = input("读音: ")
                if user_input == 'q':
                    break
                if user_input == pron:
                    print("正确")
                else:
                    print(f"错误。正确的读音是：{pron}")

            elif mode == '3':
                print(f"请输入编码 {code} 的联想词和读音：")
                user_input_word = input("联想词: ")
                if user_input_word == 'q':
                    break
                user_input_pron = input("读音: ")
                if user_input_pron == 'q':
                    break
                if user_input_word == word and user_input_pron == pron:
                    print("正确")
                    
                else:
                    print(f"错误。正确的联想词是：{word}，正确的读音是：{pron}")

            
            used_codes.add(code)
            if len(used_codes) >= 24:
                used_codes.clear()

while 1: 


    print("请选择编码群：")
    print("1: 全部编码")
    print("2: 自定义编码群（例如：A, B, C）")
    print("输入 q 退出")

    group_choice = input("输入选择的编码群编号 (1/2/q): ")

    if group_choice == 'q':
        exit()


    if group_choice == '1':
        print("请选择模式：")
        print("1: 联想词")
        print("2: 读音")
        print("3: 联想词和读音")
        print("输入 q 退出")

        mode = input("输入模式编号 (1/2/3/q): ")
        if mode == 'q':
            exit()

        if mode not in ['1', '2', '3']:
            print("无效的模式选择")
            exit()
        selected_data = data

    elif group_choice == '2':
        custom_groups = input("输入自定义编码群，用逗号分隔（例如：A,B,C）: ").split(',')
        print("请选择模式：")
        print("1: 联想词")
        print("2: 读音")
        print("3: 联想词和读音")
        print("输入 q 退出")

        mode = input("输入模式编号 (1/2/3/q): ")
        if mode == 'q':
            exit()

        if mode not in ['1', '2', '3']:
            print("无效的模式选择")
            exit()
        selected_data = {k: v for k, v in data.items() if any(k.startswith(group) for group in custom_groups)}
    else:
        print("无效的选择")
        exit()

    conduct_quiz(selected_data, mode)
