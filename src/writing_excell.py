
import os
import math
import pandas as pd
import eel

@eel.expose
def write_excell(path):
    item_names = {"氏名": 1, "所定\n勤務日数": 3, "訪問手当": 15, "待機手当": 24, "インセン\nティブ": 26, "夜勤手当": 29, "通勤\n手当": 33}

    sub_columns = {}
    max = 20
    end = 42


    for n in item_names:
        sub_columns[n] = []

    # for p in os.listdir('総合読み取りフォルダ'):
    for p in os.listdir(path):
        # input_file = pd.ExcelFile('総合読み取りフォルダ/' + p)
        input_file = pd.ExcelFile(os.path.join(path, p))
        sheet_names = input_file.sheet_names
        
        for sheet_name in sheet_names:
            main_columns = {}
            for i in range(max):
                main_columns[i] = {}
            df = pd.read_excel(os.path.join(path, p), index_col=None,header=None,sheet_name=sheet_name)
            filtered_df = df.iloc[1:,:end]
            # 読み取ったエクセルから空行を除き、表ごとに分割する:
            dfg = filtered_df.groupby((df.isnull().all(axis=1)).cumsum())

            for index, g in dfg:
                # if index == 1:
                    g = g.dropna(how="all")
                    if len(g) > 0:
                        g = g.values
                        d = pd.DataFrame( g[1:,:], columns=g[0])
                        index_num = 0
                        divisions = {}
                        for i,r in d.items():
                            index_num += 1
                            if type(i) is str:
                                divisions[index_num] = r
                        divisions[index_num] = ""
                        
                        num = 0
                        prev = 0
                        for i in divisions:
                            if num == 0:
                                column_item = d.iloc[:,0:i+1]
                            else:
                                column_item = d.iloc[:,prev-1:i-1]
                            num += 1
                            prev = i

                            for item_name in item_names:
                                empty_list = [0]*len(column_item)
                                
                                column_list = []
                                if item_name in d.columns:
                                    if item_name in column_item.columns:
                                        if column_item.shape[1] <= 1:
                                            for m in column_item[item_name]:
                                                if type(m) is str:
                                                    column_list.append(m)
                                                else:
                                                    if type(m) is int or type(m) is float:
                                                        if math.isnan(m):
                                                            column_list.append(0)
                                                        else: 
                                                            column_list.append(m)
                                                    else: 
                                                        column_list.append(0)
                                            main_columns[index][item_name] = column_list
                                        else: 
                                            for m in column_item.iloc[:,-1]:
                                                if type(m) is str:
                                                    column_list.append(m)
                                                else:
                                                    if type(m) is int or type(m) is float:
                                                        if math.isnan(m):
                                                            column_list.append(0)  
                                                        else: 
                                                            column_list.append(m)
                                                    else: 
                                                        column_list.append(0)                                 
                                            main_columns[index][item_name] = column_list

                                else:
                                    main_columns[index][item_name] = empty_list
                                                    
                                if item_name not in main_columns[index]:
                                    main_columns[index][item_name] = empty_list

            for i in range(max):
                for name in item_names:
                    if name in main_columns[i]:
                        sub_columns[name] += main_columns[i][name]

    data = {}
    for i in item_names:
        data[i] = sub_columns[i]

    df_result = pd.DataFrame(data)
    df_result = df_result[df_result["氏名"] != 0]
    print(len(df_result["氏名"]))
    df_result.to_excel('抽出結果/result.xlsx', index=False)
