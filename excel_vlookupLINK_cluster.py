import pandas as pd  
import tkinter as tk  
from tkinter import filedialog
from tkinter import messagebox  
import datetime
import time
import os  
import numpy as np
  
# 创建Tkinter界面  
root = tk.Tk()  
root.title("链路&网元连接聚类小程序")  
  
# 创建文本框  
path_text = tk.StringVar()  
path2_text = tk.StringVar() 
path3_text = tk.StringVar()  
path_text.set("请选择链路表 ")  
path2_text.set("请选择网元表 ") 
path3_text.set("点击开始处理并选择输出文件夹 ") 
path_label = tk.Label(root, textvariable=path_text)  
path2_label = tk.Label(root, textvariable=path2_text)
path3_label = tk.Label(root, textvariable=path3_text)  
path_label.pack()  
path2_label.pack()
path3_label.pack()   
  
# 创建选择文件按钮  
def select_file1():  
    filepath1 = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])  
    if filepath1 == "":  
        return  
    path_text.set(filepath1)  
  
def select_file2():  
    filepath2 = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])  
    if filepath2 == "":  
        return  
    path2_text.set(filepath2)

def select_output_folder():
    #global foldername
    foldername = filedialog.askdirectory()
    if foldername:
        path3_text.set(foldername)
    return foldername  #加个return获取的文件夹就能直接在另一个函数里用了，实现传递，不用直接用button或者用global全局变量
  
button1 = tk.Button(root, text="选择链路表", command=select_file1)  
button1.pack()  
button2 = tk.Button(root, text="选择网元表", command=select_file2)  
button2.pack()
#button3 = tk.Button(root, text="选择输出文件夹", command=select_output_folder)  
#button3.pack()  
  
# 创建处理按钮   def process_data(select_output_folder)
def process_data():  
    try:
        start_time=time.time()  
        df1 = pd.read_excel(path_text.get())  
        df2 = pd.read_excel(path2_text.get())

        foldername=select_output_folder()

        # 创建一个新的列“电路名称前缀网元”和“电路名称后缀网元”  
        df1['电路名称前缀网元'], df1['电路名称后缀网元'] = zip(*df1['电路名称'].str.split('::').apply(lambda x: [x[0], x[1]]))
        #去除这两列的端口等后缀
        df1['电路名称前缀网元'] = df1['电路名称前缀网元'].apply(lambda x: x.split(':')[0] if ':' in x else x)  
        df1['电路名称后缀网元'] = df1['电路名称后缀网元'].apply(lambda x: x.split(':')[0] if ':' in x else x)
        
        #判断是否为A设备(自己写，但是报错list下标超界)
        #df1['真正A端设备名称'] = df1["电路名称前缀网元"].apply(lambda x: x if x.split('-')[2] == 'A' else pd.NA) #是A设备就填入，否则空的

        #df1['真正A端设备名称'] = df1['真正A端设备名称'].apply(lambda x: x) if x in df1['真正A端设备名称'] x==pd.isna()

        # 判断电路名称前缀列中单元格内容以第二个短横线"-"分割后的字符是否为"A"  
        #df1["是否为A"] = df1["电路名称前缀网元"].str.split("-").str[2] == "A"
        # 判断是否为B设备，即是否含有"-B-"，因为A设备有A/A1/A2且字符位置不确定
        #df1["是否为B"] = df1["电路名称前缀网元"].apply(lambda x: "TRUE" if "-B-" in x else "FALSE") #可行 #个人想用contain实现
        df1["是否为B"] = df1["电路名称前缀网元"].apply(lambda x: True if "-B-" in x else False)  #个人想用contain实现
  
        # 根据判断结果，如果为True，则将电路名称前缀填入新列"真正A端设备名称"  
        # 否则将电路名称后缀填入新列"真正A端设备名称"  
        #df1["真正A端设备名称"] = df1[["电路名称前缀网元", "电路名称后缀网元", "是否为A"]].apply(lambda x: x[0] if x[2] == True else x[1], axis=1) #这个也能实现
        #df1["真正A端设备名称"] = df1.apply(lambda x: x["电路名称后缀网元"] if x["是否为B"] == "TRUE" else x["电路名称前缀网元"], axis=1) #可行
        df1["真正A端设备名称"] = df1.apply(lambda x: x["电路名称后缀网元"] if x["是否为B"] else x["电路名称前缀网元"], axis=1)

        # 打印结果以验证  
        #df1.to_excel("result_{:%Y%m%d%H%M%S}.xlsx".format(datetime.datetime.now()))
        
        '''
        df1 = df1.merge(df2[["区域", "网元名称"]], on="真正A端设备名称", how="left")  
        df1["真正A端区域"] = df1["区域"].values
        '''
        # 在A表中添加新列"真正A端区域"，连接B表  
        #df1 = df1.merge(df2[["网元名称", "区域"]], on="真正A端设备名称", how="left")
        df1 = df1.merge(df2[["网元名称", "区域", "环号"]], left_on="真正A端设备名称", right_on="网元名称", how="left")
          
  
        # 重命名"区域"列为"真正A端区域"  
        df1 = df1.rename(columns={"区域": "真正A端区域"})  
  
        # 删除重复的"真正A端设备名称"列  
        #df1 = df1.drop(columns="真正A端设备名称")
        df1.drop('网元名称', axis=1, inplace=True) #原地去除从网元表作为key连接过来的网元名称一列

        #本来在这里先聚类/排序，但是拼装后排序会把环号也聚一起，挺好的，环号在末端，不影响前面正常排序
        #df1.groupby('真正A端区域', sort=True)  #处理出错： This Series is a view of some other array, to sort in-place you must create a copy
        #df1['真正A端区域'].sort_values(inplace=True)  # 先排序 处理出错： This Series is a view of some other array, to sort in-place you must create a copy
        #df1['真正A端区域'].sort_values()
        #df_temp = df1.copy()
        #df_temp.groupby('真正A端区域', sort=True)
        #df_temp['真正A端区域'].sort_values()
        #df1 = df1.sort_values(by='真正A端区域')

        #拼装环号
        #f1['拼装环号'] = df1.apply(lambda x: x["真正A端区域"] + '/' + x['环号'] if "东莞" in x['真正A端区域'] else pd.NA)
        # 使用 pd.to_numeric 将区域列的NaN值转换为空字符串，但在转换之前需要将NaN替换为一个可以转换为数字的值，例如-9999  
        df1['环号'] = df1['环号'].replace(np.nan, -9999)  
  
        # 使用 pd.to_numeric 将区域列的值转换为数字，遇到-9999时返回NaN  
        #df1['真正A端区域'] = pd.to_numeric(df1['真正A端区域'], errors='coerce')  

        df1['环号'] = df1['环号'].astype(int)

        # 使用 fillna 将NaN值替换为空字符串  
        #df1['真正A端区域'] = df1['真正A端区域'].fillna('')
        df1["环号"] = df1["环号"].replace(-9999, '')
        df1['环号'] = df1['环号'].astype(str)
        df1['真正A端区域'] = df1['真正A端区域'].astype(str)
        df1['真正A端区域'] = df1['真正A端区域'].replace('nan', '')
        df1["拼装环号"] = df1.apply(lambda x: x["真正A端区域"] + '/' + x['环号'] if x['环号'] != '' else '', axis=1)
        
        #聚类/排序
        df1 = df1.sort_values(by='拼装环号')
        #global foldername
        df1.to_excel(os.path.join(foldername, "output_" + time.strftime("%Y%m%d%H%M%S") + ".xlsx"))
        #df1.to_excel(os.path.join(path3_text.get(), "output_" + time.strftime("%Y%m%d%H%M%S") + ".xlsx"))

        total_time=time.time()-start_time
        messagebox.showinfo("提示", f"处理完成，耗时{total_time:.2f}秒")

        # 处理链路表数据  
        #df1['真正的A网元'] = df1['电路名称'].apply(lambda x: ':' if ':' in x else x)  
        #df1['区域+环号'] = df1['真正的A网元'].apply(lambda x: x + df2['环号'].values[0] + df2['区域'].values[0])  
        '''
        df['真正的A网元'] = df['电路名称'].apply(lambda x: x.split('::')[0].strip() if x.endswith('::') else x)  
        df['区域+环号'] = df['真正的A网元'].apply(lambda x: f'{x["区域"]}-{x["环号"]}' if pd.notna(x["区域"]) and pd.notna(x["环号"]) else x)

        联接后的表格 = pd.merge(链路表, 网元表, on='真正的A网元', how='outer')
        '''
        '''
        # 根据“区域”做聚类  
        df3 = df1.groupby('区域').agg(lambda x: list(x)).reset_index()  
        df3['区域'] = df3['区域'].apply(lambda x: ' '.join(x))  
        df3['环号'] = df3['环号'].apply(lambda x: ' '.join(x))  
        df3['区域+环号'] = df3['区域+环号'].apply(lambda x: ' '.join(x))  
        df3 = df3[['区域', '环号', '区域+环号']]  
        df3.to_excel("result_{:%Y%m%d%H%M%S}.xlsx".format(datetime.datetime.now()))  
        print("处理完成，结果已保存到result_{:%Y%m%d%H%M%S}.xlsx".format(datetime.datetime.now()))  '''
    except Exception as e:  
        print("处理出错：", e)
          
  
#button3 = tk.Button(root, text="选择输出文件夹并开始处理", command=lambda: process_data(select_output_folder))
button3 = tk.Button(root, text="选择输出文件夹并开始处理", command=process_data)
button3.pack() 
#button4 = tk.Button(root, text="开始处理", command=process_data)  
#button4.pack()  
  
root.mainloop()