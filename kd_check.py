import os
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
# 创建Chrome选项
chrome_options = Options()
chrome_options.add_argument('--headless')  # 设置为无界面模式
import re
import time
import pandas as pd
pd.set_option('display.unicode.east_asian_width', True) #设置输出右对齐
import copy
import openpyxl
from openpyxl.styles import Font, Color
from openpyxl.utils import get_column_letter
import numpy as np



class FruitKd:
    def __init__(self,chromedriver_path):
        if chromedriver_path:
            self.driver=webdriver.Chrome(chromedriver_path)
        else:
            pass

        #无界面模式
        # self.driver=webdriver.Chrome(chromedriver_path,options=chrome_options)

    def close_chrome_driver(self):
        self.driver.quit()

    #第一次获取物流记录
    def connect_kd_web(self,phone_number,url):
        # 打开网页
        self.driver.get(url)
        time.sleep(1)

        # 设置要输入的号码
        phone_number_4 = str(phone_number)[-4:] # 替换成你想要的号码

        # 找到收件人输入框并输入号码
        try:
            recipient_input = self.driver.find_element(By.ID,"query_str")
            recipient_input.send_keys(phone_number_4)

            # 提交表单
            self.driver.find_element(By.ID,"submit_product_query").click()

            # 获取结果
            result = self.driver.page_source

            # print(result.text)

            # 关闭浏览器
            # self.driver.quit()
        except Exception as err:
            print('连接查询网址时出错')
            result=''
        # finally:
        #     result=''
        #     self.driver.quit()

        return result

    #如有重复记录的获取方法2
    def connect_kd_web_2(self,phone_number,url):
        # 打开网页
        self.driver.get(url)
        time.sleep(1)

        # 设置要输入的号码
        phone_number_4 = str(phone_number)[-4:] # 替换成你想要的号码

        # 找到收件人输入框并输入号码
        try:
            # recipient_input = self.driver.find_element(By.ID,"query_str")
            # recipient_input.send_keys(phone_number_4)

            # 提交表单
            # self.driver.find_element(By.ID,"submit_product_query").click()

            # 获取结果
            result = self.driver.page_source

            # print(result.text)

            # 关闭浏览器
            # self.driver.quit()
        except Exception as err:
            print('连接查询网址时出错')
            result=''
        # finally:
        #     result=''
        #     self.driver.quit()

        return result

    def many_or_single_result(self,res,phn):
        soup=BeautifulSoup(res,'html.parser')

        if '查询到多条记录' in res:
            print('{} 有多条记录，正在匹配姓名。'.format(phn))
            li_elements = soup.select('li[role="presentation"]')
            # <li role="presentation" class="active"><a href="http://kd.dh.cx/df66d/7329/0"> 06月16日（单号：776260474339358） </a></li>
            urls=[]
            for li in li_elements:
                a_tag=li.find('a')
                if a_tag:
                    url=a_tag['href']
                    urls.append(url)
            
            # print(urls)

            ress=[]
            for url in urls:
                # url=url.split('//')[-1]
                res=self.connect_kd_web_2(phone_number=phn,url=url)
                tmp_res=self.deal_result(res=res)
                ress.append(tmp_res[0])

            result=ress
        else:
            print('{} 有一条记录'.format(phn))
            result=self.deal_result(res)
                
        
        return result
           
    def deal_result(self,res):
        # 处理结果
        # 这里需要根据具体情况来解析和提取你想要的信息

         # 输出结果

        soup=BeautifulSoup(res,'html.parser')
        li_elements = soup.find_all('li')

        

        # 遍历每个li元素
        info=[]
        for li in li_elements:

            tt=''
            # 检查li元素的文本内容是否包含目标字符
            try:
                if '物流单号' in li.get_text():
                    # 找到了包含目标字符的li元素
                    kd_code_txt=li
                    code=re.findall(r'\d{15}',str(li.text))[0]
                    tt=tt+code
            
    
                if '收件人：' in li.get_text():
                    kd_receiver_txt=li

                    name=str(li.text).split('：')[-1]
                    name=name.strip()

                    tt=tt+name
                    # break

                if len(tt)>0:
                    info.append(tt)
            except:
                pass
        
        info_0=info[0::2]
        info_1=info[1::2]

        
        # result=[n_ph[info_0[x],info_1[x]] for x in range(len(info_0))]

        result=[]
        for x in range(len(info_0)):
            n_ph={}
            n_ph[info_1[x]]=info_0[x]
            result.append(n_ph)        

        return result

    def batch_phone_number(self,phn_name_list=['15678892330阿晓','17853297329李君'],url='http://kd.dh.cx/df66d'):
        kd_id_list=[]
        for phn_name in phn_name_list:
            phn=phn_name[:11]
            name=phn_name[11:]
            id_get=self.connect_kd_web(phone_number=phn,url=url)

            # print('id_get:',id_get)
            # with open('d:\\py\\test\\res.html', 'r', encoding='utf-8') as fhtml:
            #     id_get = fhtml.read()
            if id_get:
                kddh=self.many_or_single_result(res=id_get,phn=phn)

                # print(kddh)
                if kddh:
                    for kd in kddh:
                        # print(kd)
                        try:
                            kd_id_list.append([str(phn),name,str(kd[name])])
                        except Exception as e:
                            # print('{} 号码下未找到 {} 的订单号。'.format(phn,name))
                            kd_id_list.append([str(phn),name,np.nan])
                else:
                    #  print('{} 号码下未找到 {} 的订单号。'.format(phn,name))
                     kd_id_list.append([str(phn),name,np.nan])

        return kd_id_list

    def read_order_excel(self,xls='e:\\temp\\ejj\\团购群\\订单\\给果园的订单\\团团好果06.17订单-01.xlsx'):
        df_read=pd.read_excel(xls)
        df_new=df_read[['收件人手机','收件人姓名']]
        df=copy.deepcopy(df_new)
        df['联系电话收货人']=df_new['收件人手机'].astype(str)+df_new['收件人姓名']
        
        return df['联系电话收货人'].tolist()

    def read_dl_order_excel(self,xls='e:\\temp\\ejj\\团购群\\订单\\wuliu2023-06-17 21_27_45.xlsx.xlsx'):
        df_dl_order=pd.read_excel(xls,sheet_name='订单统计')
        df_dl_order_2=df_dl_order[['联系电话','收货人']]
        df_dl_order_new=df_dl_order_2.copy()
        df_dl_order_new['联系电话收货人']=df_dl_order_2['联系电话'].astype(str)+df_dl_order_2['收货人']

        return df_dl_order_new['联系电话收货人'].tolist()

    def order_to_guoyuan(self,dl_xls='E:\\temp\\ejj\\团购群\\订单\\0614导出.xlsx',output_dir='e:\\temp\\ejj\\团购群\\给果园的订单',expand_accounts='yes',exp='yes'):
        df_order=pd.read_excel(dl_xls,sheet_name='订单列表')
        df_order_out=df_order.copy()
        df_order_out=df_order_out[['收货人','联系电话','详细地址','商品','订单号']]
        df_order_out.rename(columns={'收货人':'收件人姓名','联系电话':'收件人手机','详细地址':'收件人地址','商品':'品名'},inplace=True)
        df_order_out['收件人手机']=df_order_out['收件人手机'].apply(lambda x: str(x))
        df_order_out['规格']=df_order_out['品名'].apply(lambda x: re.findall(r'(\d\d斤装|\d斤装)',x)[0])
        df_order_out['数量']=df_order_out['品名'].apply(lambda x: int(str(x).split('+')[-1]))
        df_order_out['发件人姓名']='团团好果'
        df_order_out['发件人手机']='18077796420'
        df_order_out['发件人电话']=''
        df_order_out['发件人地址']=''
        df_order_out['发件人单位']=''
        df_order_out['收件人电话']=''
        df_order_out['收件人单位']=''

        # df_order_out['规格']=''
        
        
        df_order_out['备注']=''
        df_order_out['代收金额']=''
        df_order_out['到付金额']=''

        df_order_out=df_order_out[['发件人姓名','发件人手机','发件人电话','发件人地址','发件人单位','收件人姓名','收件人手机','收件人电话','收件人地址','收件人单位','品名','规格','数量','备注','订单号','代收金额','到付金额']]

      
        if expand_accounts=='yes':
            df_repeated=df_order_out.loc[df_order_out.index.repeat(df_order_out['数量'])]
            df_repeated=df_repeated.reset_index(drop=True)
            df_repeated['数量']=1
            df_res=df_repeated
        else:
            df_order_out['数量']=df_order_out['品名'].apply(lambda x: int(str(x).split('+')[-1]))
            df_res=df_order_out


        # print(df_repeated)
        if exp=='yes':
            xlsname=dl_xls.split('\\')[-1].split('.')[0].split('-')
            datetxt,num=xlsname[0],xlsname[2]
            fn='团团好果'+datetxt[4:6]+'.'+datetxt[6:]+'订单'+'-'+num+'.xlsx'
            fn=os.path.join(output_dir,fn)
            df_res.to_excel(fn,index=False)
            print('导出完成')
            self.guoyuan_xlsx_style(xls=fn)
            os.startfile(output_dir)
        return df_res

    def guoyuan_xlsx_style(self,xls):
        wb=openpyxl.load_workbook(xls)
        ws=wb.active

        # for cell in ws[1]:
            
        #     font=Font(size=13,bold=True)
        #     cell.font=font

        font=Font(size=13,bold=True)
        ws['A1'].font=font

        for cls in ['A','B','C','D','F','G','H','I']:
            cell=ws[cls+'1']
            font=Font(color='FF0000',bold=True)
            cell.font=font

        
        #调整列宽
        for cell in ws[1]:
            max_length = 0
            column_letter = get_column_letter(cell.column)
            # print(column_letter,cell.value)
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass

            if column_letter=='J':
                adjusted_width = (max_length + 15) * 2.2
            else:
                adjusted_width = (max_length + 5) * 1.2
            ws.column_dimensions[column_letter].width = adjusted_width
        
        wb.save(xls)
        print('修改格式完成')

    def multi_order_to_guoyuan(self,date=20230618,input_dir='E:\\temp\\ejj\\团购群\\订单',output_dir='e:\\temp\\ejj\\团购群\\订单\\给果园的订单',expand_accounts='yes',save_each_exp='no'):
        date=str(date)
        try:
            fns=[]
            for fn in os.listdir(input_dir):
                if re.match(date+'-导出订单-\d\d.xlsx',fn):
                    fns.append(os.path.join(input_dir,fn))

            dfs=[]
            for fn in fns:
                df_res=self.order_to_guoyuan(dl_xls=fn,output_dir=output_dir,expand_accounts=expand_accounts,exp=save_each_exp)
                if df_res.shape[0]>0:
                    dfs.append(df_res)

            df_concat=pd.concat(dfs)

            fn='团团好果'+date[4:6]+'.'+date[6:]+'订单（'+str(len(fns))+'个订单合并）.xlsx'
            fn=os.path.join(output_dir,fn)
            df_concat.to_excel(fn,index=False)
            print('导出完成')
            self.guoyuan_xlsx_style(xls=fn)
            os.startfile(output_dir)

            return df_concat

        except Exception as e:
            print(e)


    def write_xlsx_back_kd(self,input_xls='e:\\temp\\ejj\\团购群\\订单\\给果园的订单\\团团好果06.17订单-01.xlsx',
                            out_dir='e:\\temp\\ejj\\团购群\\订单\\带物流信息的回传文件',
                            url='http://kd.dh.cx/df66d',
                            kd_name='申通快递',
                            method='download'):
        if method=='download':
            #根据method值获取输入表内容
            df_input=pd.read_excel(input_xls)
            #生成客户手机及姓名列表
            exp_list=self.read_dl_order_excel(xls=input_xls)
            #获取客户物流单号
            res=self.batch_phone_number(phn_name_list=exp_list,url=url)
            
            df_write=pd.DataFrame(data=res,columns=['收件人手机','收件人姓名','物流单号'])

            print('\n\n------------------------\n以下为匹配结果：\n',df_write)
            df_kd=df_write.dropna(how='any',subset=['物流单号'])

            #如无快递单号的df
            if df_kd.shape[0]==0:
                return '未返回有效快递单号'
                # pass                
            else:
                df_input['物流公司（必填）']=kd_name
                kd_id_dic=df_kd.set_index('收件人手机')['物流单号'].to_dict()
                df_input['物流单号（必填）']=df_input['联系电话'].apply(lambda x:kd_id_dic.get(str(x),''))

                if not os.path.exists(out_dir):
                    os.makedirs(out_dir)
                out_fn=os.path.join(out_dir,input_xls.split('\\')[-1][:-5]+'-已写入物流单号待上传.xlsx')
                df_input.to_excel(out_fn,index=False)

                os.startfile(out_dir)
                return '\n写入待回传文件完成'

        else:
            pass

        self.close_chrome_driver()
        
if __name__=='__main__':
    
 
    # 一、从快团团批量导入订单处理后生成给果园的订单。只能处理一个文件。
    p=FruitKd(chromedriver_path='')
    # rs=p.order_to_guoyuan(dl_xls='E:\\temp\\ejj\\团购群\\订单\\20230618-导出订单-02.xlsx',
    #                             output_dir='e:\\temp\\ejj\\团购群\\订单\\给果园的订单',
    #                             expand_accounts='yes')
    #参数说明：
    # dl_xls：从快团团批量导出的订单，文件名修改为：20230618-导出订单-02.xlsx 的格式
    # output_dir：生成给果园的订单文件后存放的文件夹
    # expand_accounts：对于购买多于1件的商品，根据实际数量生成相应的记录。例如同一客户购买了3件，同一条记录生成3条，防止商家发货漏单。
    # print(rs)

    #批量处理同一天的不同订单，可将多个订单生成一个合并发货清单文件给果园。
    rs=p.multi_order_to_guoyuan(date=20230622,
                                input_dir='E:\\temp\\ejj\\团购群\\订单',
                                output_dir='e:\\temp\\ejj\\团购群\\订单\\给果园的订单',
                                expand_accounts='yes',
                                save_each_exp='no')

    #参数说明：
    # dl_xls：从快团团批量导出的订单，文件名修改为：20230618-导出订单-02.xlsx 的格式
    # output_dir：生成给果园的订单文件后存放的文件夹
    # expand_accounts：对于购买多于1件的商品，根据实际数量生成相应的记录。例如同一客户购买了3件，同一条记录生成3条，防止商家发货漏单。
    # each_exp：处理到每个订单时，是否分别保存。默认为no，即只保存最后生成的合并订单。
    # print(rs)

    # 二、果园返单后，通过下载快团团模板文件查询快递单号并写入待上传文件
    # p=FruitKd(chromedriver_path='D:/Program Files (x86)/ChromeWebDriver/chromedriver')
    # res=p.write_xlsx_back_kd(input_xls='e:\\temp\\ejj\\团购群\\订单\\wuliu2023-06-17 21_27_45.xlsx.xlsx',
    #                         out_dir='e:\\temp\\ejj\\团购群\\订单\\带物流信息的回传文件',
    #                         url='http://kd.dh.cx/df66d',
    #                         kd_name='申通快递',
    #                         method='download')
    # print(res)

    #参数说明：
    # input_xls：从快团团导出的待回传清单
    # output_dir：生成给果园的订单文件后存放的文件夹
    # url：查询地址，可能会经常改变
    # kd_name: 快递公司名称。写入生成的回传清单中。
    # method: download——从快团团导出的待回传清单文件查询物流单号并生成回传文件以上传。目前仅支持这一选项。




    