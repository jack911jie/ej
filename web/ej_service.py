import os
import sys
sys.path.append(os.path.join(os.path.dirname(os.path.dirname(os.path.realpath(__file__))),'modules'))
sys.path.append(os.path.join(os.path.dirname(os.path.dirname(os.path.realpath(__file__)))))
import ganzhi
import week_yun
import pandas as pd
import read_config
from flask import Flask, request, Response,render_template,send_file,make_response
import zipfile
import io
from datetime import datetime
import json
import xlwings as xw


class EjService(Flask):

    def __init__(self,*args,**kwargs):
        super(EjService, self).__init__(*args, **kwargs)
        config_fn=os.path.join(os.path.dirname(os.path.dirname(os.path.realpath(__file__))),'configs','ej_service.config')
        self.config_ej=read_config.read_json(config_fn)
        # print(self.config_ej)
        
        #路由
        #渲染页面
        #首页
        self.add_url_rule('/riyun',view_func=self.riyun)

        # 快团团
        self.add_url_rule('/ktt', view_func=self.ktt,methods=['GET','POST'])
        #日运录入
        self.add_url_rule('/riyun_input_page',view_func=self.riyun_input_page,methods=['GET','POST'])
        
        
        # 生成日运
        self.add_url_rule('/generate_riyun', view_func=self.generate_riyun,methods=['GET','POST'])

        # 结果打包并返回前端
        self.add_url_rule('/zip_and_download', view_func=self.zip_and_download,methods=['GET','POST'])


        # 写入日运表
        self.add_url_rule('/write_into_riyun_xlsx', view_func=self.write_into_riyun_xlsx,methods=['GET','POST'])
        # 写入日运表
        self.add_url_rule('/riyun_menu', view_func=self.riyun_menu,methods=['GET','POST'])


    def riyun_input_page(self):
        return render_template('/riyun_input.html')

    def riyun(self):
        return render_template('/riyun.html')

    def riyun_menu(self):
        return render_template('/riyun_menu.html')

    def ktt(self):
        return render_template('/ktt.html')

    def write_into_riyun_xlsx(self):
        print('writing into riyun xlsx')
        data=request.json
        y,m,d=data['date'].split('-')
        gz=ganzhi.GanZhi().cal_dateGZ(int(y),int(m),int(d),8)
        data['tg']=gz['bazi'][4]
        data['dz']=gz['bazi'][5]

        try:
            df=pd.DataFrame(data,index=[0])
            df=df[['date', 'weekday', 'tg', 'dz', 'color-mu', 'txt-mu', 'color-huo', 'txt-huo', 'color-tu', 'txt-tu', 'color-jin', 'txt-jin', 'color-shui', 'txt-shui']]
            # print(df)
            with open(os.path.join(os.path.dirname(__file__),'config','riyun.config'),'r',encoding='utf-8') as f:
                config_ej=json.load(f)
            riyun_fn=config_ej['riyun_fn']


            app=xw.App(visible=False)
            wb=app.books.open(riyun_fn)
            sheet=wb.sheets['运势']
            last_row=sheet.range('A1048576').end('up').row
            
            last_date=sheet.range(f'A{last_row}').value

            date_input=datetime.strptime(data['date']+' 00:00:00','%Y-%m-%d %H:%M:%S')

            # print(date_input)
            # print(last_date)
    
            dates=sheet.range(f'A3:A{last_row}').value
            try:
                row_number=dates.index(date_input)+3
            except:
                row_number=last_row+1
            finally:            
                print(f'write into row {row_number}')
                
                # 将 DataFrame 数据写入 Excel 工作表的指定行号
                print(df.values.tolist)
                sheet.range(f'A{row_number}:N{row_number}').value=df.values

                wb.save(riyun_fn)
                wb.close()
                app.quit()
        except Exception as e:
            print('write xlsx error',e)
            return {'res':'failed','error':'write xlsx error'}

        

        return {'res':'ok'}

    def run_week_txt_cover(self,fn_num,prd=['20220822','20220828'],sense_word_judge='yes'):
        work_dir=self.config_ej['work_dir']
        output_dir=self.config_ej['output_dir']
        if fn_num=='1':
            xls=os.path.join(work_dir,'运势','运势.xlsx')
        else:
            xls=os.path.join(work_dir,'运势','运势2.xlsx')

        try:

            eachday_output_dir=os.path.join(output_dir,'日穿搭')
            cover_save_dir=os.path.join(output_dir,'日穿搭','0-每周运势封图')

            # print('\n正在处理每日穿搭配色图\n')
            week_pic=week_yun.ExportImage(work_dir=work_dir)
            res=week_pic.batch_deal(prd=prd,out_put_dir=eachday_output_dir,xls=xls)
            if res['res']=='ok':
                dec_txt=res['res_data']

                print('\n正在处理每日穿搭配色文案\n')
                week_txts=week_yun.ExportWeekYunTxt(work_dir=work_dir,import_dec_dic=dec_txt)
                week_txts.all_date_wx(prd=prd,xls=xls,save_dir=eachday_output_dir,sense_word_judge=sense_word_judge)

                # print('\n正在生成本周周运封图\n')
                # week_cover=week_yun.WeekYunCover(work_dir=work_dir)
                # week_cover.export(prd=['20220822','20220828'],save_dir=cover_save_dir)

                # os.startfile(eachday_output_dir)
                return {'res':'ok'}
            else:
                return {'res':'failed','error':res['error']}
        except Exception as e:
            # raise FERROR('error where generate riyun pics and txt')
            print(e)
            return {'res':'failed','error':'error when generate riyun pics and txt'}



    def generate_riyun(self):
        data=request.data.decode('utf-8')
        print(data)
        fn_num,start_date,end_date=data.split('|')
        start_date_input=start_date.replace('-','')
        end_date_input=end_date.replace('-','')
        try:
            res_generate=self.run_week_txt_cover(fn_num=fn_num,prd=[start_date_input,end_date_input])  
            # self.zip_and_download(prd=[start_date,end_date])  
            # print(res_generate)
            if res_generate['res']!='ok':
                return {'res':'failed','error':res_generate['error']}
            

            #将日期写入临时文件

            tmp_dir=tmp_fn=os.path.join(self.config_ej['output_dir'],'日穿搭','zip')
            if not os.path.exists(tmp_dir):
                os.makedirs(tmp_dir)
            tmp_fn=os.path.join(tmp_dir,'riyun_tmp')
            with open (tmp_fn, 'w') as f:
                f.write(f'{start_date},{end_date}')

            return {'res':'ok','res_data':f'{start_date},{end_date},OK'}

            
            # return zip
        except Exception as e:
            print('riyun() Error:',e)
            return {'res':'failed','error':e}

    def zip_and_download(self):
        # prd_input=request.data.decode('utf-8')
        tmp_fn=os.path.join(self.config_ej['output_dir'],'日穿搭','zip','riyun_tmp')
        with open (tmp_fn, 'r') as f:
            prd_input=f.read()
        
        # print(prd_input)
        # prd_input='2023-08-22,2023-08-22'
        prd=prd_input.split(',')
        print('zip_and_download() ',prd)
        if prd[0]==prd[1]:
            output_filename=prd[0]
        else:
            output_filename=prd[0]+'-'+prd[1]

        path=os.path.join(self.config_ej['output_dir'],'日穿搭')
        memory_file = io.BytesIO()
        with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_STORED) as zipf:
            for root, dir, files in os.walk(path):
                for file in files:                    
                    try:     
                        if datetime.strptime(prd[0],'%Y-%m-%d')<=datetime.strptime(root.split('\\')[-1],'%Y-%m-%d')<=datetime.strptime(prd[1],'%Y-%m-%d'):
                            file_path = os.path.join(root, file)
                            arcname = os.path.relpath(file_path, path)
                            zipf.write(file_path,arcname)
          
                    except:
                        pass

        memory_file.seek(0)

        return Response(memory_file.getvalue(),
                        mimetype='application/zip',
                        headers={'Content-Disposition': f'attachment;filename={output_filename}.zip'})


class FERROR(Exception):
    pass


if __name__ == '__main__':
    app = EjService(__name__)
    if len(sys.argv)>1:
        # print(f'服务器为：{sys.argv[1]}:5000')
        app.run(debug=True,host=sys.argv[1],port=5023)
    else:
        app.run(debug=True)

    # app.run(debug=True,host='127.0.0.1',port=5023)
    # app.run(debug=True,host='192.168.10.2',port=5000)
    # app.run(debug=True,host='192.168.1.41',port=5001)
    # app.run(debug=True,host='192.168.1.149',port=5000)
    # res=wecom_dir()
    # print(res)
