import os
import sys
sys.path.append(os.path.join(os.path.dirname(os.path.dirname(os.path.realpath(__file__))),'modules'))
sys.path.append(os.path.join(os.path.dirname(os.path.dirname(os.path.realpath(__file__)))))
import week_yun
import read_config
from flask import Flask, request, Response,render_template,send_file
import zipfile
import io
from datetime import datetime


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
        
        
        # 生成日运
        self.add_url_rule('/generate_riyun', view_func=self.generate_riyun,methods=['GET','POST'])

        # 结果打包并返回前端
        self.add_url_rule('/zip_and_download', view_func=self.zip_and_download,methods=['GET','POST'])

    def riyun(self):
        return render_template('/riyun.html')

    def run_week_txt_cover(self,fn_num,prd=['20220822','20220828'],sense_word_judge='yes'):
        work_dir=self.config_ej['work_dir']
        output_dir=self.config_ej['output_dir']
        if fn_num=='1':
            xls='d:\\工作目录\\ejj\\运势\\运势.xlsx'
        else:
            xls='d:\\工作目录\\ejj\\运势\\运势2.xlsx'

        eachday_output_dir=os.path.join(output_dir,'日穿搭')
        cover_save_dir=os.path.join(output_dir,'日穿搭','0-每周运势封图')

        # print('\n正在处理每日穿搭配色图\n')
        week_pic=week_yun.ExportImage(work_dir=work_dir)
        dec_txt=week_pic.batch_deal(prd=prd,out_put_dir=eachday_output_dir,xls=xls)

        # print('\n正在处理每日穿搭配色文案\n')
        week_txts=week_yun.ExportWeekYunTxt(work_dir=work_dir,import_dec_dic=dec_txt)
        week_txts.all_date_wx(prd=prd,xls=xls,save_dir=eachday_output_dir,sense_word_judge=sense_word_judge)

        # print('\n正在生成本周周运封图\n')
        # week_cover=week_yun.WeekYunCover(work_dir=work_dir)
        # week_cover.export(prd=['20220822','20220828'],save_dir=cover_save_dir)

        # os.startfile(eachday_output_dir)

    def generate_riyun(self):
        data=request.data.decode('utf-8')
        fn_num,start_date,end_date=data.split('|')
        start_date_input=start_date.replace('-','')
        end_date_input=end_date.replace('-','')
        try:
            self.run_week_txt_cover(fn_num=fn_num,prd=[start_date_input,end_date_input])  
            # self.zip_and_download(prd=[start_date,end_date])  
            return f'{start_date},{end_date},OK'
        except Exception as e:
            print('riyun() Error:',e)
            return 'error'+e

    def zip_and_download(self):
        prd_input=request.data.decode('utf-8')
        
        prd=prd_input.split(',')
        print('zip_and_download() ',prd)
        if prd[0]==prd[1]:
            output_filename=prd[0]
        else:
            output_filename=prd[0]+'-'+prd[1]

        path=os.path.join(self.config_ej['output_dir'],'日穿搭')
        memory_file = io.BytesIO()
        with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dir, files in os.walk(path):
                for file in files:                    
                    try:     
                        if datetime.strptime(prd[0],'%Y-%m-%d')<=datetime.strptime(root.split('\\')[-1],'%Y-%m-%d')<=datetime.strptime(prd[1],'%Y-%m-%d'):
                            file_path = os.path.join(root, file)
                            arcname = os.path.relpath(file_path, path)
                            zipf.write(file_path,arcname)
          
                    except:
                        pass
            
        # print(zipf)
        memory_file.seek(0)
        response = Response(memory_file.read(), content_type='application/zip')
        response.headers['Content-Disposition'] = f'attachment; filename={output_filename}.zip'
        return response
        # return send_file(memory_file, download_name='zip.zip',as_attachment=True)
        # send_file()

        


if __name__ == '__main__':
    app = EjService(__name__)
    if len(sys.argv)>1:
        print(f'服务器为：{sys.argv[1]}:5000')
        app.run(debug=True,host=sys.argv[1],port=5023)
    else:
        app.run(debug=True)
    # app.run(debug=True,host='127.0.0.1',port=5001)
    # app.run(debug=True,host='192.168.10.2',port=5000)
    # app.run(debug=True,host='192.168.1.41',port=5000)
    # app.run(debug=True,host='192.168.1.149',port=5000)
    # res=wecom_dir()
    # print(res)
