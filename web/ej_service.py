import os
import sys
sys.path.append(os.path.join(os.path.dirname(os.path.dirname(os.path.realpath(__file__))),'modules'))
sys.path.append(os.path.join(os.path.dirname(os.path.dirname(os.path.realpath(__file__)))))
import ganzhi
import week_yun
import ktt_order_export
import pandas as pd
import read_config
from flask import Flask, request,jsonify, Response,render_template,send_file,make_response
import zipfile
import io
from datetime import datetime
import json
import xlwings as xw
import pymysql


class EjService(Flask):

    def __init__(self,*args,**kwargs):
        super(EjService, self).__init__(*args, **kwargs)
        config_fn=os.path.join(os.path.dirname(os.path.dirname(os.path.realpath(__file__))),'configs','ej_service.config')
        self.config_ej=read_config.read_json(config_fn)

        #读取快团团有关配置文件存放路径
        ktt_config=os.path.join(os.path.dirname(__file__),'config','ktt.config')
        with open(ktt_config,'r',encoding='utf-8') as f:
            self.ktt_config=json.loads(f.read())

        #读取不同发货商导表时与快团团对应的规格名称
        with open(self.ktt_config['ktt_col_map_config'],'r',encoding='utf-8') as f:
            self.col_map_config=json.loads(f.read())

        # 读取不同发货商发货表的列名
        with open(self.ktt_config['col_config_fn'],'r',encoding='utf-8') as f:
            self.config_ktt_order=json.loads(f.read())

        # 读取规格的默认名称返回前端
        with open(self.ktt_config['page_config_fn'],'r',encoding='utf-8') as f:
            config_page_default=json.loads(f.read())
        self.spec0=config_page_default['spec0']
        self.spec1=config_page_default['spec1']
        self.fn_info=config_page_default['fn_info']
        self.sender_name=config_page_default['sender_name']
        self.sender_tel=config_page_default['sender_tel']
        
        
        #路由
        #渲染页面
        #首页
        self.add_url_rule('/riyun',view_func=self.riyun)

        # 快团团
        self.add_url_rule('/ktt', view_func=self.ktt,methods=['GET','POST'])
        #快团团导单
        self.add_url_rule('/ktt_buy_list', view_func=self.ktt_buy_list_page,methods=['GET','POST'])
        #处理快团团导单
        self.add_url_rule('/ktt_deal_list', view_func=self.ktt_deal_list,methods=['GET','POST'])
        # 快团团结果打包并返回前端
        self.add_url_rule('/zip_and_download_ktt_exp_order', view_func=self.zip_and_download_ktt_exp_order,methods=['GET','POST'])
        #获取不同供应商下的不同规格的在表格中的名称（与快团团名称有对应，但不一定相同）
        self.add_url_rule('/ktt_return_spec', view_func=self.ktt_return_spec,methods=['GET','POST'])
        #渲染茶叶入库页面
        self.add_url_rule('/ktt_purchase_tea', view_func=self.ktt_purchase_tea_page,methods=['GET','POST'])
        #渲染茶叶出库页面
        self.add_url_rule('/ktt_sale_tea', view_func=self.ktt_sale_tea_page,methods=['GET','POST'])
        #渲染茶叶查询页面
        self.add_url_rule('/ktt_query_tea', view_func=self.ktt_query_tea_page,methods=['GET','POST'])
        #渲染茶叶菜单页面
        self.add_url_rule('/ktt_tea_menu', view_func=self.ktt_tea_menu_page,methods=['GET','POST'])

        #处理前端
        #获取操作人信息
        self.add_url_rule('/ktt_fetch_opr', view_func=self.ktt_fetch_opr,methods=['GET','POST'])
        #获取产品列表
        self.add_url_rule('/ktt_fetch_product', view_func=self.ktt_fetch_product,methods=['GET','POST'])
        #获取规格列表
        self.add_url_rule('/ktt_fetch_specs', view_func=self.ktt_fetch_specs,methods=['GET','POST'])
        #处理入库
        self.add_url_rule('/ktt_purchase_deal',view_func=self.ktt_purchase_deal,methods=['GET','POST'])
        #处理出库
        self.add_url_rule('/ktt_sale_deal',view_func=self.ktt_sale_deal,methods=['GET','POST'])
        #返回统计剩余数量        
        self.add_url_rule('/ktt_stat_deal',view_func=self.ktt_stat_deal,methods=['GET','POST'])
        #获取包装列表
        self.add_url_rule('/ktt_fetch_pkgs', view_func=self.ktt_fetch_pkgs,methods=['GET','POST'])        

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


    def connect_mysql(self):
        with open(os.path.join(os.path.dirname(os.path.realpath((__file__))),'config','db.config'),'r',encoding='utf-8') as f:
            cfg=json.load(f)
        # 连接数据库
        conn = pymysql.connect(
            host=cfg['host'],       # 数据库主机地址
            user=cfg['user'],     # 数据库用户名
            password=cfg['password'], # 数据库密码
            database=cfg['database'],  # 要连接的数据库名称
            port=cfg['port']
        )

        return conn
    
    def ktt_tea_menu_page(self):
        return render_template('/ktt_tea_menu.html')

    def ktt_query_tea_page(self):
        return render_template('/ktt_query_tea.html')

    def ktt_sale_tea_page(self):
        return render_template('/ktt_sale_tea.html')

    def ktt_purchase_tea_page(self):
        return render_template('/ktt_purchase_tea.html')
    
    def ktt_purchase_deal(self):
        data=request.json
        # print(data)
        conn=self.connect_mysql();
        cursor=conn.cursor();
        try:
            sql=f'select goods_id,goods_type,goods_name,goods_name2,goods_producer from products where goods_id=%s'
            print(data['product_id'])
            cursor.execute(sql,data['product_id'])
            product_info=cursor.fetchall()
            print(product_info)
            goods_id=data['product_id']
            goods_type=product_info[0][1]
            goods_name=product_info[0][2]
            cmt=data['cmt'] 
            total_price=data['total_price']
            flow_in_time=data['date_time']
            opr_name=data['opr_name']
            goods_unit=data['product_unit']
            current_time=datetime.now()
            opr_time=current_time.strftime('%Y-%m-%d %H:%M:%S')
            print(current_time,opr_time)

            print(goods_unit)
            if goods_unit=='g':
                # goods_unit=data['product_unit']
                goods_unit2=goods_unit
                goods_price=data['price']
                qty2=qty=data['qty']
        
            else:
                equal_g=data['equal_g']
                qty2=data['qty']
                qty=float(qty2)*float(equal_g)

                goods_price=float(total_price)/float(qty);
                goods_unit2=data['product_unit']
                goods_unit='g'

            values=(goods_id,goods_type,goods_name,goods_unit,qty,goods_price,total_price,goods_unit2,qty2,cmt,flow_in_time,opr_name,opr_time)
            print(values)

            print('qty2:',qty2)

            sql=f'''
                 insert into flow_purchase (goods_id,goods_type,goods_name,goods_unit,qty,goods_price,total_price,goods_unit2,qty2,cmt,flow_in_time,opr_name,opr_time) 
                 values 
                 (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            '''
            cursor.execute(sql,values)
            conn.commit()
            cursor.close()
            conn.close()     
            return  jsonify({'res':'ok','prompt':(flow_in_time,goods_name,qty2,goods_unit2,opr_name)})    
            
        except Exception as e:
            print(e)
            cursor.close()
            conn.close()
            return jsonify({'res':'failed'})



        return {'res':'ok'}

    def ktt_sale_deal(self):
        data=request.json
        # print(data)
        conn=self.connect_mysql()
        cursor=conn.cursor()
        try:
            sql=f'select spec_id,spec_unit,spec_name,goods_type,goods_type2,pkg_list from specs where spec_id=%s'
            # print(data['spec_id'])
            cursor.execute(sql,data['spec_id'])
            spec_info=cursor.fetchone()
            ####################################
            # 写入flow_sale表中，即销售流水，按销售规格记录：
            # id, goods_type, spec_id, spec_name, spec_unit, qty, spec_price, total_spec_price, use_for, cmt, sale_time, opr_name, opr_time
            # print(spec_info)
            goods_type=spec_info[3]
            spec_id=spec_info[0]
            spec_name=spec_info[2]
            spec_unit=spec_info[1]
            qty=data['qty']
            spec_price=data['price']
            total_price=data['total_price']
            use_for=data['use_for']
            cmt=data['cmt']
            sale_time=data['date_time']
            opr_name=data['opr_name']
            current_time=datetime.now()
            opr_time=current_time.strftime('%Y-%m-%d %H:%M:%S')

            values=(goods_type, spec_id, spec_name, spec_unit, qty, spec_price, total_price, use_for, cmt, sale_time, opr_name, opr_time)

            sql=f'''
                insert into flow_sale (goods_type, spec_id, spec_name, spec_unit, qty, spec_price, total_spec_price, use_for, cmt, sale_time, opr_name, opr_time) 
                values 
                (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            '''
            cursor.execute(sql,values)
            # conn.commit()

            ####################################
            # 写入flow_sale_product表中，也是销售流水，按进货的规格记录（通常为克（g)）
            # id, goods_id, goods_name, pkg_name, goods_unit, goods_qty, use_for, cmt, sale_time, opr_name, opr_time
            pkg_list=spec_info[5]
            pkgs=pkg_list.split(';')[:-1]
            for pkg in pkgs:
                
                pkg_id=pkg.split(',')[0]
                pkg_qty=float(pkg.split(',')[1])
                sql=f'select goods_list,pkg_name from packages where pkg_id=%s'
                cursor.execute(sql,pkg_id)
                goods_list=cursor.fetchone()
                sql=f'select goods_name,goods_name2,wt_unit from products where goods_id=%s'
                # print(goods_list)
                product_info=goods_list[0].split(';')[0].split(',')
                goods_id=product_info[0]
                # print(goods_id)
                cursor.execute(sql,goods_id)
                goods_info=cursor.fetchone()
                # print(goods_info)
                goods_name=goods_info[0]
                pkg_name=goods_list[1]
                goods_unit=goods_info[2]
                #总克数=每小包克数*小包数*规格数
                goods_qty=float(product_info[1])*pkg_qty*float(qty)
                # use_for
                # cmt
                # sale_time
                # opr_name
                # opr_time
                values=(goods_id,goods_name,pkg_name,goods_unit,goods_qty,use_for,cmt,sale_time,opr_name,opr_time)
                sql=f'''
                    insert into flow_sale_product (goods_id, goods_name, pkg_name, goods_unit, goods_qty, use_for, cmt, sale_time, opr_name, opr_time) 
                    values 
                    (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                '''
                cursor.execute(sql,values)


            conn.commit()
            cursor.close()
            conn.close()

        except Exception as e:
            conn.rollback()
            print(e)
            cursor.close()
            conn.close()
            return jsonify({'res':'failed'})



        return {'res':'ok','prompt':(sale_time,spec_name,qty,spec_unit,opr_name)}
    
    #计算库存
    def ktt_stat_deal(self):
        conn=self.connect_mysql()
        cursor=conn.cursor()
        try:
            sql=f'''
                SELECT 
                p.goods_id, p.goods_name, SUM(p.qty) AS total_purchase_qty, 
                COALESCE((SELECT SUM(s.goods_qty) FROM flow_sale_product s WHERE s.goods_id = p.goods_id), 0) AS total_sale_qty, 
                (SUM(p.qty) - COALESCE((SELECT SUM(s.goods_qty) FROM flow_sale_product s WHERE s.goods_id = p.goods_id), 0)) AS current_stock 
                FROM flow_purchase p 
                GROUP BY p.goods_id, p.goods_name
            '''
            cursor.execute(sql)
            res=cursor.fetchall()
            print(res)            
            return jsonify({'res':'ok','stocks':res})

        except Exception as e:
            conn.rollback()
            print(e)
            cursor.close()
            conn.close()
            return jsonify({'res':'failed'})
        

    def riyun_input_page(self):
        return render_template('/riyun_input.html')

    def riyun(self):
        return render_template('/riyun.html')

    def riyun_menu(self):
        return render_template('/riyun_menu.html')

    def ktt(self):
        return render_template('/ktt.html')

    #获取商品规格
    def ktt_fetch_specs(self):
        print('connecting database specs...')
        conn=self.connect_mysql();
        cursor=conn.cursor();
        try: 
            #如果没有记录，写入salt值，如有，更新。
            sql=f'select spec_id,spec_unit,spec_name,goods_type,goods_type2,pkg_list,price,price2,price2_on_pkg,price_cmt from specs'
            cursor.execute(sql)
            res=cursor.fetchall()
            print(res)
            cursor.close()
            conn.close()
            return jsonify({'res':'ok','specs':res})   
        except Exception as e:
            print(e)
            cursor.close()
            conn.close()
            return jsonify({'res':'failed'})

    #获取人员列表
    def ktt_fetch_opr(self):
        # print('connecting database...')
        conn=self.connect_mysql();
        cursor=conn.cursor();
        try: 
            #如果没有记录，写入salt值，如有，更新。
            sql=f'select opr_id,opr_name from opr'
            cursor.execute(sql)
            res=cursor.fetchall()
            print(res)
            cursor.close()
            conn.close()
            return jsonify({'res':'ok','opr':res})   
        except Exception as e:
            print(e)
            cursor.close()
            conn.close()
            return jsonify({'res':'failed'})

    #获取包装列表
    def ktt_fetch_pkgs(self):
        print('connecting database...')
        conn=self.connect_mysql();
        cursor=conn.cursor();
        try: 
            sql=f'select pkg_id,pkg_name,goods_list from packages'
            cursor.execute(sql)
            res=cursor.fetchall()
            # print(res)
            cursor.close()
            conn.close()
            return jsonify({'res':'ok','pkgs':res})   
        except Exception as e:
            print(e)
            cursor.close()
            conn.close()
            return jsonify({'res':'failed'})

    #获取产品列表
    def ktt_fetch_product(self):
        print('connecting database...')
        conn=self.connect_mysql();
        cursor=conn.cursor();
        try: 
            sql=f'select goods_id,goods_name,goods_producer from products'
            cursor.execute(sql)
            res=cursor.fetchall()
            print(res)
            cursor.close()
            conn.close()
            return jsonify({'res':'ok','products':res})   
        except Exception as e:
            print(e)
            cursor.close()
            conn.close()
            return jsonify({'res':'failed'})

    def ktt_buy_list_page(self):
        # print(list(self.config_ktt_order.keys()))
        #读取config文件里的导单模板设置，并传送到前端
        # with open(self.ktt_config['ktt_col_map_config'],'r',encoding='utf-8') as f_colmap:
        #     colmap_config=json.loads(f_colmap.read())
        # # print(self.fn_info)
        
        #读取config文件里的导单模板设置，并传送到前端
        return render_template('/ktt_buy_list.html',expMode=list(self.config_ktt_order.keys()),
                                                    specName0=self.spec0,specName1=self.spec1,
                                                    senderNameDefault=self.sender_name,
                                                    senderTelDefault=self.sender_tel,
                                                    fnInfoDefault=self.fn_info)
    def ktt_return_spec(self):
        data=request.json
        supplier=data['supplier']
        specs=self.col_map_config[supplier]
        return {'res':'ok','data':specs}     
                                                    

    def ktt_deal_list(self):
        data=request.json
        try:
            fn_info=data['fn_info'].split('，')
        except:
            fn_info=data['fn_info'].split(',')

        exp_mode=data['exp_mode']
        sender_name=data['sender_name']
        sender_tel=data['sender_tel']
        spec0=data['spec0']
        buy_list0=data['buy_list0']
        spec1=data['spec1']
        buy_list1=data['buy_list1']
        spec2=data['spec2']
        buy_list2=data['buy_list2']
        # odrs=[[spec0,buy_list0],[spec1,buy_list1]]
        odrs=[]
        if spec0 and buy_list0:
            odrs.append([spec0,buy_list0])
        if spec1 and buy_list1:
            odrs.append([spec1,buy_list1])
        if spec2 and buy_list2:
            odrs.append([spec2,buy_list2])
    

        p=ktt_order_export.KttList()
        res=p.multi_spec_output(supplier=exp_mode,sender_name=sender_name,sender_tel=sender_tel,odrs=odrs,save='yes',save_cfg=fn_info,save_dir=self.ktt_config['save_dir'])
        return jsonify({'res':'ok','wx_txt':res['wx_txt']})

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

    def zip_and_download_ktt_exp_order(self):
        # prd_input=request.data.decode('utf-8')
        tmp_fn=os.path.join('e:\\temp\\ktt\\exp_order','newfn.tmp')
        with open (tmp_fn, 'r') as f:
            newfn=f.read()
        return send_file(os.path.join('e:\\temp\\ktt\\exp_order',newfn),as_attachment=True)

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
