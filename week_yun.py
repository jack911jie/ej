import os
import sys
sys.path.append(os.path.join(os.path.dirname(__file__),'modules'))
import ganzhi
import copy
from PIL import Image,ImageFont,ImageDraw
from datetime import date, datetime,timedelta
import json
import random
import pandas as pd
# pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.unicode.east_asian_width', True)

class WeekYun:
    def __init__(self,work_dir='D:\\工作目录\\ejj'):
        self.work_dir=work_dir
        # print('weekyun:',self.work_dir)

    def read_day_cmt(self,xls='d:\\工作目录\\ejj\\运势\\运势.xlsx'):
        df_dayyun=pd.read_excel(xls,sheet_name='运势',skiprows=1)
        df_dayyun.columns=['日期','星期','干','支','木颜色','木描述','火颜色','火描述','土颜色','土描述','金颜色','金描述','水颜色','水描述']
        df_dayyun['干支']=df_dayyun['干']+df_dayyun['支']
        df_dayyun=df_dayyun[['日期','星期','干支','木颜色','木描述','火颜色','火描述','土颜色','土描述','金颜色','金描述','水颜色','水描述']]
        return df_dayyun

    def read_color_para(self,xls='d:\\工作目录\\ejj\\运势\\运势.xlsx'):
        df_color_para=pd.read_excel(xls,sheet_name='参数表-色系')
        return df_color_para

    def read_dec_para(self,xls='d:\\工作目录\\ejj\\运势\\运势.xlsx'):
        df_dec_para=pd.read_excel(xls,sheet_name='参数表-饰品')
        return df_dec_para

    def day_cmt(self,date_input='20220822',xls='d:\\工作目录\\ejj\\运势\\运势.xlsx'):
        df_cmt=self.read_day_cmt(xls=xls)
        day_cmt=df_cmt[df_cmt['日期']==datetime.strptime(date_input,'%Y%m%d')].copy(deep=True)
        dategz=ganzhi.GanZhi().cal_dateGZ(int(date_input[:4]),int(date_input[4:6]),int(date_input[6:]),8,0)['bazi']
        gz=''.join(dategz)
        day_cmt['日期干支']=gz[:2]+'年'+gz[2:4]+'月'+gz[4:6]+'日'
        return day_cmt

    def wuxing(self,date_input='20220822',wx='木',xls='d:\\工作目录\\ejj\\运势\\运势.xlsx'):
        daycmt=self.day_cmt(date_input=date_input,xls=xls)
        df=pd.DataFrame()
        
        df['日期']=daycmt['日期']
        df['星期']=daycmt['星期']
        df['日期干支']=daycmt['日期干支']
        df['颜色']=daycmt[wx+'颜色'].str.replace(r'[，,、]','',regex=True)
        df['颜色']=df['颜色'].apply(lambda x:''.join(sorted(x)))
        df['描述']=daycmt[wx+'描述']
        df['五行']=wx

        #色系图地址
        # print(df['颜色'].tolist()[0])
        # print(map(sorted,df['颜色'].tolist()[0]))
        color_list=[]
        for color in os.listdir(os.path.join(self.work_dir,'素材','色系图')):
            if color[-3:].lower()=='png':
                clr_name=color.split('_')[0]
                # if ''.join(sorted(clr_name))==list(map(sorted,df['颜色'].tolist()[0]))
                if ''.join(sorted(clr_name))==df['颜色'].tolist()[0]:
                    color_list.append(color)
        
        try:
            _pick_color=random.choice(color_list)
        except IndexError:
            print(color_list)
            print(df['颜色'].tolist()[0]+':目录中没有对应的色系图片'+'\n')
            exit(0)
        pick_color=os.path.join(self.work_dir,'素材','色系图',_pick_color)
        df['颜色图地址']=pick_color

        #色系图描述
        color_txt=self.read_color_para(xls=xls)
        color_txt['排序颜色名']=color_txt['颜色'].apply(lambda x: ''.join(sorted(x)))
        res_colors=color_txt[color_txt['排序颜色名']==df['颜色'].tolist()[0]]
        # print(df['颜色'].tolist()[0],color_txt['排序颜色名'])
        res_color=res_colors.sample(1)
        # print(res_colors,'\n\n',res_color)
        color=res_color['标签'].tolist()[0].replace('，',' ')
        df['颜色标签']=color
        df['实际颜色']=df['颜色'].tolist()[0]

        #饰品图地址
        dec_txt=self.read_dec_para(xls=xls)

        decs=[]
        for clr in df['颜色'].tolist()[0]:
            if clr=='黑':
                clr='蓝'
            if clr=='金':
                clr='白'
            if clr=='粉':
                clr='红'
            if clr=='紫':
                clr='红'
            if clr=='棕':
                clr='黄'
            if clr=='粉红':
                clr='红'
            dec_wxs=dec_txt[dec_txt['颜色']==clr].sample(1)['五行属性'].tolist()[0]
            if dec_wxs not in decs:
                decs.append(dec_wxs)
        
        dec_urls=[]
        for dec_wx in decs:
            for dec_fn in os.listdir(os.path.join(self.work_dir,'素材','饰品图')):
                if dec_fn[0]==dec_wx:
                    dec_urls.append(os.path.join(self.work_dir,'素材','饰品图',dec_fn))

        df['饰品图地址']=','.join(dec_urls)

        return df
        
class ExportImage(WeekYun):
    def __init__(self,work_dir='D:\\工作目录\\ejj'):
        self.work_dir=work_dir
        with open(os.path.join(os.path.dirname(__file__),'configs','fonts.config'),'r',encoding='utf-8') as f:
            self.config=json.load(f)        
    
    def font(self,ft_name,ft_size):
        font_fn=self.config[ft_name].replace('$',self.work_dir)
        return ImageFont.truetype(font_fn,ft_size)

    def font_color(self,txt):
        clist={
            "红":"#DC262B",
            "黄":"#EFF94A",
            "蓝":"#1A3E68",
            "绿":"#267B23",
            "白":"#B3B5B3",
            "黑":"#000000",
            "灰":"#474847",
            "金":"#DDD5A6"
        }

        try:
            return clist[txt]
        except:
            return clist['灰']

    def wx_color(self,wx):
        wxlist={
            '木':'#99A16D',
            '火':'#DB3D29',
            '土':'#AB896E',
            '金':'#EEE1B1',
            '水':'#36475C'
        }
        return wxlist[wx]

    def back_img(self,wx='木'):
        img=Image.open(os.path.join(self.work_dir,'素材','穿搭底图',wx+'.jpg'))
        return img

    def draw_img(self,date_input='20220822',wx='木',xls='d:\\工作目录\\ejj\\运势\\运势.xlsx',exp_txt='no'):
        bg_img=self.back_img(wx=wx)
        infos=self.wuxing(date_input=date_input,wx=wx,xls=xls)
        # print(infos.columns)
    
        #文字部分
        wxtitle=infos['五行'].tolist()[0]+'宝宝'
        weekay='星期'+infos['星期'].tolist()[0]
        date_txt=datetime.strftime(infos['日期'].tolist()[0],'%Y年%m月%d日')
        date_gz=infos['日期干支'].tolist()[0]
        color_txt=infos['颜色标签'].tolist()[0]       

        
        #饰品      

        #随机打乱顺序
        _decs=infos['饰品图地址'].tolist()[0].split(',')
        random.shuffle(_decs)
        
        #去重，去类型相同的饰品，如手镯，只有一个
        decs_drop_dup=[]
        type_dup_list=[]
        for _pre_decs in _decs:
            if _pre_decs.split('_')[2] not in type_dup_list:
                decs_drop_dup.append(_pre_decs)
                type_dup_list.append(_pre_decs.split('_')[2])

        #如饰品数>3，则随机选出3个
        if len(decs_drop_dup)>3:
            decs=random.sample(decs_drop_dup,3)
        else:
            decs=decs_drop_dup

        #饰品名字文字
        dec_txts=[]
        dec_txt_out_list=[]
        for dec in decs:
            dec_txt=''
            _dec_t=dec.split('\\')[-1].split('_')
            dec_txt+=_dec_t[1]+_dec_t[2]
            dec_txts.append([dec_txt,_dec_t[0]])
            dec_txt_out_list.append(dec_txt)
        
        #输出到文本的饰品描述内容
        dec_txt_out='、'.join(dec_txt_out_list)

        #url
        color_url=infos['颜色图地址'].tolist()[0]
        dec_urls=decs
        

        # print(wxtitle,weekay,date_txt,date_gz,color_txt,decs,dec_txts)
        
        #色系图
        x_colorblock,y_colorblock=290,224
        img_color=Image.open(color_url)
        img_color=img_color.resize((298,43))
        
        bg_img.paste(img_color,(x_colorblock,y_colorblock))

       #饰品图 
        x_dec,y_dec,dec_gap=270,454,50
        x_dec_txt_init=copy.deepcopy(x_dec)
        
        for dec_url in dec_urls:            
            img_dec=Image.open(dec_url)            
            img_dec=img_dec.resize((80,80))
            img_dec_a=img_dec.split()[3]
            bg_img.paste(img_dec,(x_dec,y_dec),mask=img_dec_a)
            x_dec+=(img_dec.size[0]+dec_gap)
        
        #标题
        draw=ImageDraw.Draw(bg_img)
        draw.text((120,74),text=wxtitle,font=self.font('汉仪心海行楷W',50),fill=self.wx_color(wxtitle[0]))

        #日期
        draw.text((210,730),text=date_txt+'  '+weekay,font=self.font('字由文艺黑体',24),fill='#72726E') 
        draw.text((220,770),text=date_gz,font=self.font('字由文艺黑体',28),fill='#72726E') 

        #饰品描述
        ft_size_dectxt=20
        for dec_id,dec_txt in enumerate(dec_txts):
            # print(dec_txt)
            #计算饰品描述文字坐标
            x_dec_txt,y_dec_txt=x_dec_txt_init+img_dec.size[0]*dec_id+dec_gap*dec_id+img_dec.size[0]//2-len(dec_txt[0])*ft_size_dectxt//2,y_dec+img_dec.size[1]+10
            draw.text((x_dec_txt,y_dec_txt),dec_txt[0],font=self.font('华康海报体W12(P)',ft_size_dectxt),fill=self.wx_color(dec_txt[1]))

        #色系描述
        # for real_color in infos['实际颜色'].tolist()[0]:
        ft_size_colortxt=20
        x_colortxt,y_colortxt=x_colorblock+img_color.size[0]//2-(len(color_txt)*ft_size_colortxt)//2,y_colorblock+80
        draw.text((x_colortxt,y_colortxt),color_txt,font=self.font('华康海报体W12(P)',ft_size_colortxt),fill='#9E9E9D')

        return bg_img,dec_txt_out

    def batch_deal(self,prd=['20220822','20220828'],out_put_dir='e:\\temp\\ejj\日穿搭',xls='d:\\工作目录\\ejj\\运势\\运势.xlsx'):
        stime,etime=datetime.strptime(prd[0],'%Y%m%d'),datetime.strptime(prd[1],'%Y%m%d')
        datelist=[]
        while stime<=etime:
            datelist.append(stime.strftime('%Y%m%d'))
            stime+=timedelta(days=1)
        
        out_decs_txt=Vividict()
        for nowtime in datelist:
            date_dir=nowtime[:4]+'-'+nowtime[4:6]+'-'+nowtime[6:]
            save_dir=os.path.join(out_put_dir,date_dir)
            if not os.path.exists(save_dir):
                os.makedirs(save_dir)
            for num,wx in enumerate(['木','火','土','金','水']):
                print('正在生成 '+date_dir+' '+wx+'  穿搭')
                res_img,dec_txt=self.draw_img(date_input=nowtime,wx=wx,xls=xls)
                save_name=os.path.join(save_dir,str(num+1)+'-'+wx+'.jpg')
                res_img.save(save_name,quality=95,subsampling=0)
                out_decs_txt[nowtime][wx]=dec_txt

        print('完成')

        return out_decs_txt

class ExportWeekYunTxt(WeekYun):
    def __init__(self,work_dir='D:\\工作目录\\ejj',import_dec_dic=''):
        super(ExportWeekYunTxt,self).__init__(work_dir=work_dir)
        self.import_dec_dic=import_dec_dic

    def wx_icon(self,wx='木'):
        wxchr_list={
            "木":"🌳",
            "火":"🔥",
            "土":"⛰︎",
            "金":"🗡",
            "水":"💧",
            "日历":"📅"
        }

        return wxchr_list[wx]

    def exp_txt(self,date_input='20220822',wx='木',xls='d:\\工作目录\\ejj\\运势\\运势.xlsx'):
        infos=self.wuxing(date_input=date_input,wx=wx,xls=xls)
        #衣服配色语句clr_txt
        clrs=infos['颜色标签'].tolist()[0].split(' ')
        if len(clrs)==1:
            clr_txt='建议穿'+clrs[0]+'衣服，'
        elif len(clrs)==2:
            clr_txt='建议穿'+clrs[0]+'、'+clrs[1]+'衣服，'
        else:
            clr_txt='建议穿'+'、'.join(clrs[:-1])+'以及'+clrs[-1]+'衣服，'

        
        
        #佩戴饰品语句dec_txt
        #无饰品语句导入
        # print('self.import dec dic',self.import_dec_dic)
        if self.import_dec_dic=='':
            decs=infos['饰品图地址'].tolist()[0].split(',')
            dec_names=[x.split('\\')[-1].split('_')[1]+x.split('\\')[-1].split('_')[2] for x in decs]
            for dec_name in dec_names:
                if len(dec_names)==1:
                    dec_txt='佩戴'+dec_names[0]+'等饰品。'
                elif len(dec_names)==2:
                    dec_txt='佩戴'+dec_names[0]+'及'+dec_names[1]+'等饰品。'
                else:
                    dec_txt='佩戴'+'、'.join(dec_names[:-1])+'以及'+clrs[-1]+'的饰品。'
        #有饰品语句导入
        else:
            dec_txt='佩戴'+self.import_dec_dic[date_input][wx]+'以及'+clrs[-1]+'的饰品。'
         
            
        daycmt=self.day_cmt(date_input=date_input,xls=xls)
        df=pd.DataFrame()        
        df['日期']=daycmt['日期']
        df['星期']=daycmt['星期']
        df['日期干支']=daycmt['日期干支']
        df['描述']=daycmt[wx+'描述']
        df['五行']=wx
        wxtitle=wx+'宝宝'

        title=datetime.strftime(df['日期'].tolist()[0],'%Y年%m月%d日')+'（星期'+df['星期'].tolist()[0]+'）运势|穿搭配色\n\n'+self.wx_icon('日历')+ \
                                '  '+df['日期干支'].tolist()[0]+'\n'

        wxtxt=self.wx_icon(wx=wx)+'  '+wxtitle+clr_txt+dec_txt+'\n'+df['描述'].tolist()[0]

        return [title,wxtxt]

    def all_wx_txt(self,date_input='20220822',xls='d:\\工作目录\\ejj\\运势\\运势.xlsx',save='yes',save_dir='e:\\temp\\ejj\\日穿搭'):        
        all_txt=''
        for wx in ['木','火','土','金','水']:
            txts=self.exp_txt(date_input=date_input,wx=wx,xls=xls)
            all_txt+=txts[1]+'\n\n'
        all_txt=txts[0]+'\n'+all_txt

        if save=='yes':
            date_dir=date_input[:4]+'-'+date_input[4:6]+'-'+date_input[6:]
            save_dir=os.path.join(save_dir,date_dir)
            if not os.path.exists(save_dir):
                os.makedirs(save_dir)

            # print(os.path.join(save_dir,date_dir+'_日运.txt'))
            with open(os.path.join(save_dir,date_dir+'_日运.txt'),'w',encoding='utf-8') as f:
                f.write(all_txt)
        
        return all_txt
    
    def all_date_wx(self,prd=['20220822','20220828'],xls='d:\\工作目录\\ejj\\运势\\运势.xlsx',save='yes',save_dir='e:\\temp\\ejj\\日穿搭',import_dec_dic=''):
        stime,etime=datetime.strptime(prd[0],'%Y%m%d'),datetime.strptime(prd[1],'%Y%m%d')
        datelist=[]
        while stime<=etime:
            datelist.append(stime.strftime('%Y%m%d'))
            stime+=timedelta(days=1)
        
        for nowtime in datelist:
            print('正在处理 '+nowtime[:4]+'-'+nowtime[4:6]+'-'+nowtime[6:]+' 穿搭配色文案')
            self.all_wx_txt(date_input=nowtime,xls=xls,save=save,save_dir=save_dir)
        
        print('完成')

class WeekYunCover(ExportImage):
    def __init__(self,work_dir='D:\\工作目录\\ejj'):
        self.work_dir=work_dir
        with open(os.path.join(os.path.dirname(__file__),'configs','fonts.config'),'r',encoding='utf-8') as f:
            self.config=json.load(f)      
    
    def export(self,prd=['20220822','20220828'],save_dir='e:\\temp\\ejj\\日穿搭\\0-每周运势封图'):
        print('正在生成',prd[0]+'-'+prd[1]+'周运封图')
        bg=Image.open(os.path.join(self.work_dir,'素材','周运底图','周运封图竖.jpg'))
        txt=prd[0][:4]+'年'+prd[0][4:6]+'月'+prd[0][6:]+'日'+'-'+prd[1][:4]+'年'+prd[1][4:6]+'月'+prd[1][6:]+'日'
        draw=ImageDraw.Draw(bg)
        draw.text((210,80),text=txt,fill='#968351',font=self.font('字由文艺黑体',20))

        if not os.path.exists(save_dir):
            os.makedirs(save_dir)

        bg.save(os.path.join(save_dir,prd[0]+'-'+prd[1]+'周运封图.jpg'),quality=95,subsampling=0)

        print('完成')

class Vividict(dict):
    def __missing__(self, key):
        value = self[key] = type(self)()
        return value

if __name__=='__main__':
    #######################  一周日穿搭配色文案 + 周运封图   #######################
    # week_txts=ExportWeekYunTxt(work_dir='d:\\工作目录\\ejj')
    # p.all_date_wx(prd=['20220822','20220828'],xls='d:\\工作目录\\ejj\\运势\\运势.xlsx')



    #######################  导出一周日穿搭文案   #######################
    # p=ExportWeekYunTxt()
    # p.all_wx_txt(date_input='20220822',xls='d:\\工作目录\\ejj\\运势\\运势.xlsx',save='yes',save_dir='e:\\temp\\ejj\\日穿搭')
    # p.all_date_wx(prd=['20220822','20220828'],xls='d:\\工作目录\\ejj\\运势\\运势.xlsx',save='yes',save_dir='e:\\temp\\ejj\\日穿搭')


    #######################  导出周运封图   #######################
    # p=WeekYunCover(work_dir='D:\\工作目录\\ejj')
    # p.export(prd=['20220822','20220828'],save_dir='e:\\temp\\ejj\\日穿搭\\0-每周运势封图')





    # p=WeekYun() 
    # df=p.wuxing(date_input='20220822',wx='木',xls='d:\\工作目录\\ejj\\运势\\运势.xlsx')
    # print(df)


    p=ExportImage()
    res=p.draw_img(date_input='20220828',wx='木',xls='d:\\工作目录\\ejj\\运势\\运势.xlsx')
    res.show()
    res=p.batch_deal(prd=['20220912','20220913'],out_put_dir='e:\\temp\\ejj\日穿搭',xls='d:\\工作目录\\ejj\\运势\\运势.xlsx')
    # print(res)
  


    

