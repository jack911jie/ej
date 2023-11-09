import os
import sys
import json
import re
from datetime import datetime
from datetime import timedelta
from itertools import combinations
import copy
import pandas as pd
import openpyxl
from openpyxl.styles import Font
import logging
 
logging.basicConfig(level=logging.WARNING, format=' %(levelname)s - %(lineno)d 行- %(message)s')
logger = logging.getLogger(__name__)
        
class GanZhi:
    def __init__(self):
        self.jqnames=["小寒", "大寒","立春", "雨水", "惊蛰", "春分", "清明", "谷雨", \
                        "立夏", "小满", "芒种", "夏至", "小暑", "大暑", \
                        "立秋", "处暑", "白露", "秋分", "寒露", "霜降", \
                        "立冬", "小雪", "大雪", "冬至"]
        self.tg=['甲','乙','丙','丁','戊','己','庚','辛','壬','癸']
        self.dz=['子','丑','寅','卯','辰','巳','午','未','申','酉','戌','亥']
        self.jieling_names=["立春",  "惊蛰", "清明",  \
                        "立夏", "芒种","小暑",  \
                        "立秋", "白露", "寒露", \
                        "立冬",  "大雪", "小寒"]
        self.tg_wx={'甲':1,'乙':-1,'丙':2,'丁':-2,'戊':3,'己':-3,'庚':4,'辛':-4,'壬':5,'癸':-5}
        self.dz_wx_yy={'子':5,'丑':-3,'寅':1,'卯':-1,'辰':3,'巳':-2,'午':2,'未':-3,'申':4,'酉':-4,'戌':3,'亥':-5}
        self.dz_wx_ss={'子':-5,'丑':-3,'寅':1,'卯':-1,'辰':3,'巳':2,'午':-2,'未':-3,'申':4,'酉':-4,'戌':3,'亥':5}
        self.data_dz_chong=['子午','丑未','寅申','卯酉','辰戌','巳亥']
        self.data_tg_he={'甲己':'土','乙庚':'金','丙辛':'水','丁壬':'木','戊癸':'火'}
        self.data_dz_sanhe={'申子辰':'水','亥卯未':'木','寅午戌':'火','巳酉丑':'金'}
        self.data_dz_liuhe={'子丑':'土','寅亥':'木','卯戌':'火','辰酉':'金','巳申':'水','午未':'土'}
        self.data_dz_banhe={'申子':'水','子辰':'水','亥卯':'木','卯未':'木','寅午':'火','午戌':'火','巳酉':'金','酉丑':'金'}
        self.data_dz_gong={'申辰':'水','巳丑':'金','寅戌':'火','亥未':'木'}
        self.data_dz_hui={'亥子丑':'水','寅卯辰':'木','巳午未':'火','申酉戌':'金'}
        self.data_dz_po=['子酉','丑辰','寅亥','卯午','未戌','申巳']
        self.data_dz_chuan=['子未','丑午','卯辰','酉戌','申亥','寅巳']
        self.data_dz_xing={'子卯':'无礼之刑','寅巳申':'无恩之刑','丑戌未':'恃势之刑','亥亥':'自刑','酉酉':'自刑','午午':'自刑','辰辰':'自刑'}
        self.shishen_digit={
            '-1':'比肩',
            '1':'劫财',
            '-2':'食神',
            '2':'伤官',
            '-3':'偏财',
            '3':'正财',
            '-4':'七杀',
            '4':'正官',
            '-5':'偏印',
            '5':'正印',
        }
        self.ss_geju={
            "食伤生财":[[-2,3],[-2,-3],[2,3],[2,-3]],
            "食伤架杀":[[2,-4],[-2,-4]],
            "伤官见官":[[2,4]],
            "杀印相生":[[-4,5]],
            "财破印":[[3,5]],
            "印制食伤":[[-2,5],[2,-5],[-2,5],[-2,-5]],
            "财滋杀":[[3,-4],[-3,-4]]
        }
        
        with open(os.path.join(os.path.dirname(__file__),'1800-2100jieqi.json'),'r',encoding='utf-8',errors='ignore') as f: 
            # for line in f.readlines():
            #     self.dic = json.loads(line)
            self.dic=json.load(f)
        
        with open(os.path.join(os.path.dirname(__file__),'ping_and_real_sun_time_adjust.json'),'r',encoding='utf-8',errors='ignore') as f_suntime: 
            self.realsuntime_adjust=json.load(f_suntime)

    
    def only_tg(self,bz):
        bz=bz[::2]
        bz.pop(2)
        return bz
    
    def only_dz(self,bz):
        bz=bz[1::2]
        return bz

    def rel_tg_dz(self,pairs,act='tg_he'):    
        data=eval('self.data_'+act)
        if act=='tg_he':            
            items=2
            link='合化'
        elif act=='dz_chong':
            items=2
            link='冲'
        elif act=='dz_sanhe':
            items=3
            link='合化'
        elif act=='dz_liuhe':
            items=2
            link='合化'
        elif act=='dz_banhe':
            items=2
            link='半合'
        elif act=='dz_gong':
            items=2
            link='拱'
        elif act=='dz_hui':
            items=2
            link='会'
        elif act=='dz_po':
            items=2
            link='破'
        elif act=='dz_chuan':
            items=2
            link='穿（害）'
        elif act=='dz_xing':
            items=2
            link=''


        try:
            std_pairs=[[sorted(list(cb)),cb+link+data[cb]] for cb in list(data)]
        except(TypeError):
            std_pairs=[[sorted(list(cb)),cb+link] for cb in list(data)]

        odr_pairs=[sorted(x) for x in list(combinations(pairs,items))]

        res_he=[]
        for odr_pair in odr_pairs:
            for std_pair in std_pairs:
                if odr_pair==std_pair[0]:
                    if std_pair[1] not in res_he:
                        res_he.append(std_pair[1])

        return res_he

    def rels(self,pairs):
        acts_1=['tg_he','dz_chong','dz_liuhe','dz_hui','dz_po','dz_chuan','dz_xing']
        # acts_2=['dz_sanhe','dz_banhe','dz_gong']

        ress=[]
        for act in acts_1:           
            res=self.rel_tg_dz(pairs,act=act)
            if res:
                ress.append(res)

        res_sanhe=self.rel_tg_dz(pairs,act='dz_sanhe')
        if res_sanhe:
            ress.append(res_sanhe)

        res_banhe=self.rel_tg_dz(pairs,act='dz_banhe')
        if res_banhe:
            if res_sanhe:
                std_dz=''.join([itm[:3] for itm in res_sanhe])
                list_banhe=[]
                for _res_banhe in res_banhe:
                    if _res_banhe[0] in std_dz:
                        pass
                    else:
                        list_banhe.append(_res_banhe)
                if list_banhe:
                    ress.append(list_banhe)
            else:
                ress.append(res_banhe)

        res_gong=self.rel_tg_dz(pairs,act='dz_gong')
        if res_gong:
            if res_sanhe:
                std_dz=''.join([itm[:3] for itm in res_sanhe])
                list_gong=[]
                for _res_gong in res_gong:
                    if _res_gong[0] in std_dz:
                        pass
                    else:
                        list_gong.append(_res_gong)
                if list_gong:
                    ress.append(list_gong)
            else:
                ress.append(res_gong)

        return ress
       

    def shishen(self,pairs):
        # print(pairs)
        wx=copy.deepcopy(self.tg_wx)
        wx.update(self.dz_wx_ss)
        # print(wx,self.tg_wx)
        pairs_wx=[wx[itm] for itm in pairs]
        rz_yy=1 if pairs_wx[4]>0 else -1
        pairs_ss=[abs(itm)*-1 if itm*rz_yy>0 else abs(itm)*1 for itm in pairs_wx]
        _tmp=abs(pairs_ss[4])-1
        fh=[1 if itm>0 else -1 for itm in pairs_ss]
        pairs_abs_ss=[abs(itm)-_tmp if abs(itm)-_tmp>0 else abs(itm)-_tmp+5 for itm in pairs_ss]

        #将八字转成标准的数字格式（带正负，日元为1或-1）
        ss=[itm*fh[id] for id,itm in enumerate(pairs_abs_ss)]
        ss_txt=['（'+self.shishen_digit[str(itm)]+'）' for itm in ss]

        # print(ss,ss_txt)

        ss_txt[4]='（日主）'

        # print(ss,'\n',ss_txt)

        #去除日主
        ss_drop_rz=copy.deepcopy(ss)
        ss_drop_rz.pop(4)

        #去除日主后生成两两组合
        ss_combs=list(combinations(ss_drop_rz,2))

        std_gejus=list(self.ss_geju)


        gejus=[]
        for ss_comb in ss_combs:
            # ss_comb=[2,3]
            for std_geju in std_gejus:
                std_geju_combs=[sorted(itm) for itm in self.ss_geju[std_geju]]
                if sorted(ss_comb) in std_geju_combs:
                    if std_geju not in gejus:
                        gejus.append(std_geju)

        return {'ss_txt':ss_txt,'geju':gejus }

    
    def inputdate(self,y,m,d,h=0,min=0,zishi=0):
        self.jqlist=self.dic[str(y)]
        
        ymdhm=str(y)+"-"+ str(m)+ '-' +str(d)+'-'+str(h)+'-'+str(min)

        self.jqdate=[]
        for i in self.jqnames:
            ptn=r'\d+年.+日'
            ptn_time=r'\d\d:\d\d:\d\d'
            ptn_lunar=r'农历.+'
            tmp_jq=re.search(ptn,self.jqlist[i])
            tmp_jq_time=re.search(ptn_time,self.jqlist[i])
            tmp_jq_lunar=re.search(ptn_lunar,self.jqlist[i])
            self.jqdate.append([i,tmp_jq[0],tmp_jq_time[0],tmp_jq_lunar[0]])
            
        logger.debug(self.jqdate) #当年的节气时间总表,数组形式
            
        for i in self.jqdate:
            i[1]=i[1].replace('年','-').replace('月','-')[0:-1]
            
        
        input_time=datetime.strptime(ymdhm, '%Y-%m-%d-%H-%M')
        self.jq_section=[]
        n=0

        for i in self.jqdate:
            list_time=datetime.strptime(i[1]+'-'+i[2][0:5], '%Y-%m-%d-%H:%M')
            if list_time>input_time:
                self.jq_section.append(n-1)
                self.jq_section.append(self.jqdate[n-1][0])
                self.jq_section.append(self.jqdate[n][0])
                break            
            n+=1
        
        # print(self.jq_section)
        if self.jq_section:
            pass
        else:
            # self.jq_section[0]=24+self.jq_section[0]
            # self.jq_section[1]=self.jqdate[self.jq_section[0]][0]
            # self.jq_section[2]=self.jqdate[self.jq_section[0]-24+1][0]
            self.jq_section.append(23)
            self.jq_section.append('冬至')
            self.jq_section.append('小寒')

            
        if (self.jq_section[0])%2==0:
            m_odr=self.jq_section[0]/2
        else:
            m_odr=(self.jq_section[0]-1)/2
            
        m_odr=m_odr if m_odr>0 else m_odr+12
            
            
        self.jq_section.append(m_odr)
        
        logger.debug(['输入的日期在今年的节气分段：',self.jq_section])  #[节气序数，节气，节气，月份的序号]

        return {'jq_section':self.jq_section}

    def dayun(self,y,m,d,h,min,sex='m',zishi=0,real_sun_time='no',longtitude=120,cal_mode='old'):
        ymdhm=str(y)+"-"+ str(m)+ '-' +str(d)+' '+str(h)+':'+str(min)
        bazi=self.cal_dateGZ(y=y,m=m,d=d,h=h,min=min,zishi=zishi,real_sun_time=real_sun_time,longtitude=longtitude)['bazi']
        jq_section=self.inputdate(y=y,m=m,d=d,h=h,min=0,zishi=0)
        # print(self.jq_section[0])
        if self.jq_section[0]<0:
            self.jq_section[0]=self.jq_section[0]+24
        elif self.jq_section[0]>23:
            self.jq_section[0]=self.jq_section[0]-24
         #计算生日的下一节令（不是节气！），计算大运须用到。        

        # print(jq_section,self.jq_section)
        #阳年生男，阴年生女：

        real_y=y
        if bazi[0] in ['甲','丙','戊','庚','壬'] and sex=='m' or bazi[0] in ['乙','丁','己','辛','癸'] and sex=='f':    
  
            #偶数，即节
            if self.jq_section[0]//2==self.jq_section[0]/2:
                if self.jq_section[0]<22:
                    next_jieling=self.jqnames[self.jq_section[0]+2]
                else:
                    next_jieling=self.jqnames[self.jq_section[0]+2-24]
                    real_y+=1
            #奇数，即气
            else:
                if self.jq_section[0]<23:
                    next_jieling=self.jqnames[self.jq_section[0]+1]
                else:
                    next_jieling=self.jqnames[self.jq_section[0]+1-24]
                    if m==12: #12月份的小寒节气，推到下一年的1月
                        real_y+=1

            next_jl_time=' '.join(self.dic[str(real_y)][next_jieling].split(' ')[0:2]).replace('年','-').replace('月','-').replace('日','')      

            time_delta=abs(datetime.strptime(next_jl_time[:-3],'%Y-%m-%d %H:%M')-datetime.strptime(ymdhm,'%Y-%m-%d %H:%M'))

            if cal_mode=='new':
                start_dayun=self.dayun_days_transfer_new(start_date=datetime.strptime(ymdhm,'%Y-%m-%d %H:%M'),delta=time_delta)
                dif_ymdh='no dif days info'
            elif cal_mode=='old':
                start_dayun=self.dayun_days_transfer_old(start_date=datetime.strptime(ymdhm,'%Y-%m-%d %H:%M'),delta=time_delta)
                dif_ymdh=start_dayun[1]
                start_dayun=start_dayun[0]

            start_dayun_txt=start_dayun.strftime('%Y-%m-%d %H:%M')
            
            
            dayun=[start_dayun_txt,start_dayun.year-y]

            bz_yue_tg,bz_yue_dz=self.tg.index(bazi[2]),self.dz.index(bazi[3])


            dy_list=[]
            dy_year=[]
            dy_y=start_dayun.year
            for odr in range(8): #八步大运
                bz_yue_tg+=1
                if bz_yue_tg>9:
                    bz_yue_tg-=10
                dy_list.append(self.tg[bz_yue_tg])

                bz_yue_dz+=1
                if bz_yue_dz>11:
                    bz_yue_dz-=12
                dy_list.append(self.dz[bz_yue_dz])                
                dy_year.append(dy_y)
                dy_y+=10


            # print(dy_list,dy_year)
            
        #阴生年男，阳年生女:
        elif bazi[0] in ['甲','丙','戊','庚','壬'] and sex=='f' or bazi[0] in ['乙','丁','己','辛','癸'] and sex=='m':
             #偶数，即节           
            if self.jq_section[0]//2==self.jq_section[0]/2:
                if self.jq_section[0]==0 and m==1:
                    # next_jieling=self.jqnames[self.jq_section[0]-2+24]
                    pass
                else:
                    pass
                    # next_jieling=self.jqnames[self.jq_section[0]-2]
                next_jieling=self.jqnames[self.jq_section[0]]
            #奇数，即气
            else:
                if self.jq_section[0]==23 and m==1:
                    # next_jieling=self.jqnames[self.jq_section[0]-1]
                    real_y-=1
                else:
                    pass
                    # next_jieling=self.jqnames[self.jq_section[0]-1+24]
                next_jieling=self.jqnames[self.jq_section[0]-1]

            next_jl_time=' '.join(self.dic[str(real_y)][next_jieling].split(' ')[0:2]).replace('年','-').replace('月','-').replace('日','')  
        
            #如生日与节令同日，则按时辰计。       

            time_delta=abs(datetime.strptime(ymdhm,'%Y-%m-%d %H:%M')-datetime.strptime(next_jl_time[:-3],'%Y-%m-%d %H:%M'))
            
            # start_dayun=datetime.strptime(ymdhm,'%Y-%m-%d %H:%M')+timedelta(seconds=time_delta.total_seconds()*120)
            if cal_mode=='new':
                start_dayun=self.dayun_days_transfer_new(start_date=datetime.strptime(ymdhm,'%Y-%m-%d %H:%M'),delta=time_delta)
                dif_ymdh='no dif days info'
            elif cal_mode=='old':
                start_dayun=self.dayun_days_transfer_old(start_date=datetime.strptime(ymdhm,'%Y-%m-%d %H:%M'),delta=time_delta)
                dif_ymdh=start_dayun[1]
                start_dayun=start_dayun[0]
            start_dayun_txt=start_dayun.strftime('%Y-%m-%d %H:%M')
            
            dayun=[start_dayun_txt,start_dayun.year-y]


            bz_yue_tg,bz_yue_dz=self.tg.index(bazi[2]),self.dz.index(bazi[3])
            dy_list=[]
            dy_year=[]
            dy_y=start_dayun.year
            for odr in range(8): #八步大运
                bz_yue_tg-=1
                # if bz_yue_tg<9:zzxxx
                #     bz_yue_tg+=10
                dy_list.append(self.tg[bz_yue_tg])

                bz_yue_dz-=1
                # if bz_yue_dz>11:
                #     bz_yue_dz-=12

                dy_list.append(self.dz[bz_yue_dz])                
                dy_year.append(dy_y)
                dy_y+=10


            # print(dy_list,dy_year)
        
        else:
            print('无此性别及年份的组合。')

        
        return {'dayun_date':dayun,'dayun_dis':dif_ymdh,'dayun_list':[dy_list,dy_year],'next_jieling':[next_jieling,next_jl_time],'input_date':ymdhm,'sex':sex}




    def cal_dateGZ(self,y,m,d,h=-1,min=0,zishi=0,real_sun_time='no',longtitude=120):

        txt_input_time=str(y)+'-'+str(m)+'-'+str(d)+' '+str(h).zfill(2)+':'+str(min).zfill(2)
        
        if real_sun_time=='yes':
            new_time=self.real_sun_time_transfer(y,m,d,h,min,long=longtitude)

            y=int(new_time['real_sun_time'].year)
            m=int(new_time['real_sun_time'].month)
            d=int(new_time['real_sun_time'].day)
            h=int(new_time['real_sun_time'].hour)
            min=int(new_time['real_sun_time'].minute)
            

            res_real_sun_time=new_time['real_sun_time'].strftime('%Y-%m-%d %H:%M')
            res_ping_sun_time=new_time['ping_sun_time'].strftime('%Y-%m-%d %H:%M')
            # print('真太阳时：',res_real_sun_time)
        else:
            res_real_sun_time='real sun time not calculate'
            res_ping_sun_time='ping sun time not calculate'
        
        
        # 因为23-0点涉及日期变动，先按输入参数校正
        # zishi==0, 23-0点按下一天的子时处理，默认
        # zishi==1, 23-0点按当天的子时处理
        _date=str(y)+'-'+str(m)+'-'+str(d)
        if h==23 and zishi==0:
            date_correct=datetime.strptime(_date,'%Y-%m-%d')+timedelta(days=1)     
            y=int(date_correct.strftime('%Y'))
            m=int(date_correct.strftime('%m'))
            d=int(date_correct.strftime('%d'))     
            
        dateGZ=self.inputdate(y,m,d,h,min)['jq_section']
        
        logger.info(['按子时校正后的日期：',y,m,d])
        
        #年干支
        odr_yg=int(str(y-3)[-1])-1
        odr_yz=int((y-3)%12)-1

        #月干支
#         odr_mg_0=self.month_gz_adjust(odr_yg)
      
        
        if dateGZ[3]==0:
            realMon=12
        else:
            realMon=dateGZ[3]
        
        # print(dateGZ)

        odr_mg=int(self.gzodr((odr_yg+1)*2+realMon,'g'))%10-1
        odr_mz=int(self.gzodr(dateGZ[3]+1,'z'))  #1月是丑月，列表从子开始，故+1
        
        logger.debug(['月序数：',odr_mg,odr_mz,'未按节气校正后的年干支序数：', \
                      odr_yg,odr_yz,dateGZ[3],realMon,(odr_yg+1)*2+realMon])
        
        if m==1 and dateGZ[3]>10  or m==2 and dateGZ[3]==12: #寅月前，算上一年
            odr_yg=int(str(y-1-3)[-1])-1

            odr_yz=odr_yz-1
            odr_mg=int(self.gzodr((odr_yg+1)*2+realMon,'g'))%10-1

            
        logger.debug(['按节气校正后的年干支序数：',odr_yg,odr_yz,odr_mg,dateGZ[3],realMon])
        
        #日干支
        #先计算这年的元旦干支,已知1900年元旦是 甲戌日
        this_yd=self.cal_yd(y)
        if this_yd[0]!=-1: #如果年份<1900，将返回-1
            days_diff=(datetime.strptime(str(y)+'-'+str(m)+'-'+str(d),'%Y-%m-%d')- \
                                            datetime.strptime(str(y)+'-1-1','%Y-%m-%d')).days
            odr_dg=(this_yd[0]+days_diff)%10
            odr_dz=(this_yd[1]+days_diff)%12
            
        else:
            pass       
        
        
        #时干支        
        if h!=-1:
            if h==23:
                h=0
                
            if h%2==0:  #按十二个时辰来计算
                H=h/2
            else:
                H=(h+1)/2    
                
            if odr_dg==0 or odr_dg==5:
                odr_hg_0=0
            elif odr_dg==1 or odr_dg==6:
                odr_hg_0=2
            elif odr_dg==2 or odr_dg==7:
                odr_hg_0=4
            elif odr_dg==3 or odr_dg==8:
                odr_hg_0=6
            elif odr_dg==4 or odr_dg==9:
                odr_hg_0=8

            odr_hg=int(odr_hg_0+H)
            odr_hz=int(H)
        
            res=[self.tg[self.gzodr(odr_yg,'g')], \
                    self.dz[self.gzodr(odr_yz,'z')], \
                    self.tg[self.gzodr(odr_mg,'g')], \
                    self.dz[self.gzodr(odr_mz,'z')], \
                    self.tg[self.gzodr(odr_dg,'g')], \
                    self.dz[self.gzodr(odr_dz,'z')], \
                    self.tg[self.gzodr(odr_hg,'g')], \
                    self.dz[self.gzodr(odr_hz,'z')] 
                ]
        else:
            res=[self.tg[self.gzodr(odr_yg,'g')], \
                    self.dz[self.gzodr(odr_yz,'z')], \
                    self.tg[self.gzodr(odr_mg,'g')], \
                    self.dz[self.gzodr(odr_mz,'z')], \
                    self.tg[self.gzodr(odr_dg,'g')], \
                    self.dz[self.gzodr(odr_dz,'z')], \
                    '', \
                    ''] 

        logger.debug(['输出结果:',res])
        logger.debug('----------------------------')
#         print(res)

        
        return {'bazi':res,'input_time':txt_input_time,'real_sun_time':res_real_sun_time,'ping_sun_time':res_ping_sun_time}

    def gzodr(self,n,j): #校准超过10或12,或负数的天干地支序数
        if j=='g':
            if n>9:
                n=n-10
            if n<0:
                n=n+10
        elif j=='z':
            if n>11:
                n=n-12
            if n<0:
                n=n+12
        return n
    
    def cal_yd(self,n,o=1899): #先计算这年的元旦干支,已知1900年元旦是 甲戌日
        if n==1900:
            return [0,10]
        elif n>1900:            
            days=(datetime.strptime(str(n)+'-1-1','%Y-%m-%d')-datetime.strptime('1900-1-1','%Y-%m-%d')).days
            odr_tg=self.gzodr(days%10+0,'g')
            odr_dz=self.gzodr(days%12+10,'z')
            logger.debug(['距1900年元旦的天数：',days,'元旦的干支序数：',odr_tg,odr_dz])
            
            return [odr_tg,odr_dz]
        else:
            return [-1,-1]

    def dif_dates(self,n,o):
        d1=datetime.strptime(n,'%Y-%m-%d')
        d2=datetime.strptime(o,'%Y-%m-%d')
        return (d1-d2).days
          
    def real_sun_time_transfer(self,y,m,d,h,min,long=120):
        old=datetime.strptime(str(y)+'-'+str(m)+'-'+str(d)+'-'+str(h)+'-'+str(min),'%Y-%m-%d-%H-%M')
        long_delta=(int(long)-120)*4*60
        #平太阳时和真太阳时的差距
        date_str=str(m).zfill(2)+str(d).zfill(2)
        ping_real_delta=int(self.realsuntime_adjust[date_str])
        

        if long_delta==abs(long_delta):
            new_time=old+timedelta(seconds=abs(long_delta))            
        else:
            new_time=old-timedelta(seconds=abs(long_delta))

        if ping_real_delta==abs(ping_real_delta):
            real_time=new_time+timedelta(seconds=abs(ping_real_delta))
        else:
            real_time=new_time-timedelta(seconds=abs(ping_real_delta))



        return {'ping_sun_time':new_time,'real_sun_time':real_time}

    def dayun_days_transfer_new(self,start_date,delta):
        return start_date+timedelta(seconds=delta.total_seconds()*120)

    def dayun_days_transfer_old(self,start_date,delta):
        days=delta.days
        secs=delta.seconds

        # print(days,secs)

        #按天计算
        n_d=(secs//3600)*5 #剩余秒整数 to 天   1小时相当于5天
        n_d=n_d+(secs%3600)//720 #剩余秒小数 to 天  12分钟即720秒相当于1天
        n_h=(secs%3600)%720//30 #再剩余的秒数 to 小时   30秒相当于1小时

        #按除了天以后剩余的秒数计算
        if n_d>30:
            n_m=n_d//30
            n_d=n_d%30
        else:
            n_m=0
        
        n_y=days//3 #年
        n_m=n_m+days%3*4

        if n_m>12:
            n_y=n_y+n_m//12
            n_m=n_m+n_m%12
        

        dif_ymdh=[n_y,n_m,n_d,n_h]

        new_date=self.cal_new_date(s_date=start_date,dis=dif_ymdh)

        return [new_date,dif_ymdh]

    def cal_new_date(self,s_date,dis):
        y,m,d,h,minute=s_date.year,s_date.month,s_date.day,s_date.hour,s_date.minute

        hh=h+dis[3]
        dd=d+dis[2]
        mm=m+dis[1]
        yy=y+dis[0]

        
        if minute>=60:
            hh=hh+minute//60
            minute=minute%60        
        if hh>=24:
            dd=dd+hh//24  #超过24小时进一天
            hh=hh%24  #余下的小时数
        if dd>30:
            mm=mm+dd//30
            dd=dd%30

        if mm>12:
            if mm%12==0:
                yy=yy+(mm//12-1)
                mm=12
            else:
                yy=yy+mm//12
                mm=mm%12

        # print('new date',str(yy)+'-'+str(mm)+'-'+str(dd)+' '+str(hh)+':'+str(minute))

        new_date=datetime.strptime(str(yy)+'-'+str(mm)+'-'+str(dd)+' '+str(hh)+':'+str(minute),'%Y-%m-%d %H:%M')

        
        return new_date

class LiuYue(GanZhi):

    def yue(self,y=2022):
        st_jl=self.jqnames[::2]
        #原来的list中小寒是第一位，把小寒放到最后一位。
        st_jl.append(st_jl[0])
        st_jl=st_jl[1:]

        months=self.dic[str(y)]
        months1=self.dic[str(y+1)]
        
        time_jl=[]
        for jl in st_jl:
            time_jl.append([jl,months[jl].split(' ')[0]])
        # time_jl.append(months1['立春'])

        #下一年立春时间
        next_xiaohan=months1['小寒'].split(' ')[0]
        next_lichun=months1['立春'].split(' ')[0]
        time_jl[-1]=['小寒',next_xiaohan]
        
        for odr,dates in enumerate(time_jl):
            if odr<11:
                nxt_jl_1day=datetime.strptime(time_jl[odr+1][1],'%Y年%m月%d日')-timedelta(days=1)
                nxt_jl_1day=nxt_jl_1day.strftime('%Y年%m月%d日')
                _temp_month=datetime.strptime(dates[1],'%Y年%m月%d日')+timedelta(days=2) #往后推两天来计算，以确保当时间没跨越时还停留在上一个月。
                month_bz=self.cal_dateGZ(_temp_month.year,_temp_month.month,_temp_month.day,8,0)['bazi']  #按8点来计算             
                dates.append(nxt_jl_1day)
                dates.append([month_bz[2],month_bz[3]])
            else:
                nxt_jl_1day=datetime.strptime(next_lichun,'%Y年%m月%d日')-timedelta(days=1)
                nxt_jl_1day=nxt_jl_1day.strftime('%Y年%m月%d日')
                _temp_month=datetime.strptime(next_xiaohan,'%Y年%m月%d日')+timedelta(days=2)
                month_bz=self.cal_dateGZ(_temp_month.year,_temp_month.month,_temp_month.day,8,0)['bazi']
                dates.append(nxt_jl_1day)
                dates.append([month_bz[2],month_bz[3]])    
            
        return time_jl

    def bz_liuyue(self,which_year,sex,y,m,d,h,min,zishi=0,real_sun_time='no',longtitude=120,dy_mode='old'):
        bazi=self.cal_dateGZ(y=y,m=m,d=d,h=h,min=min,zishi=0,real_sun_time=real_sun_time,longtitude=longtitude)
        dy=self.dayun(y=y,m=m,d=d,h=h,min=min,sex=sex,zishi=zishi,real_sun_time=real_sun_time,longtitude=longtitude,cal_mode=dy_mode)
        months=self.yue(y=which_year)

        # print('bazi:',bazi)
        
        #目前的大运
        for dy_year in dy['dayun_list'][1]:
            if which_year<dy_year:
                #算出是第几步大运
                dy_id=dy['dayun_list'][1].index(dy['dayun_list'][1][dy['dayun_list'][1].index(dy_year)-1])
                break
        dy_tg=dy['dayun_list'][0][dy_id*2]
        dy_dz=dy['dayun_list'][0][dy_id*2+1]
        this_dy_gz=[dy_tg,dy_dz]

        this_year=self.cal_dateGZ(which_year,3,1,8,10) #用今年的随便一个日子（3月1日8时10分）算出今年的干支
        this_year_gz=[this_year['bazi'][0],this_year['bazi'][1]]

        # print(bazi['bazi'],this_dy_gz,this_year_gz)

        this_yr_liuyue=[]
        
        for month in months:
            _tmp_total=[]
            _tmp_total.extend(bazi['bazi'])
            _tmp_total.extend(this_dy_gz)
            _tmp_total.extend(this_year_gz)
            _tmp_total.extend(month[3])
            this_yr_liuyue.append(_tmp_total)
        
    

        return {'bazi':bazi,'dayun':dy,'liuyue':months,'this_year_liuyue':this_yr_liuyue}


    def pai_liu_yue(self,yy,sex,y,m,d,h,min,zishi=0,real_sun_time='no',longtitude=120,dy_mode='old'):
        mybz=self.bz_liuyue(which_year=yy,sex=sex,y=y,m=m,d=d,h=h,min=min,zishi=zishi,real_sun_time=real_sun_time,longtitude=longtitude,dy_mode=dy_mode)
        my_this_year_liuyue=mybz['this_year_liuyue']
        # print(my_this_year_liuyue)
        # print(mybz['liuyue'])
        res=[]

        
        for id,ly in enumerate(mybz['liuyue']):
            res_row=[]
            #月起始日，干支
            res_row.extend(['-'.join([ly[1],ly[2]])])

            #天干注意关系
            tg=self.only_tg(my_this_year_liuyue[id])
            tmp_tg=self.rels(tg)
            tmp_tg_total=['，'.join(itm) for itm in tmp_tg]
            tg_total='，'.join(tmp_tg_total)
            res_row.append(tg_total)

            #地支注意关系
            dz=self.only_dz(my_this_year_liuyue[id])
            tmp_dz=self.rels(dz)
            tmp_dz_total=['，'.join(itm) for itm in tmp_dz]
            dz_total='，'.join(tmp_dz_total)
            res_row.append(dz_total)

            #十神注意
            ss_info=self.shishen(my_this_year_liuyue[id])
            ss='，'.join(ss_info['geju'])
            ss_txt=ss_info['ss_txt']
            res_row.append(ss)

            #流月
            res_row.append(ly[3][0]+ss_txt[12]+'\n'+ly[3][1]+ss_txt[13])

            #流年
            # print(my_this_year_liuyue[id])
            res_row.append(my_this_year_liuyue[id][10]+ss_txt[10]+'\n'+my_this_year_liuyue[id][11]+ss_txt[11])

            #大运
            res_row.append(my_this_year_liuyue[id][8]+ss_txt[8]+'\n'+my_this_year_liuyue[id][9]+ss_txt[9])

            #原局
            res_row.append(my_this_year_liuyue[id][0]+ss_txt[0]+'  '+my_this_year_liuyue[id][2]+ss_txt[2]+'  '+
                            my_this_year_liuyue[id][4]+ss_txt[4]+'  '+my_this_year_liuyue[id][6]+ss_txt[6]+'\n'+
                            my_this_year_liuyue[id][1]+ss_txt[1]+'  '+my_this_year_liuyue[id][3]+ss_txt[3]+'  '+
                            my_this_year_liuyue[id][5]+ss_txt[5]+'  '+my_this_year_liuyue[id][7]+ss_txt[7])

            res.append(res_row)

        # print(res)

        df_liuyue=pd.DataFrame(data=res,columns=['流月起始日','天干注意','地支注意','十神注意','流月','流年','大运','原局'])
        df_liuyue=df_liuyue[['原局','大运','流年','流月','流月起始日','天干注意','地支注意','十神注意']]
        df_liuyue['描述']=''
        df_liuyue['助运颜色']=''
        df_liuyue['头像提示']=''
        # print(df_liuyue)
        return df_liuyue

    def export_liuyue_xlsx(self,cus_name,yy,sex,y,m,d,h,min,zishi=0,real_sun_time='no',birth_place='',longtitude=120,dy_mode='old',cal_from='jan',out_dir='e:/temp/ejj/客户流年',show_mode='save'):
        print('\n正在排月运……',end='')

        cus_bazi=self.cal_dateGZ(y=y,m=m,d=d,h=h,min=min,zishi=0,real_sun_time=real_sun_time,longtitude=longtitude)

        if cal_from=='jan':
            df_y_1=self.pai_liu_yue(yy=yy-1,sex=sex,y=y,m=m,d=d,h=h,min=min,zishi=zishi,real_sun_time=real_sun_time,longtitude=longtitude,dy_mode=dy_mode)
            df_jan=df_y_1.iloc[[11]]
            df_y_this=self.pai_liu_yue(yy=yy,sex=sex,y=y,m=m,d=d,h=h,min=min,zishi=zishi,real_sun_time=real_sun_time,longtitude=longtitude,dy_mode=dy_mode)
            df_feb_dec=df_y_this.iloc[0:,:]
            df_res=pd.concat([df_jan,df_feb_dec])

        else:
            df_res=self.pai_liu_yue(yy=yy,sex=sex,y=y,m=m,d=d,h=h,min=min,zishi=zishi,real_sun_time=real_sun_time,longtitude=longtitude,dy_mode=dy_mode)
        
        if show_mode=='save':
            print('完成\n\n正在保存……',end='')
            save_dir=os.path.join(out_dir,cus_name)
            if not os.path.exists(save_dir):
                os.makedirs(save_dir)

            if sex=='m':
                sex_txt='男'
            elif sex=='f':
                sex_txt='女'
            else:
                sex_txt='男'
            save_name=os.path.join(save_dir,str(yy)+'-'+cus_name+'-'+sex_txt+'-流月.xlsx')
            df_res.to_excel(save_name,sheet_name='流月',index=False)

            #客户信息
            birth_info=[cus_name,sex,y,m,d,h,min,birth_place,'\n平太阳时：'+cus_bazi['ping_sun_time'],'\n真太阳时：'+cus_bazi['real_sun_time']]
            self.xlsx_format(src=save_name,birth_info=birth_info)

            print('完成。文件名：{}'.format(save_name))

            os.startfile(save_dir)
        else:
            print('完成\n\n结果如下：\n\n')
    
    def xlsx_format(self,src,birth_info):
        birthday=str(birth_info[2])+'-'+str(birth_info[3]).zfill(2)+'-'+str(birth_info[4]).zfill(2)+'  '+str(birth_info[5]).zfill(2)+':'+str(birth_info[6]).zfill(2)
        if birth_info[1]=='m':
            sex_txt='男'
        elif birth_info[1]=='f':
            sex_txt='女'
        else:
            sex_txt='男'

        if birth_info[7]:
            txt_a16=','.join([str(birth_info[0]),sex_txt,birthday,birth_info[8],birth_info[9],'\n'+birth_info[7]])
        else:
            txt_a16=','.join([str(birth_info[0]),sex_txt,birthday])

        wb=openpyxl.load_workbook(src)
        sht=wb['流月']

        sht['A15']='客户信息'        
        sht['A16']=txt_a16
        sht['A18']='总体描述'



        #行高
        for row_id in range(2,13):
            sht.row_dimensions[row_id].height=38
        # sht.row_dimensions[13].height=38

        #列宽
        sht.column_dimensions['a'].width=55
        sht.column_dimensions['b'].width=11
        sht.column_dimensions['c'].width=11
        sht.column_dimensions['d'].width=11    
        sht.column_dimensions['e'].width=29
        sht.column_dimensions['f'].width=29
        sht.column_dimensions['g'].width=35
        sht.column_dimensions['h'].width=35
        sht.column_dimensions['i'].width=35
        sht.column_dimensions['j'].width=13
        sht.column_dimensions['k'].width=20

        
        #自动换行
        for rows in sht:
            for cell in rows:
            # print(cell)
                cell.alignment=openpyxl.styles.Alignment(wrap_text=True)

        #A15,A18加粗
        sht['A15'].font = Font(bold=True)
        sht['A18'].font = Font(bold=True)
        # sht['a1'].alignment=openpyxl.styles.Alignment(wrap_text=True)

        wb.save(src)
    


if __name__=='__main__':
    bz=LiuYue()
    bz.export_liuyue_xlsx('测试',2024,'f',1981,2,4,3,15,zishi=0,real_sun_time='yes',birth_place='江苏南京',
                            longtitude=119,dy_mode='old',out_dir='e:/temp/ejj/客户流年',show_mode='save')
    # cal_from: jan：从1月-12月 否则按立春，从2月到次年1月
    # dy_mode:大运的计算模式，
    # new是简单的计算用秒计算，可能不准，
    # old是根据传统的方法：按三天计为一岁行大运，即一天计四个月行大运、一个时辰计十天行大运
    # run_ejj()

    # mybz=bz.bz_liuyue(2023,'m',1980,5,23,2,10)
    # test=res['this_year_liuyue'][0]
    # print(test)

    # my_this_year_liuyue=['庚', '申', '辛', '巳', '丙', '申', '己', '丑', '乙', '酉', '壬', '寅', '壬', '辰']

    # bz.shishen(my_this_year_liuyue)

    # bz.rel_he(pairs=k) 
    # res=[]
    # for ly in mybz['liuyue']:
    #     pass

    # # for liuyue in my_this_year_liuyue:
    # #     tg=bz.only_tg(liuyue)
    # #     tmp_tg=bz.rels(tg)
    # #     res.append(tmp_tg)
    
    # for liuyue in my_this_year_liuyue:
    #     res=bz.shishen(liuyue)

    #     print(res)

    

    # for id,ly in enumerate(mybz['liuyue']):
    #     # print('\n',ly,res[id])
    #     pass

    # bz.pai_liu_yue(2023,'m',1980,5,23,2,10,zishi=0,real_sun_time='no',longtitude=120,dy_mode='old')
    
    # mybz=GanZhi()


    # s_date=datetime.strptime('1980-5-23 2:10','%Y-%m-%d %H:%M')
    # delta=abs(s_date-datetime.strptime('1980-6-5 21:14','%Y-%m-%d %H:%M'))
    # res=mybz.dayun_days_transfer_old(s_date,delta)
    # print(res)
    
    # k=mybz.cal_dateGZ(1984,10,24,23,48,zishi=0,real_sun_time='yes',longtitude=108)

    # k=mybz.inputdate(1992,1,6,8,10)

    # k=mybz.dayun(1980,5,23,2,10,sex='f')
    # print(k)

    # zishi==0, 23-0点按下一天的子时处理，默认
    # zishi==1, 23-0点按当天的子时处理

    # k=mybz.real_sun_time_transfer(1980,5,23,2,10,108)

    # data=[
    #     [1980,8,23,2,10,'m'],
    #     [1980,8,23,2,10,'f'],
    #     [1981,10,21,6,10,'m'],
    #     [1981,10,21,6,10,'f'],
    #     [1980,12,25,2,10,'m'],
    #     [1980,12,25,2,10,'f'],
    #     [1981,12,25,2,10,'m'],
    #     [1981,12,25,2,10,'f'],
    #     [1982,1,4,2,10,'m'],
    #     [1982,1,4,2,10,'f'],
    #     [1981,1,4,2,10,'m'],
    #     [1981,1,4,2,10,'f']
    # ]

    
    # for dat in data:
    #     # print(dat)
    #     k=mybz.dayun(dat[0],dat[1],dat[2],dat[3],dat[4],sex=dat[5])
    #     print(k,'\n')


    