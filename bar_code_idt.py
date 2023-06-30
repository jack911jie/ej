import os
import cv2
import pyzbar.pyzbar as pyzbar
import numpy as np
from PIL import Image, ImageDraw, ImageFont
import re

class BarCode:


    def decodeDisplay(self,img_path):        
        # img_data = cv2.imread(decoded_path) #如路径中有中文，则报错。应用下一句：
        img_data=cv2.imdecode(np.fromfile(img_path,dtype=np.uint8),-1)
        # 转为灰度图像
        gray = cv2.cvtColor(img_data, cv2.COLOR_BGR2GRAY)
        

        
        barcodes = pyzbar.decode(gray)

        res=[]
        for barcode in barcodes:

            # 提取条形码的边界框的位置
            # 画出图像中条形码的边界框
            (x, y, w, h) = barcode.rect
            cv2.rectangle(img_data, (x, y), (x + w, y + h), (0, 255, 0), 2)

            # 条形码数据为字节对象，所以如果我们想在输出图像上
            # 画出来，就需要先将它转换成字符串
            barcodeData = barcode.data.decode("utf-8")
            barcodeType = barcode.type

            #不能显示中文
            # 绘出图像上条形码的数据和条形码类型
            #text = "{} ({})".format(barcodeData, barcodeType)
            #cv2.putText(imagex1, text, (x, y - 10), cv2.FONT_HERSHEY_SIMPLEX,5, (0, 0, 125), 2)

            #更换为：
            img_PIL = Image.fromarray(cv2.cvtColor(img_data, cv2.COLOR_BGR2RGB))

            # 参数（字体，默认大小）
            font = ImageFont.truetype('msyh.ttc', 35)
            # 字体颜色（rgb)
            fillColor = (0,255,255)
            # 文字输出位置
            position = (x, y-10)
            # 输出内容                              
            str = barcodeData

            # 需要先把输出的中文字符转换成Unicode编码形式(  str.decode("utf-8)   )

            draw = ImageDraw.Draw(img_PIL)
            draw.text(position, str, font=font, fill=fillColor)
            # 使用PIL中的save方法保存图片到本地
            img_PIL.save('02.jpg', 'jpeg')

            res.append(barcodeData)

            # 向终端打印条形码数据和条形码类型
            # print("扫描结果==》 类别： {0} 内容： {1}".format(barcodeType, barcodeData))
        return res

    


    def batch_identify(self,dir):
        res=[]
        for fn in os.listdir(dir):
            if re.match(r'.*.jpg',fn):
                fn=os.path.join(dir,fn)
                result=self.decodeDisplay(fn)
                res.append(result)
        
        return res

    def exp_res(self,res_list):
        codes=[itm for res in res_list for itm in res]
        res_codes=[]
        for k in codes:
            if re.match(r'SF\d{13}',k):
                if k not in res_codes:
                    res_codes.append(k)
        return res_codes

if __name__ == '__main__':
    p=BarCode()
    # decodeDisplay("e:\\temp\\temp_ejj\\01bar.jpg")
    res=p.batch_identify('E:\\temp\\ejj\\团购群\\快递')
    res=p.exp_res(res)
    print(res)