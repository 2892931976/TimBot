# -*- coding:utf-8 -*-
import sys
#QQ超人打码支持类库
import ctypes
from os.path import join, dirname, abspath, exists
dll = ctypes.windll.LoadLibrary(join(dirname(__file__),'dc.dll'))
class dcVerCode:
    #user QQ超人打码账号
    #pwd QQ超人打码密码
    #softId 软件ID 缺省为0,作者务必提交softId,已保证分成
    def __init__(self,user,pwd,softId="0"):
        self.user = user
        self.pwd = pwd
        self.softId = softId

    #获取账号剩余点数
    #成功返回剩余点数
    #返回"-1"----网络错误
    #返回"-5"----账户密码错误

    def getUserInfo(self):
        p = dll.GetUserInfo(self.user,self.pwd)
        if p:
            return ctypes.string_at(p,-1)
        return ''

    #解析返回结果,成功返回(验证码,验证码ID),失败返回错误信息
    #点数不足:Error:No Money!
    #账户密码错误:Error:No Reg!
    #上传失败，参数错误或者网络错误:Error:Put Fail!
    #识别超时:Error:TimeOut!
    #上传无效验证码:Error:empty picture!
    #账户或IP被冻结:Error:Account or Software Bind!
    #软件被冻结:Error:Software Frozen!
    def parseResult(self,result):
        list = result.split('|')
        if len(list)==3:
            return (list[0],list[2])
        return (result,'')

    #recByte 根据图片二进制数据识别验证码,返回验证码,验证码ID
    #buffer 图片二进制数据

    def recByte(self,buffer):
        p = dll.RecByte_A(buffer,len(buffer),self.user,self.pwd,self.softId)
        if p:
            str = ctypes.string_at(p,-1)
            return self.parseResult(str)
        return ''

    #recYZM 根据验证码路径识别,返回验证码,验证码ID
    #path 图片路径
    def recYZM(self,path):
        p = dll.RecYZM_A(path,self.user,self.pwd,self.softId)
        if p:
            str = ctypes.string_at(p,-1)
            return self.parseResult(str)
        return ''

    #reportErr 提交识别错误验证码
    #imageId 验证码ID
    def reportErr(self,imageId):
        dll.ReportError(self.user,imageId)

    #reportErr 提交识别错误验证码
    #返回"-1",提交失败,返回"1",提交成功
    def reportErrA(self,imageId):
        return  dll.ReportError_A(self.user,imageId)


def super_recognize(username, password, file_path):
    client = dcVerCode(username, password, "0")
    yzm, imageId = client.recYZM(file_path)

    return yzm


def super_query_fee(username, password):
    client = dcVerCode(username, password, "0")
    return client.getUserInfo()


if __name__ == '__main__':
    client = dcVerCode("djq1991811","qq5858885","0"); #超人打码帐号,超人打码密码,软件ID
    img = open('code.jpg','rb')
    buffer = img.read()
    img.close()

    #查询帐号余额
    print (client.getUserInfo())

    #按图片字节数据识别
    yzm,imageId = client.recByte(buffer)
    print(yzm,imageId)

    #按图片本地路径识别
    yzm,imageId = client.recYZM("code.jpg")
    print ( yzm,imageId )
    #client.reportErrA(imageId) 只有在验证码识别错误时才运行这个方法,恶意提交将会受到惩罚