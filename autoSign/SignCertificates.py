#!c:\Python27\python
# -*- coding:UTF-8 -*-
# signCertificates.py 2012-08-09 13:45 Weiguang Liu

#Auto Sign Certificates 
#Copyright (C) 2012 name of Weiguang Liu

#此程序用于封装signtool工具，可以应用一些规则，并且签名用的证书不需要安装在员工的机器上,员工可以在自己机器上调用服务器上的程序进行签名工作。
'''
Purpose: Call this script to automatically sign all the executables
Usage: python SignCertificates.py -s <SourcePath> -e <Extensions> [-m <MailList>]
Parameters:
  -s   --sourcepath: - specify the path that put all the files for signing.
  -e   --extensions: - specify the file extensions for signing. the extensions 
          are seperated by semicolons (;)
  -m   --mailList  - specify the mail addresses that will receive the sign results.

Authors:weiguang liu
You can contact authors by email at wliu@arcsoft.com.cn

  File Version:
  2012/8/9 Initial version by wliu
  
'''
import sys,os,time
import getopt
import datetime
import re
import subprocess
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication


def main(argv):
    config = Config(argv)
    if config.is_not_good(): return config.usage()
    if config.want_help(): return config.show_help()
    filelist = FileList(config)
    checkpolice = CheckPolice(config)
    signed = Signed(config)
    mail = Mail(config)
    mail_body = mail.get_body()
    is_write = check_network_share(config)
    if is_write != True:
        mail_body = "Can not Write! %s,%s" % (config._source_path,is_write)
    else:
        for f in filelist.get_file_glob_extensions():
            if checkpolice.check_with_arcsoft(f) == True and \
               checkpolice.check_with_signed(f) == False:
                mail_body += mail.get_sign_files(signed.sign_files(f))

    write_log(config,mail.get_mail_list())
    write_log(config,mail_body)
    basemail = BaseMail(config)
    basemail.send(mail.get_mail_title(),mail_body,mail.get_mail_list())
    
class Config:
    '''
    包含所有程序的配置参数
    '''
    def __init__(self,argv):
        '''
        初始化，参数来自于命令行
        '''
        self._argv = argv
        self._broken = False
        self._want_help = False
        self._error_message = None
        self._source_path = None
        self._extensions = None
        self._mail_list = None
        self._getversioninfo = "D:\\AutoSign\\Bin\\GetVersionInfo.exe"
        self._sigcheck = "D:\\AutoSign\\Bin\\sigcheck.exe"
        self._signtool = "D:\\Tools\\signtool.exe"
        self._is_with_arcsoft = "ArcSoft"
        self._timestamp_url = "http://timestamp.verisign.com/scripts/timestamp.dll"
        self._mscv = "D:\Signtool\MSCV-VSClass3.cer"
        self._is_not_check_arcosft_extensions = ".cat;.cab"
        self._smtpserver = "mail.arcsoft.com.cn"
        self._mail_from = "ArcBuilder@arcsoft.com.cn"
        self._mail_title = "Automatic sign files"
        self._log_file = "debug.txt"
        self._test_write_file = "write_me_ok.txt"

        try:
            optlist, args = getopt.getopt(
                argv[1:],
                'hs:e:m:',
                [
                    'help',
                    'sourcepath=',
                    'extensions=',
                    'maillist='
                ])
        except getopt.GetoptError, e:
            self._broken = True
            self._error_message = str(e)
            return None

        optdict = {}

        for k,v in optlist:
            optdict[k] = v

        if optdict.has_key('-h') or optdict.has_key('--help'):
            self._want_help = True
            return None

        for key,value in optlist:
            if key == '-s':self._source_path = value
            elif key =='--sourcepath':self._source_path = value
            elif key =='-e':self._extensions = value
            elif key =='--extensions':self._extensions = value
            elif key =='-m':self._mail_list = value
            elif key =='--maillist':self._mail_list = value

        if self._source_path is None or self._extensions is None:
            self._broken = True
            
        else:
            print 'source_path:',self._source_path
            print 'extensions:',self._extensions
            print 'mail_list:',self._mail_list
        
    def is_not_good(self):
        return self._broken

    def usage(self):
        if self._error_message is not None:print >> sys.stderr,'Error:%s' % self._error_message
        print >>sys.stderr,'Usage: %s' % (self.argv[0])
        print >>sys.stderr,'Use %s --help to get help.' % (self.argv[0])
        return -1

    def want_help(self):
        return self._want_help

    def show_help(self):
        print __doc__
        return None

    def get_source_path(self):
        return self._source_path

    def get_extensions(self):
        return self._extensions
    
        
class FileList:
    def __init__(self,config):
        self._source_path = config.get_source_path()
        self._extensions = config.get_extensions()


    def get_file_glob_extensions(self):
        self._filelist = []
        for root,dirs,files in os.walk(self._source_path):
            for f in files:
                file_ext = os.path.splitext(f)
                if re.search(format_extensions(self._extensions),file_ext[1],re.I):
                    self._filelist.append(root+os.sep+f)
        return self._filelist

class CheckPolice:
    def __init__(self,config):
        self._getversioninfo = config._getversioninfo
        self._sigcheck = config._sigcheck
        self._is_with_arcsoft = config._is_with_arcsoft
        self._is_not_check_arcosft_extensions = config._is_not_check_arcosft_extensions
        self._signed_flag = "Signing date"
        
    def check_with_arcsoft(self,f):
        self._cmd = self._getversioninfo+" \"%s\"" % f
        get_version_info = get_cmd_output(self._cmd)
        file_ext = os.path.splitext(f)
        if re.search(format_extensions(self._is_not_check_arcosft_extensions),file_ext[1],re.I) or \
           self._is_with_arcsoft.lower() in get_version_info.lower():
            self._is_file_created_by_arcsoft = True
        else:
            self._is_file_created_by_arcsoft = False
        return self._is_file_created_by_arcsoft
    
    def check_with_signed(self,f):
        self._cmd = self._sigcheck+" -q \"%s\"" % f
        self._get_signed_info = get_cmd_output(self._cmd)
        if self._signed_flag in self._get_signed_info:
            self._is_with_signed = True
        else:
            self._is_with_signed = False
        return self._is_with_signed


class Signed:
    def __init__(self,config):
        self._signtool = config._signtool
        self._timestamp_url = config._timestamp_url
        self._mscv = config._mscv
        self._successfully_signed = "Number of files successfully Signed: 1"
        
    def sign_files(self,f):
        self._cmd = self._signtool+" sign /v /ac "+self._mscv \
                    + " /s my /n ArcSoft /t " + self._timestamp_url \
                    + " \"%s\"" % f
        #print self._cmd
        self._get_info = get_cmd_output(self._cmd)
        if self._successfully_signed in self._get_info:
            self._is_success = "Successfully signed! %s" % f
        else:
            self._is_success = "Error! %s:%s" % (f,self._get_info)
        return self._is_success

class Mail:
    def __init__(self,config):
        self._mail_title = config._mail_title
        self._extensions = config._extensions
        self._source_path = config._source_path
        self._mail_list = config._mail_list
                
    def get_mail_title(self):
        return self._mail_title

    def get_body(self):
        r = "<br/>"
        r += "Hi,Guys,<br/>"
        r += "The files under %s with extensions {%s} have been signed.<br/>" % (self._source_path,self._extensions)
        r += "<br/>"
        r += "===Signed files list===<br/>"
        return r
        
    def get_sign_files(self,f):
        r = "<br/>"
        r += "%s" % f
        r += "<br/>"
        return r
    
    def get_mail_list(self):
        return self._mail_list.split(';')
            
        
class BaseMail:
    def __init__(self,config):
        self.smtp = config._smtpserver
        self.sender = config._mail_from
        #self.pwd = pwd
        
    def _parserSend(self,sSubject,sContent,lsPlugin):
        return sSubject,sContent,lsPlugin
    
    def send(self,sSubject,sContent,lsTo,lsCc=[],lsPlugin=[]):
        mit = MIMEMultipart()
        mit['from'] = self.sender
        #lsTo += lsCc
        mit['to'] = ','.join(set(lsTo))
        #if lsCc: mit['cc'] = ','.join(lsCc)

        codeSubject, codeContent, codePlugin = self._parserSend(sSubject, sContent, lsPlugin) 
        mit.attach( MIMEText( codeContent, 'html', 'utf-8' ) ) 
        mit['subject'] = codeSubject 
        for plugin in codePlugin: 
            mitFile = MIMEApplication( plugin['content'], ) 
            mitFile.add_header( 'content-disposition', 'attachment', filename=plugin['subject'] ) 
            mit.attach( mitFile ) 
        try:     
            server = smtplib.SMTP()
            server.connect(self.smtp)
            #server.set_debuglevel(smtplib.SMTP.debuglevel) 
            #server.login( self.sender, self.pwd ) 
            server.sendmail( self.sender, lsTo , mit.as_string() ) 
            server.close()
        except Exception ,e:
            print e


def format_extensions(ext_list):
    exts = []
    for ext in ext_list.split(';'):
            exts.append(ext)
    return "|".join(exts)+'\Z'

            
def get_cmd_output(cmd):
    p = subprocess.Popen(cmd,
#                        stdin = subprocess.PIPE,
                         shell = True,
                         stdout = subprocess.PIPE,
                         stderr = subprocess.PIPE
                         )
#    p.stdin.close()
    #print p.stdout.readlines()
    return p.communicate()[0]
#    return p.stdout.readlines()

def write_log(config,log_message):
    f = open(config._log_file,'at')
    print >>f, "--------------------------------------------"
    print >>f, str(datetime.datetime.now()),log_message
    f.close()

def check_network_share(config):
    network_share_flag = config._source_path+os.sep+config._test_write_file
    try:
        f = open(network_share_flag,'at')
    except IOError, e:
        return e
    else:
        print >>f, "write ok!"
        f.close()
        os.remove(network_share_flag)
        return True
    

if __name__ == '__main__':
    main(sys.argv)
