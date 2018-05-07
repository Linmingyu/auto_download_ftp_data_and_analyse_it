#coding=gbk
import ftplib
import os
import time
import shutil
from pandas import Series, DataFrame
import pandas as pd
import numpy as np
import datetime
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import subprocess
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import re
import itchat

#文件名及文件类型
target_file_name = 'AAAAAAA_' #脱敏
target_file_type = '.txt.Z'

#ftp文件存放目录
ftpfilepath = '/临时清单'

#下载文件存放目录
saverootpath = 'D:\\'

#winrar安装目录
winrarpath = 'D:\\winrar\\winrar'

#Excel安装目录
officepath = 'D:\\Microsoft Office\\Office12\\excel.exe'

def get_file_name(ftpnamelist):
	'''
	获取下载的文件名字
	'''
	
	all_target_file_name = []
	
	for item in ftpnamelist:
		if item[:len(target_file_name)] == target_file_name:
			all_target_file_name.append(int(item[len(target_file_name):len(target_file_name) + 8]))

	newestdate = max(all_target_file_name)

	print('当前FTP最新清单为： ' + str(newestdate))
	
	print('\n输入回车获取最新清单，输入其他日期，获取其他日期清单:')
	
	while True:
		get = input()

		if get == '':
			a = str(newestdate)
		elif int(get) in all_target_file_name:
			a = get
		else:
			print('你输入的日期格式错误或不存在该日期清单，请重新输入：')
			continue
			
		downloadfilename = target_file_name + a + target_file_type
		break
	
	return downloadfilename


def down_file_from_ftp():
	'''
	登录FTP下载文件，需要把python目录中的ftplib.py的encoding = "lantin-1"修改为encoding = "gbk"
	'''
	
	#FTP登录信息
	host = 'AAA.AA.AAA.153' #脱敏
	username = 'AAAAAA' #脱敏
	password = 'AAAAAAAAA' #脱敏

	#登录FTP
	ftp = ftplib.FTP(host)
	ftp.login(username, password)
	print('成功登录FTP')
	print(ftp.getwelcome())	

	#更改路径
	ftp.cwd(ftpfilepath)
	current_path = ftp.pwd()
	print('\nFTP当前路径更改为：', current_path)

	#获取路径下所有文件名
	allftpfilename = ftp.nlst()

	#设置需要下载的文件名
	downloadfilename = get_file_name(allftpfilename)

	#设置下载文件的存放目录
	savepath = saverootpath + downloadfilename

	#下载文件到指定目录
	print('正在下载"' + downloadfilename + '"  ...')
	copyfile = open(savepath, 'wb')  
	needfilename = 'RETR ' + downloadfilename 
	ftp.retrbinary(needfilename, copyfile.write)
	
	#下载完毕后隔1秒再打印信息，避免解压时文件正在使用
	time.sleep(1)
	print('文件下载完毕')

	#登出FTP
	ftp.quit() 
	print('\n登出FTP')
	
	#返回下载的文件名及保存的路径
	return downloadfilename, savepath


def unrar_the_file(rarfilename, rarpath):
	'''
	调用系统winrar软件解压文件
	'''

	#解压下载的文件
	print('\n正在解压"' + rarfilename + '"  ...')
	cmd = winrarpath + ' x -ibck "' + rarpath+ '" "' + saverootpath + '"'
	os.system(cmd)#待完善
	time.sleep(1)
	print('解压完毕')
	
	unrarfilename = rarfilename.replace('.Z', '')
	unrarpath = rarpath.replace('.Z', '')
	return unrarfilename, unrarpath


def change_filename_to_english(chinesefilename, chinesepath):
	'''
	将中文文件名修改为英文文件名
	'''
	englishfilename = chinesefilename[len(target_file_name):]
	englishpath = chinesepath[:len(chinesepath)-len(chinesefilename)] + englishfilename
	shutil.move(chinesepath, englishpath)
	
	return englishfilename, englishpath
	

def data_statistic(englishfilename, englishpath):
	'''
	进行数据统计，并输出Excel报表
	'''

	#将数据导入为DataFrame
	df = pd.read_table(englishpath, sep='$', encoding='gbk', low_memory = False)
	print('\n正在进行数据统计...')

	#将日期列格式从String修改为datetime
	timelist_str = pd.Series(df['兑换日期'])
	timelist_datetime = pd.Series(datetime.datetime.strptime(k, '%Y-%m-%d') for k in timelist_str)
	df['兑换日期'] = timelist_datetime

	#变量设置
	districtindex = ['A', 'AA', 'AAA', 'AAA', 'A', 'A', 'A', 'A', 'A', 'A', 'A', 'A', 'A', 'A'] #脱敏
	filedatestr = englishfilename.replace('.txt', '')
	filedatedatime = datetime.datetime.strptime(filedatestr, '%Y%m%d')
	delta = datetime.timedelta(days = 7)
	groupbykeyword = 'AAAA' #脱敏
	sumkeyword = 'AAAA' #脱敏

	#统计全年情况
	yearstatistic_unsort = df.groupby(groupbykeyword).agg({groupbykeyword : 'count', sumkeyword : sum})
	yearstatistic_sort = yearstatistic_unsort.reindex(districtindex,fill_value=0)

	#统计当月情况
	monthformat = lambda x : x.month
	df_month = df[df['兑换日期'].map(monthformat) == filedatedatime.month]
	monthstatistic_unsort = df_month.groupby(groupbykeyword).agg({groupbykeyword : 'count', sumkeyword : sum})
	monthstatistic_sort = monthstatistic_unsort.reindex(districtindex,fill_value=0)

	#统计上周情况
	df_week = df[df['兑换日期'] > (filedatedatime - delta)]
	weekstatistic_unsort = df_week.groupby(groupbykeyword).agg({groupbykeyword : 'count', sumkeyword : sum})
	weekstatistic_sort = weekstatistic_unsort.reindex(districtindex,fill_value=0)

	#合并3个DataFrame、修改列名，以及增加合计行
	week_month_df = pd.merge(weekstatistic_sort, monthstatistic_sort, left_index=True, right_index=True, suffixes=('_week', '_month'))
	week_month_year_df = pd.merge(week_month_df, yearstatistic_sort, left_index=True, right_index=True,)
	week_month_year_df.columns = [['本周','本周','本月','本月','2018年','2018年'], ['AAA（户）', 'AAAA（万元）']*3] #脱敏
	week_month_year_df.loc['合计'] = week_month_year_df.sum()
	print('数据统计完毕')

	#输出excel文件
	print('\n正在输出Excel报表...')
	excelfilepath = englishpath.replace('.txt', '.xlsx')
	week_month_year_df.to_excel(excelfilepath, encoding = 'gbk')
	print('Excel报表输出完毕，存放地址为：' + excelfilepath)
	
	return excelfilepath
	

def adjust_excel_style(excelfilepath):
	'''
	调整Excel报表样式
	'''
	#打开Excel文件
	wb = openpyxl.load_workbook(excelfilepath)
	sheet = wb.active

	print('\n正在调整Excel报表样式...')

	#处理表头
	sheet.insert_rows(0)
	sheet.unmerge_cells('B1:C1')
	sheet.unmerge_cells('D1:E1')
	sheet.unmerge_cells('F1:G1')
	sheet.merge_cells('B2:C2')
	sheet.merge_cells('D2:E2')
	sheet.merge_cells('F2:G2')
	sheet['A1'] = 'AAAA受理情况' #脱敏
	sheet.merge_cells('A1:G1')
	sheet['A2'] = sheet['A4'].value
	sheet.merge_cells('A2:A3')
	sheet.delete_rows(4)
	
	#处理备注
	

	#调整行高
	sheet.row_dimensions[1].height = 40
	sheet.row_dimensions[2].height = 17.4
	sheet.row_dimensions[3].height = 34.8
	sheet.row_dimensions[18].height = 17.4

	for i in range(4,18):
		sheet.row_dimensions[i].height = 15


	#调整列宽
	sheet.column_dimensions['A'].width = 10
	for i in 'BCDEFG':
		sheet.column_dimensions[i].width = 15


	#设置单元格字体、对齐、边框、背景颜色样式
	titleFont = Font(name = '微软雅黑', size = 20, bold = True, italic = False)
	headFont = Font(name = '微软雅黑', size = 12, bold = True, italic = False)
	cellFont = Font(name = '微软雅黑', size = 11, bold = False, italic = False)

	titleAlignment = Alignment(horizontal = 'center', vertical = 'bottom')
	cellAlignment = Alignment(horizontal = 'center', vertical = 'center', wrap_text= True)

	cellBorder = Border(left = Side(border_style = 'thin', color = 'FF000000'),
					right = Side(border_style = 'thin', color = 'FF000000'),
					top = Side(border_style = 'thin', color = 'FF000000'),
					bottom = Side(border_style = 'thin', color = 'FF000000'))

	cellPatternFill = PatternFill(fill_type='solid', start_color='FFFFFF00', end_color='FFFFFF00')

	#将样式应用于标题
	sheet['A1'].font = titleFont
	sheet['A1'].alignment = titleAlignment

	#将样式应用于表头
	for RowofCell in sheet['A2':'G3']:
		for Cell in RowofCell:
			Cell.font = headFont
			Cell.alignment = cellAlignment
			Cell.border = cellBorder
			Cell.fill = cellPatternFill

	#将样式应用于单元格
	for RowofCell in sheet['A4':'G17']:
		for Cell in RowofCell:
			Cell.font = cellFont
			Cell.alignment = cellAlignment
			Cell.border = cellBorder

	#将样式应用于合计行
	for RowofCell in sheet['A18':'G18']:
		for Cell in RowofCell:
			Cell.font = headFont
			Cell.alignment = cellAlignment
			Cell.border = cellBorder
			Cell.fill = cellPatternFill
        
	print('Excel报表样式调整完毕')

	#保存Excel表格
	wb.save(excelfilepath)
	
	return True
	
	
def open_excel(state):
	if state:
		print('\n正在打开Excel报表...')
		subprocess.Popen([officepath, excelfilepath]).wait()
		print('Excel报表已关闭')
		

def add_input_to_list(emaillist, to_list, terminal_staff, Falseemail):
	for item in emaillist:
		if '@' in item:
			to_list.append(item)
		elif item in terminal_staff:
			to_list.append(terminal_staff[item])
		else:
			Falseemail.append(item)		
		
		
def add_email_to_list(inputstr, to_list, terminal_staff):
	'''
	将输入的字符串转换为email列表
	'''
			
	email_list = inputstr.replace(' ', '').split(',')
	Falseemail = []
			
	add_input_to_list(email_list, to_list, terminal_staff, Falseemail)

	while len(Falseemail):
				
		if len(to_list) > 1:
			print('\n以下电子邮箱地址已加入发件人列表：')
			for item in to_list[1:]:
				print(item)
						
		print('您输入的' + ', '.join(Falseemail) + '不属于指定人员，请输入详细邮件地址，')
		print('输入no则不再添加：')
		secondinput = input()
				
		if secondinput == 'no' and to_list == 1:
			print('不发送电子邮件')
			return False
		elif secondinput == 'no':
			break
		else:
			email_list2 = secondinput.replace(' ', '').split(',')
			Falseemail = []
			add_input_to_list(email_list2, to_list, terminal_staff, Falseemail)

		
def send_email(excelfilepath):
	'''
	将Excel报表发送电子邮件给指定人员
	'''

	#设置变量
	my_email_address = 'AAAA@AAAA.cn' #脱敏
	my_pas = 'AAAAAAA' #脱敏
	to_list = [my_email_address,]
	terminal_staff = {'AAA' : 'AAAA@AAAA.cn',
					  'AAAAA' : 'AAAAAA@AAAA.cn',} #脱敏

	
	print('\n是否将Excel报表发送电子邮件给指定人员？')
	print('输入电子邮箱或OA账户或姓名拼音（OA及拼音仅限指定人员），发送多个人员用英文逗号分开，输入no不发送：')
	sendornot = input()
	
	#是否发送邮件
	if sendornot == 'no':
		print('不发送电子邮件')
		return False
	else:
		add_email_to_list(sendornot, to_list, terminal_staff)
		

	#设置邮件标题及正文内容
	msg = MIMEMultipart()
	msg['Subject'] = 'AAAA业务受理情况' + re.sub('\D', '', excelfilepath) #脱敏
	puretext = MIMEText('无正文')
	msg.attach(puretext)

	#设置邮件附件内容
	xlsxpart = MIMEApplication(open(excelfilepath, 'rb').read())
	xlsxpart.add_header('Content-Disposition', 'attachment', filename = excelfilepath)
	msg.attach(xlsxpart)

	#发送电子邮件
	print('\n正在发送电子邮件...')
	smtpObj = smtplib.SMTP_SSL('smtp.chinatelecom.cn', 465)
	smtpObj.ehlo()
	smtpObj.login(my_email_address, my_pas)
	smtpObj.sendmail(my_email_address, to_list, msg.as_string())
	smtpObj.quit()
	
	print('\n已将Excel报表发送至以下邮箱：')
	for item in to_list:
		print(item)
		

def send_excel_by_wechat(excelfilepath):
	'''
	通过微信将Excel报表发给指定人员
	'''

	print('\n是否将Excel报表通过微信发送给指定人员？')
	print('输入yes或no：')
	sendornot = input()
	
	#是否通过微信发送
	if sendornot == 'yes':
		print('\n正在登录微信...')
		itchat.auto_login(hotReload = True)
		print('登录微信成功')
	
		lmy = itchat.search_friends(nickName = '林明煜')
		lmy_username = lmy[0]['UserName']
	
		print('\n正在使用微信发送Excel报表')
		itchat.send_file(excelfilepath, lmy_username)
		itchat.send_msg('以上为' + re.sub('\D', '', excelfilepath) + 'AAA业务受理情况，请查收。', lmy_username)#脱敏
		print('发送完成')
	
		itchat.logout()
		print('登出微信')
	
	
def remove_file(rarfilepath, txtfilepath, excelfilepath):
	'''
	删除文件
	'''
	print('\n是否需要删除源文件？输入y删除，输入其他不删除：')
	delornot = input()
	
	if delornot == 'y':
		os.remove(rarfilepath)
		os.remove(txtfilepath)
		print('\n已删除以下文件：')
		print(rarfilepath)
		print(txtfilepath)
		
		print('\n是否需要删除Excel报表文件？输入y删除，输入其他不删除：')
		exceldelornot = input()
		
		if exceldelornot == 'y':
			os.remove(excelfilepath)
			print('\n已删除以下文件：')
			print(excelfilepath)
	
	print('\n感谢您的使用')


#登录FTP下载数据文件，并返回文件名、保存路径
downloadfilename, savepath = down_file_from_ftp()

#解压文件
unrarfilename, unrarpath = unrar_the_file(downloadfilename, savepath)

#修改中文名为英文名
englishfilename, englishpath = change_filename_to_english(unrarfilename, unrarpath)

#统计数据并输出Excel报表
excelfilepath = data_statistic(englishfilename, englishpath)

#调整Excel报表样式
state = adjust_excel_style(excelfilepath)

#打开Excel报表
open_excel(state)
	
#将Excel报表发送电子邮件给指定人员
send_email(excelfilepath)

#将Excel报表通过微信发送给指定人员
send_excel_by_wechat(excelfilepath)

#删除文件
remove_file(savepath, englishpath, excelfilepath)


