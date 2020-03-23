#-*- coding:utf-8 -*-
#new thought: when download it rename it imediatelly. 
import requests
from bs4 import BeautifulSoup as BS
import re, itertools
import pandas as pd 
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC
import smtplib
import imaplib
import email
import os




wait_time=3
email_account='autotaxcode@gmail.com'
password='auto-tax0124'
driver_path='/Users/Mac/Desktop/chromedriver'
download_path='/Users/Mac/Downloads/'
rename_path='/Users/Mac/Desktop/taxcode/'

SMTP_SERVER = "imap.gmail.com"
SMTP_PORT   = 993
initial_url='https://www.wkcheetah.com/#/home/TaxStateAndLocal'
state_xpath_pattern="//*[@id=\"body\"]/div[1]/div/ui-view/div/ui-view/div/ui-view/div/div/div[2]/div/div/div/div/div/div/div/div[2]/ul[1]/span[2]/span/dnd-nodrag/div/ch-content-root-item/div/div/div/ch-content-drop-down-item/div/div[1]/span/div[2]/div/ul/li["
year_xpath_pattern="//*[@id=\"body\"]/div[1]/div/ui-view/div/ui-view/div/ui-view/div/div/div[2]/div/div/div/div/div/div/div/div[2]/ul[1]/span[2]/span/dnd-nodrag/div/ch-content-root-item/div/div/div/ch-content-drop-down-item/div/div[2]/div/div/div/ch-content-drop-down-item/div/div[1]/span/div[2]/div/ul/li["
# set the download path to chrome driver

driver = webdriver.Chrome(driver_path)
wait = WebDriverWait(driver,120)

class autoCHH():

	download_url_list=list()#the url of file download by email
	rename_list=list()
	count_list=[0,0,0,0,0,0,0]
	d_num=0
	total_num=0
	list_year=[2018,2017,2016,2015,2014,2012,2011,2010,2009,2008,2007,2006,2005,2004,2003,2002,2001,2000,1999,1998,1997,1996,1995,1994]
	list_state=['alabama','alaska','arizona','arkansas','california','colorado','connecticut','delaware','district of columbia','florida','georgia','hawaii','idaho','illinois','indiana','iowa','kansas','kentucky','louisiana','maine','maryland','massachusetts','michigan','minnesota','mississippi','missouri','montana','nebraska','nevada','new hamshire','new jersey','new mexico','New York','New York City','north carolina','north dakota','ohio','oklahoma','oregon','pennsylvania','rhode island','south carolina','south dakota','tennessee','texas','utah','vermont','virginia','washington','west virginia','wisconsin','wyoming']
	
	def __init__(self,state_id,year_id,code_str):

		i=0
		print("welcome to autotax download system!")
		print("===================================")
		print("*your email account set for donwload is: "+email_account)
		print("*your file will be donwload to: "+download_path)
		print("*after download, your file should be moved to: "+rename_path)
		print("*your target year is: "+str(self.list_year[year_id[0]-2])+"-"+str(self.list_year[year_id[-1]-2]))
		print("*your target state is: "+ str(self.list_state[state_id[0]-2:state_id[-1]-2]))
		print("*your target code is: ")
		for code in code_str:
			print("["+str(i)+"]"+code)
			i+=1
		self.total_num=int(len(state_id))*int(len(year_id))*int(len(code_str))
		print("*total number of tax code: "+str(self.total_num))
		print("===================================")

							

	def read_content(self,data):

		for response_part in data:
			if isinstance(response_part, tuple):
				msg = email.message_from_string(response_part[1])

				email_subject = msg['subject']
				email_from = msg['from']
				
				print ('From : ' + email_from + '\n')
				print ('Subject : ' + email_subject + '\n')
									
				text=msg.__str__()				
				m = re.findall(r"\<+a+\s+href+\=+3D+\"+http+\:+\/+\/+intelliconn+.+\s+.+\s.+\"",text)# some url is 2 line
				n = re.findall(r"\<+a+\s+href+\=+3D+\"+http+\:+\/+\/+intelliconn+.+\s+.+\s.+\s+.+\"",text)#some url is 3 line
			
				if len(m)>0:
					url_text=m[0]
				if len(n)>0:
					url_text=n[0]

				url=url_text[11:-2]
				temp1_url=url.replace('=','')
				temp2_url=temp1_url.replace(' ','')
				temp3_url=temp2_url.replace('\n','')
				temp4_url=temp3_url.replace('cpid','cpid=')
				temp5_url=temp4_url.replace('userId','userId=')
				temp6_url=temp5_url.replace('jobId','jobId=')

				try:
					driver.get(temp6_url)
					self.download_url_list.append(temp6_url)
				except:
					print ("no such url in this email")
				print("download file from:  "+temp6_url)						
				print("==========")

	
	def read_email_from_gmail(self):
		
		try:
			#login email
			mail = imaplib.IMAP4_SSL(SMTP_SERVER)
			mail.login(email_account,password)
			#open inbox
			mail.select('inbox')
			print("login into email account:"+email_account)
			type, data = mail.search(None, 'ALL')
			mail_ids = data[0]
			id_list = mail_ids.split()   
			#identify the first and the last id
			first_email_id = int(id_list[0])
			latest_email_id = int(id_list[-1])

			if first_email_id!=latest_email_id:
				for i in range(first_email_id,latest_email_id+1):
					typ, data = mail.fetch(i, '(RFC822)' )
					self.read_content(data)
			else:
				typ, data = mail.fetch(1, '(RFC822)' )
				self.read_content(data)
		except Exception, e:
			print (str(e))
		
	def xpath_soup(self,element):

		components = []
		child = element if element.name else element.parent
		for parent in child.parents:
			siblings = parent.find_all(child.name, recursive=False)
			components.append(
				child.name
				if siblings == [child] else
				'%s[%d]' % (child.name, 1 + siblings.index(child))
				)
			child = parent
		components.reverse()
		return '/%s' % '/'.join(components)	

	def find_tag(self,textcode,code_soup):

		tag_find=False
		for tag in code_soup:
			# if tag is found
			if tag.find(text=textcode):
				tag_find=True
				#find tag with tax code and its xpath
				python_button10_soup=tag
				python_button10_xpath=self.xpath_soup(python_button10_soup.previous_element.previous_element)
				#click the tag button
				try:
				 	wait_button10 = wait.until(EC.presence_of_element_located((By.XPATH,python_button10_xpath)))
					python_button10 = driver.find_elements_by_xpath(python_button10_xpath)[0]
					driver.execute_script("arguments[0].click();", python_button10)
				except NoSuchElementException:
					print("something goes wrong with xpath")
		return tag_find

	def download_directly(self,filename,code,state,year):

		print("["+str(self.d_num)+"]"+"Exception :tax code: "+code+"for:"+self.list_state[state-2]+"in: "+str(self.list_year[year-2])+"is download directly ")				
		print("file is download in: "+download_path+"; name as: "+filename)
		

	def find_disable(self,filename,code,state,year):
		tag=False
		try:
			python_button15 = driver.find_elements_by_xpath("//*[@id=\"body\"]/div[1]/div/div/div/bmb-popup[1]/div/div[2]/div[2]/div/div[2]/button")
			if len(python_button15)>0:
				python_button17=python_button15[0]
				driver.execute_script("arguments[0].click();", python_button17)
				print("["+str(self.d_num)+"]"+"Exception :tax code: "+code+"for:"+self.list_state[state-2]+"in: "+str(self.list_year[year-2])+"is unable to download")
				tag=True
			else:
				tag=False				
		except NoSuchElementException:
			print("This element is able to download.")
		if tag==True:
			try:
				python_button16 = driver.find_elements_by_xpath("//*[@id=\"body\"]/div[1]/div/div/div/bmb-popup/div/div[2]/div[2]/div/div[2]/button")
				if len(python_button16)>0:
					driver.execute_script("arguments[0].click();", python_button16[0])
			except NoSuchElementException:
				print("cancel button disappear!")
		return tag
	
	def search_code(self,st_list,yr_list,code_list):
		
		#set the state:
		for state in st_list: 
			time.sleep(9)
			driver.get(initial_url)
			#time.sleep(30)
			# without enough time there will be no click button so it will tell you index out of range
			#click the arrow to begin
			try:
				wait_button1 = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"body\"]/div[1]/div/ui-view/div/ui-view/div/ui-view/div/div/div[2]/div/div/div/div/div/div/div/div[2]/ul[1]/span[2]/span/dnd-nodrag/div/ch-content-root-item/div/div/div/ch-content-drop-down-item/div/div/span/button/span[3]")))
			finally:
				python_button1=driver.find_elements_by_xpath("//*[@id=\"body\"]/div[1]/div/ui-view/div/ui-view/div/ui-view/div/div/div[2]/div/div/div/div/div/div/div/div[2]/ul[1]/span[2]/span/dnd-nodrag/div/ch-content-root-item/div/div/div/ch-content-drop-down-item/div/div/span/button/span[3]")[0]
				driver.execute_script("arguments[0].click();", python_button1)
			time.sleep(wait_time)
			#click the state
			try:
				wait_button2 = wait.until(EC.presence_of_element_located((By.XPATH, state_xpath_pattern+str(state)+"]/a")))				
			finally:
				python_button2 = driver.find_elements_by_xpath(state_xpath_pattern+str(state)+"]/a")[0]
				driver.execute_script("arguments[0].click();", python_button2)
			time.sleep(wait_time)
			for year in yr_list:
				#update progress so you have an idea of how long you are going to wait
				driver.get(initial_url)

				print("process complete: "+str(round(self.d_num/self.total_num))+" %")
				#click arrow
				try:
					wait_button3 = wait.until(EC.presence_of_element_located((By.XPATH,"//*[@id=\"body\"]/div[1]/div/ui-view/div/ui-view/div/ui-view/div/div/div[2]/div/div/div/div/div/div/div/div[2]/ul[1]/span[2]/span/dnd-nodrag/div/ch-content-root-item/div/div/div/ch-content-drop-down-item/div/div[2]/div/div/div/ch-content-drop-down-item/div/div/span/button/span[3]")))
				finally:
					time.sleep(wait_time)
				python_button3 = driver.find_elements_by_xpath("//*[@id=\"body\"]/div[1]/div/ui-view/div/ui-view/div/ui-view/div/div/div[2]/div/div/div/div/div/div/div/div[2]/ul[1]/span[2]/span/dnd-nodrag/div/ch-content-root-item/div/div/div/ch-content-drop-down-item/div/div[2]/div/div/div/ch-content-drop-down-item/div/div/span/button/span[3]")[0]					
				driver.execute_script("arguments[0].click();", python_button3)
				
				#set year
				try:
					wait_button4 = wait.until(EC.presence_of_element_located((By.XPATH,year_xpath_pattern+str(year)+"]/a")))										
				finally:
					time.sleep(wait_time)
				python_button4 = driver.find_elements_by_xpath(year_xpath_pattern+str(year)+"]/a")[0]
				driver.execute_script("arguments[0].click();", python_button4)
				
				# click herf
				try:
					wait_button5 = wait.until(EC.presence_of_element_located((By.XPATH,"//*[@id=\"body\"]/div[1]/div/ui-view/div/ui-view/div/ui-view/div/div/div[2]/div/div/div/div/div/div/div/div[2]/ul[1]/span[2]/span/dnd-nodrag/div/ch-content-root-item/div/div/div/ch-content-drop-down-item/div/div[2]/div/div/div/ch-content-drop-down-item/div/div[2]/div/div/div/div/div/ch-content-link-item/li/a")))										
				finally:
					time.sleep(wait_time)
				python_button5 = driver.find_elements_by_xpath("//*[@id=\"body\"]/div[1]/div/ui-view/div/ui-view/div/ui-view/div/div/div[2]/div/div/div/div/div/div/div/div[2]/ul[1]/span[2]/span/dnd-nodrag/div/ch-content-root-item/div/div/div/ch-content-drop-down-item/div/div[2]/div/div/div/ch-content-drop-down-item/div/div[2]/div/div/div/div/div/ch-content-link-item/li/a")[0]
				driver.execute_script("arguments[0].click();", python_button5)
				
				#open new pages, save url 			
				
				for code in code_list: 
					Downloadpage_Url=driver.current_url
					#this is set for longer because when you download a small file directly, it is really slow.		
					#click download 1
					time.sleep(6)
					try:
						wait_button6 = wait.until(EC.presence_of_element_located((By.XPATH,"//*[@id=\"body\"]/div[1]/div/ui-view/div/ui-view/div/ui-view/div/div[3]/rs-document-view/div[2]/div/ng-include/div[2]/div/rs-document-view-items/div/div/div[2]/ng-include/div/div/span[2]/ch-common-toolbar/div/span/button[5]")))
						python_button6 = driver.find_elements_by_xpath("//*[@id=\"body\"]/div[1]/div/ui-view/div/ui-view/div/ui-view/div/div[3]/rs-document-view/div[2]/div/ng-include/div[2]/div/rs-document-view-items/div/div/div[2]/ng-include/div/div/span[2]/ch-common-toolbar/div/span/button[5]")[0]
						driver.execute_script("arguments[0].click();", python_button6)					
					finally:
						time.sleep(wait_time)
								

					try:
						wait_button6 = wait.until(EC.presence_of_element_located((By.XPATH,"//*[@id=\"body\"]/div[1]/div/ui-view/div/ui-view/div/ui-view/div/div[3]/rs-document-view/div[2]/div/ng-include/div[2]/div/rs-document-view-items/div/div/div[2]/ng-include/div/div/span[2]/ch-common-toolbar/div/span/button[5]")))
						python_button6 = driver.find_elements_by_xpath("//*[@id=\"body\"]/div[1]/div/ui-view/div/ui-view/div/ui-view/div/div[3]/rs-document-view/div[2]/div/ng-include/div[2]/div/rs-document-view-items/div/div/div[2]/ng-include/div/div/span[2]/ch-common-toolbar/div/span/button[5]")[0]
						driver.execute_script("arguments[0].click();", python_button6)					
					finally:
						time.sleep(wait_time)
						
											
					#open new pop_window set word version
					try:
						wait_button7 = wait.until(EC.presence_of_element_located((By.XPATH,"//*[@id=\"body\"]/div[1]/div/div/div/bmb-popup/div/div[2]/div[2]/div/div[1]/div/div/div[1]/fieldset[1]/span/button/span[3]")))
						python_button7=driver.find_elements_by_xpath("//*[@id=\"body\"]/div[1]/div/div/div/bmb-popup/div/div[2]/div[2]/div/div[1]/div/div/div[1]/fieldset[1]/span/button/span[3]")[0]
						driver.execute_script("arguments[0].click();", python_button7)											
					finally:
						time.sleep(wait_time)
							
					
					try:
						wait_button8 = wait.until(EC.presence_of_element_located((By.XPATH,"//*[@id=\"body\"]/div[1]/div/div/div/bmb-popup/div/div[2]/div[2]/div/div[1]/div/div/div[1]/fieldset[1]/span/div[2]/div/ul/li[2]/a")))
						python_button8 = driver.find_elements_by_xpath("//*[@id=\"body\"]/div[1]/div/div/div/bmb-popup/div/div[2]/div[2]/div/div[1]/div/div/div[1]/fieldset[1]/span/div[2]/div/ul/li[2]/a")[0]	
						driver.execute_script("arguments[0].click();", python_button8)									
					finally:
						time.sleep(wait_time)
					
											
					#uncheck the first box
					if year ==18:
						frt_path="//*[@id=\"body\"]/div[1]/div/div/div/bmb-popup/div/div[2]/div[2]/div/div[1]/div/div/div[2]/div[3]/ch-content-tree/div/div/div/div[2]/div/ul/li/div/div/div[2]/div/ul/li[2]/div/div/div[2]/div/ul/li[1]/div/div/div[2]/div/ul/li[1]/div/div/div/label/input"
					else:
						frt_path="//*[@id=\"body\"]/div[1]/div/div/div/bmb-popup/div/div[2]/div[2]/div/div[1]/div/div/div[2]/div[3]/ch-content-tree/div/div/div/div[2]/div/ul/li[1]/div/div/div[2]/div/ul/li[1]/div/div/div[2]/div/ul/li[1]/div/div/div[2]/div/ul/li[1]/div/div/div/label/input"
					try:
						wait_button9 = wait.until(EC.presence_of_element_located((By.XPATH,frt_path)))	
						python_button9 = driver.find_elements_by_xpath(frt_path)[0]
						driver.execute_script("arguments[0].click();", python_button9)									
					finally:
						time.sleep(wait_time)

						
					
											
					#check the required box: match by title
					#update 1/20 used BS4:get the soup content
					html = driver.page_source
					soup = BS(html, 'html.parser')
					code_soup=soup.find_all("span",class_="content-tree-node-text ng-binding")
					textcode= re.compile('^'+code+'$')

					tag_find=False
					#match the tag
					tag_find=self.find_tag(textcode,code_soup)
					#if a element is never find in all code_soup			
					if tag_find==False:
						#check whether there is a button called "Laws, Regs and Summaries by Tax Type"
						for tag in code_soup:
							#if this button exist, click it
							b=re.findall(r'.+Laws+.+Regs+.+and+.+Summaries+.+by+.+Tax+.+Type',tag.text)

							if len(b)>0: 
								print ("!!!!a special category is identified: "+str(b[0]))
								time.sleep(wait_time)
								python_button13 = driver.find_elements_by_xpath("//*[@id=\"body\"]/div[1]/div/div/div/bmb-popup/div/div[2]/div[2]/div/div[1]/div/div/div[2]/div[3]/ch-content-tree/div/div/div/div[2]/div/ul/li[2]/div/div/div/span/span/span/span[1]")[0]
								driver.execute_script("arguments[0].click();", python_button13)
								#renew the page content
								html2 = driver.page_source
								soup2 = BS(html2, 'html.parser')
								code_soup2=soup2.find_all("span",class_="content-tree-node-text ng-binding")
								#match the tag again
								tag_find=self.find_tag(textcode,code_soup2)
								print("*mannual correction is highly recommended")
					# if tag still could not be find
					if tag_find==False:
						self.d_num+=1
						try:
							python_button12 = driver.find_elements_by_xpath("//*[@id=\"body\"]/div[1]/div/div/div/bmb-popup/div/div[2]/div[2]/div/div[2]/button")[0]
							driver.execute_script("arguments[0].click();", python_button12)# this button is not visible						
						finally:
							print("["+str(self.d_num)+"]"+"Exception :tax code: "+code+"for:"+self.list_state[state-2]+"in: "+str(self.list_year[year-2])+"is not able to find ")
						continue
										
					#click download 2
					python_button11 = driver.find_elements_by_xpath("//*[@id=\"body\"]/div[1]/div/div/div/bmb-popup/div/div[2]/div[2]/div/div[2]/div/button/span")[0]
					driver.execute_script("arguments[0].click();", python_button11)# this button is not visible
					

					#open a new pop window type in email or download 
					filename=self.list_state[state-2]+"_"+str(self.list_year[year-2])+"_"+code[0:-2]+".rtf"
					self.d_num+=1
					time.sleep(5)
					try:					
						box= driver.find_element_by_xpath("//*[@id=\"EmailToField\"]")
						box.send_keys(email_account)						
						print("["+str(self.d_num)+"]"+"found tax code: "+code+"for state_id:"+self.list_state[state-2]+"in year_id: "+str(self.list_year[year-2]))
						
						box.submit()
						continue

					except NoSuchElementException:
						# this is sometimes because you aleady download it as the file is small
						time.sleep(3)
						self.d_num+=1
						flag_unable=False
						flag_unable=self.find_disable(filename,code,state,year)
						if flag_unable==False:
							self.download_directly(filename,code,state,year)
									
					#every round the xls will be update in case any crash									

state_list=list(range(53,55))
year_list=list(range(2,21))
#tax_code_list=['Franchise/Capital Stock Taxes\\n','Income Taxes, Corporate\\n','Income Taxes, Personal\\n','Inheritance, Estate and Gift\\n','Property\\n','Sales and Use\\n','Unemployment Compensation\\n']
#match_code=[r'FranchiseCapital+.+Stock+.+Taxes',r'Income+.+Taxes+.+Corporate',r'Income+.+Taxes+.+Personal',r'Inheritance+.+Estate+.+and+.+Gift',r'Property',r'Sales+.+and+.+Use',r'Unemployment+.+Compensation']

tax_code_list=['Franchise/Capital Stock Taxes\\n']
match_code=[r'Franchise+.+Capital+.+Stock+.+Taxes+.+']

t=autoCHH(state_list,year_list,tax_code_list)
#t.search_code(state_list,year_list,tax_code_list)
t.read_email_from_gmail()


