#Execução TalentPlug por item no filtro
import smtplib
import time
from datetime import date, datetime, timedelta
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from time import sleep, strftime

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait

print(datetime.now().strftime("%d/%m/%Y %H:%M:%S"))

dtTable = {
'ExecDate': ['24/01/2022', '25/01/2022', '26/01/2022', '27/01/2022', '22/02/2022', '23/02/2022', '24/02/2022', '22/03/2022', '23/03/2022', '24/03/2022', '19/04/2022', '20/04/2022', '21/04/2022', '17/05/2022', '19/05/2022', '20/05/2022', '14/06/2022', '15/06/2022', '16/06/2022', '12/07/2022', '13/07/2022', '14/07/2022', '09/08/2022', '10/08/2022', '11/08/2022', '06/09/2022', '07/09/2022', '08/09/2022', '04/10/2022', '05/10/2022', '06/10/2022', '01/11/2022', '02/11/2022', '03/11/2022', '01/12/2022', '02/12/2022', '03/12/2022', '20/12/2022', '21/12/2022', '22/12/2022', '17/01/2023', '18/01/2023', '19/01/2023', '14/02/2023', '15/02/2023', '16/02/2023', '15/03/2023', '16/03/2023', '17/03/2023', '12/04/2023', '13/04/2023', '14/04/2023', '10/05/2023', '11/05/2023', '12/05/2023', '07/06/2023', '08/06/2023', '09/06/2023', '04/07/2023', '05/07/2023', '06/07/2023', '01/08/2023', '02/08/2023', '03/08/2023', '29/08/2023', '30/08/2023', '31/08/2023', '19/09/2023', '20/09/2023', '21/09/2023', '17/10/2023', '18/10/2023', '19/10/2023', '14/11/2023', '15/11/2023', '16/11/2023', '12/12/2023', '13/12/2023', '14/12/2023'],
'À partir du': ['29/12/2021', '29/12/2021', '29/12/2021', '29/12/2021', '26/01/2022', '27/01/2022', '28/01/2022', '23/02/2022', '24/02/2022', '25/02/2022', '23/03/2022', '24/03/2022', '25/03/2022', '20/04/2022', '21/04/2022', '22/04/2022', '18/05/2022', '20/05/2022', '21/05/2022', '15/06/2022', '16/06/2022', '17/06/2022', '13/07/2022', '14/07/2022', '15/07/2022', '10/08/2022', '11/08/2022', '12/08/2022', '07/09/2022', '08/09/2022', '09/09/2022', '05/10/2022', '06/10/2022', '07/10/2022', '03/11/2022', '04/11/2022', '05/11/2022', '02/12/2022', '03/12/2022', '04/12/2022', '21/12/2022', '22/12/2022', '23/12/2022', '18/01/2023', '19/01/2023', '20/01/2023', '15/02/2023', '16/02/2023', '17/02/2023', '16/03/2023', '17/03/2023', '18/03/2023', '13/04/2023', '14/04/2023', '15/04/2023', '11/05/2023', '12/05/2023', '13/05/2023', '08/06/2023', '09/06/2023', '10/06/2023', '05/07/2023', '06/07/2023', '07/07/2023', '02/08/2023', '03/08/2023', '04/08/2023', '30/08/2023', '31/08/2023', '01/09/2023', '20/09/2023', '21/09/2023', '22/09/2023', '18/10/2023', '19/10/2023', '20/10/2023', '15/11/2023', '16/11/2023', '17/11/2023'],
"Jusqu'au": ['20/01/2022', '20/01/2022', '27/01/2022', '20/01/2022', '26/01/2022', '27/01/2022', '22/02/2022', '23/02/2022', '24/02/2022', '22/03/2022', '23/03/2022', '24/03/2022', '19/04/2022', '20/04/2022', '21/04/2022', '17/05/2022', '19/05/2022', '20/05/2022', '21/05/2022', '15/06/2022', '16/06/2022', '12/07/2022', '13/07/2022', '14/07/2022', '09/08/2022', '10/08/2022', '11/08/2022', '06/09/2022', '07/09/2022', '08/09/2022', '04/10/2022', '05/10/2022', '06/10/2022', '02/11/2022', '03/11/2022', '04/11/2022', '01/12/2022', '02/12/2022', '03/12/2022', '20/12/2022', '21/12/2022', '22/12/2022', '17/01/2023', '18/01/2023', '19/01/2023', '14/02/2023', '15/02/2023', '16/02/2023', '15/03/2023', '16/03/2023', '17/03/2023', '12/04/2023', '13/04/2023', '14/04/2023', '10/05/2023', '11/05/2023', '12/05/2023', '07/06/2023', '08/06/2023', '09/06/2023', '04/07/2023', '05/07/2023', '06/07/2023', '01/08/2023', '02/08/2023', '03/08/2023', '29/08/2023', '30/08/2023', '31/08/2023', '19/09/2023', '20/09/2023', '21/09/2023', '17/10/2023', '18/10/2023', '19/10/2023', '14/11/2023', '15/11/2023', '16/11/2023', '12/12/2023'],
}

dT = pd.DataFrame(dtTable)
for item in dT.index:
    if (dT['ExecDate'][item]) == date.today().strftime("%d/%m/%Y"):

        try:
            navegador = webdriver.Chrome()

            WebWait = WebDriverWait(navegador,30)

            navegador.maximize_window()

            navegador.implicitly_wait(30)

            navegador.get("https://app.mytalentplug.com/Account/Login?ReturnUrl=%2Frecruiter%2Fdefault.aspx")

            Login = navegador.find_element(by=By.XPATH, value='//*[@id="UserName"]')
            Login.send_keys('kmerini@efficity.com')

            Senha = navegador.find_element(by=By.XPATH, value='//*[@id="Password"]')
            Senha.send_keys('effiCity2019')


            navegador.find_element(by=By.XPATH, value='//*[@id="c_loginControl_Login"]').click()

            time.sleep(15)
            navegador.find_element(by=By.XPATH, value='//*[@id="main"]/div[2]/div/div/div[4]/a').click()

            dtTable = {
                'ExecDate': ['24/01/2022', '25/01/2022', '26/01/2022', '27/01/2022', '22/02/2022', '23/02/2022', '24/02/2022', '22/03/2022', '23/03/2022', '24/03/2022', '19/04/2022', '20/04/2022', '21/04/2022', '17/05/2022', '19/05/2022', '20/05/2022', '14/06/2022', '15/06/2022', '16/06/2022', '12/07/2022', '13/07/2022', '14/07/2022', '09/08/2022', '10/08/2022', '11/08/2022', '06/09/2022', '07/09/2022', '08/09/2022', '04/10/2022', '05/10/2022', '06/10/2022', '01/11/2022', '02/11/2022', '03/11/2022', '01/12/2022', '02/12/2022', '03/12/2022', '20/12/2022', '21/12/2022', '22/12/2022', '17/01/2023', '18/01/2023', '19/01/2023', '14/02/2023', '15/02/2023', '16/02/2023', '15/03/2023', '16/03/2023', '17/03/2023', '12/04/2023', '13/04/2023', '14/04/2023', '10/05/2023', '11/05/2023', '12/05/2023', '07/06/2023', '08/06/2023', '09/06/2023', '04/07/2023', '05/07/2023', '06/07/2023', '01/08/2023', '02/08/2023', '03/08/2023', '29/08/2023', '30/08/2023', '31/08/2023', '19/09/2023', '20/09/2023', '21/09/2023', '17/10/2023', '18/10/2023', '19/10/2023', '14/11/2023', '15/11/2023', '16/11/2023', '12/12/2023', '13/12/2023', '14/12/2023'],
                'À partir du': ['29/12/2021', '29/12/2021', '29/12/2021', '29/12/2021', '26/01/2022', '27/01/2022', '28/01/2022', '23/02/2022', '24/02/2022', '25/02/2022', '23/03/2022', '24/03/2022', '25/03/2022', '20/04/2022', '21/04/2022', '22/04/2022', '18/05/2022', '20/05/2022', '21/05/2022', '15/06/2022', '16/06/2022', '17/06/2022', '13/07/2022', '14/07/2022', '15/07/2022', '10/08/2022', '11/08/2022', '12/08/2022', '07/09/2022', '08/09/2022', '09/09/2022', '05/10/2022', '06/10/2022', '07/10/2022', '03/11/2022', '04/11/2022', '05/11/2022', '02/12/2022', '03/12/2022', '04/12/2022', '21/12/2022', '22/12/2022', '23/12/2022', '18/01/2023', '19/01/2023', '20/01/2023', '15/02/2023', '16/02/2023', '17/02/2023', '16/03/2023', '17/03/2023', '18/03/2023', '13/04/2023', '14/04/2023', '15/04/2023', '11/05/2023', '12/05/2023', '13/05/2023', '08/06/2023', '09/06/2023', '10/06/2023', '05/07/2023', '06/07/2023', '07/07/2023', '02/08/2023', '03/08/2023', '04/08/2023', '30/08/2023', '31/08/2023', '01/09/2023', '20/09/2023', '21/09/2023', '22/09/2023', '18/10/2023', '19/10/2023', '20/10/2023', '15/11/2023', '16/11/2023', '17/11/2023'],
                "Jusqu'au": ['20/01/2022', '20/01/2022', '27/01/2022', '20/01/2022', '26/01/2022', '27/01/2022', '22/02/2022', '23/02/2022', '24/02/2022', '22/03/2022', '23/03/2022', '24/03/2022', '19/04/2022', '20/04/2022', '21/04/2022', '17/05/2022', '19/05/2022', '20/05/2022', '21/05/2022', '15/06/2022', '16/06/2022', '12/07/2022', '13/07/2022', '14/07/2022', '09/08/2022', '10/08/2022', '11/08/2022', '06/09/2022', '07/09/2022', '08/09/2022', '04/10/2022', '05/10/2022', '06/10/2022', '02/11/2022', '03/11/2022', '04/11/2022', '01/12/2022', '02/12/2022', '03/12/2022', '20/12/2022', '21/12/2022', '22/12/2022', '17/01/2023', '18/01/2023', '19/01/2023', '14/02/2023', '15/02/2023', '16/02/2023', '15/03/2023', '16/03/2023', '17/03/2023', '12/04/2023', '13/04/2023', '14/04/2023', '10/05/2023', '11/05/2023', '12/05/2023', '07/06/2023', '08/06/2023', '09/06/2023', '04/07/2023', '05/07/2023', '06/07/2023', '01/08/2023', '02/08/2023', '03/08/2023', '29/08/2023', '30/08/2023', '31/08/2023', '19/09/2023', '20/09/2023', '21/09/2023', '17/10/2023', '18/10/2023', '19/10/2023', '14/11/2023', '15/11/2023', '16/11/2023', '12/12/2023'],
            }

            dT = pd.DataFrame(dtTable)

            for item in dT.index:
                if (dT['ExecDate'][item]) == date.today().strftime("%d/%m/%Y"):
                    WebWait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="grid"]/div[1]/div/table/thead/tr/th[3]/a[1]')))
                    navegador.find_element(by=By.XPATH, value='//*[@id="grid"]/div[1]/div/table/thead/tr/th[3]/a[1]').click()
                    time.sleep(1)
                    navegador.find_element(by=By.XPATH, value='/html/body/div[3]/form/div[1]/span[2]/span/input').send_keys(dT['À partir du'][item])
                    
                    navegador.find_element(by=By.XPATH, value='/html/body/div[3]/form/div[1]/span[5]/span/input').send_keys(dT["Jusqu'au"][item])
                    
                    navegador.find_element(by=By.XPATH, value='/html/body/div[3]/form/div[1]/div[2]/button[1]').click()

                    time.sleep(5)
                    page=1
                    References=[]
                    Line_id=[]
                    df=pd.DataFrame({"References":References, "Line_id":Line_id})

                    next_page = navegador.find_element(by=By.XPATH, value='//*[@id="grid"]/div[3]/a[3]').get_attribute('data-page')
                    next_pagestr = str(next_page)

                    actual_page = navegador.find_element(by=By.XPATH, value='//*[@id="grid"]/div[3]/ul/li/span').text
                    actual_pagestr = str(actual_page)

                    while next_pagestr!=actual_pagestr:
                    #while page<=9:
                        next_page = navegador.find_element(by=By.XPATH, value='//*[@id="grid"]/div[3]/a[3]').get_attribute('data-page')
                        next_pagestr = str(next_page)
                        rows = WebWait.until(EC.visibility_of_all_elements_located((By.XPATH, '//*[@id="grid"]/div[2]/table/tbody/tr')))
                        for row in rows:
                            Line_id.append(row.find_element(by=By.XPATH, value='./td[4]').get_attribute('innerText'))
                            References.append(row.find_element(by=By.XPATH, value='./td[5]').text)            
                        WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="grid"]/div[3]/a[3]/span'))).click()
                        page=page+1
                        actual_page = navegador.find_element(by=By.XPATH, value='//*[@id="grid"]/div[3]/ul/li/span').text
                        actual_pagestr = str(actual_page)
                        time.sleep(3)
                    df=pd.DataFrame({'References':References, 'Line_id':Line_id})
                    #print(df)

                    for line in df.index:
                        #print(df['References'][line])
                        line_str = str(df['References'][line])[-3:]
                        if line_str != "LBC" and line_str != "BC)":
                            #print(line_str)
                            monster_str = str(df['References'][line])[-8:]
                            #print(monster_str)
                            if monster_str != "Monster)":
                                time.sleep(2)
                                WebWait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="grid"]/div[1]/div/table/thead/tr/th[5]/a[1]/span')))
                                WebWait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="grid"]/div[1]/div/table/thead/tr/th[5]/a[1]/span')))               
                                time.sleep(2)
                                navegador.find_element(by=By.XPATH, value='//*[@id="grid"]/div[1]/div/table/thead/tr/th[5]/a[1]/span').click()
                                navegador.find_element(by=By.XPATH, value='/html/body/div[4]/form/div[1]/input[1]').clear()
                                navegador.find_element(by=By.XPATH, value='/html/body/div[4]/form/div[1]/input[1]').send_keys(df['References'][line])
                                navegador.find_element(by=By.XPATH, value='/html/body/div[4]/form/div[1]/div[2]/button[1]').click()
                                time.sleep(1)
                                WebWait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="deletejb-'+df['Line_id'][line]+'"]/i')))
                                WebWait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="deletejb-'+df['Line_id'][line]+'"]/i')))
                                time.sleep(2)
                                navegador.find_element(by=By.XPATH, value='//*[@id="deletejb-'+df['Line_id'][line]+'"]/i').click()            
                                time.sleep(1)
                                WebWait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="validerRepublish"]/span')))
                                #navegador.find_element(by=By.XPATH, value='//*[@id="annulerRepublish"]/span').click()
                                navegador.find_element(by=By.XPATH, value='//*[@id="validerRepublish"]/span').click()
                                WebWait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="republishSuccessPopup"]/div[2]/a')))
                                navegador.find_element(by=By.XPATH, value='//*[@id="republishSuccessPopup"]/div[2]/a').click()
                                #print(line)
                            else:
                                time.sleep(2)
                                WebWait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="grid"]/div[1]/div/table/thead/tr/th[5]/a[1]/span')))
                                WebWait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="grid"]/div[1]/div/table/thead/tr/th[5]/a[1]/span')))               
                                time.sleep(2)
                                navegador.find_element(by=By.XPATH, value='//*[@id="grid"]/div[1]/div/table/thead/tr/th[5]/a[1]/span').click()
                                navegador.find_element(by=By.XPATH, value='/html/body/div[4]/form/div[1]/input[1]').clear()
                                navegador.find_element(by=By.XPATH, value='/html/body/div[4]/form/div[1]/input[1]').send_keys(df['References'][line][:-9])
                                print(df['References'][line][:-9])
                                navegador.find_element(by=By.XPATH, value='/html/body/div[4]/form/div[1]/div[2]/button[1]').click()
                                time.sleep(1)
                                WebWait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="deletejb-'+df['Line_id'][line]+'"]/i')))
                                WebWait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="deletejb-'+df['Line_id'][line]+'"]/i')))
                                time.sleep(2)
                                navegador.find_element(by=By.XPATH, value='//*[@id="deletejb-'+df['Line_id'][line]+'"]/i').click()            
                                time.sleep(1)
                                WebWait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="validerRepublish"]/span')))
                                #navegador.find_element(by=By.XPATH, value='//*[@id="annulerRepublish"]/span').click()
                                navegador.find_element(by=By.XPATH, value='//*[@id="validerRepublish"]/span').click()
                                WebWait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="republishSuccessPopup"]/div[2]/a')))
                                navegador.find_element(by=By.XPATH, value='//*[@id="republishSuccessPopup"]/div[2]/a').click()
                                #print(line)
                    
                    WebWait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="grid"]/div[1]/div/table/thead/tr/th[5]/a[1]/span')))
                    WebWait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="grid"]/div[1]/div/table/thead/tr/th[5]/a[1]/span')))               
                    time.sleep(2)
                    navegador.find_element(by=By.XPATH, value='//*[@id="grid"]/div[1]/div/table/thead/tr/th[5]/a[1]/span').click()
                    navegador.find_element(by=By.XPATH, value='/html/body/div[4]/form/div[1]/input[1]').clear()
                    WebWait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="grid"]/div[1]/div/table/thead/tr/th[3]/a[1]')))
                    navegador.find_element(by=By.XPATH, value='//*[@id="grid"]/div[1]/div/table/thead/tr/th[3]/a[1]').click()
                    time.sleep(1)  
                
                    navegador.find_element(by=By.XPATH, value='/html/body/div[3]/form/div[1]/span[2]/span/input').clear()
                    navegador.find_element(by=By.XPATH, value='/html/body/div[3]/form/div[1]/span[2]/span/input').send_keys((date.today()+timedelta(days=1)).strftime("%d/%m/%Y"))
                    navegador.find_element(by=By.XPATH, value='/html/body/div[3]/form/div[1]/span[2]/span/input').clear()
                    navegador.find_element(by=By.XPATH, value='/html/body/div[3]/form/div[1]/span[5]/span/input').send_keys((date.today()+timedelta(days=1)).strftime("%d/%m/%Y"))
                    
                    navegador.find_element(by=By.XPATH, value='/html/body/div[3]/form/div[1]/div[2]/button[1]').click()

                    time.sleep(5)
                    
                    TotalElements = navegador.find_element(by=By.XPATH, value='//*[@id="grid"]/div[3]/span').text
                    #print(TotalElements)
                    TotalElementsStr = str(TotalElements)
                    #print(TotalElementsStr)
                    TotalElementsStr = TotalElementsStr[11:14].strip()

            #The mail addresses and password
            sender_address = 'boombots002@gmail.com'
            #sender_pass = 'Passe123$'
            sender_pass = 'maitxsneemsyiane'
            receiver_address = ['chris@boombots.co', 'cintia.tominaga@boombots.co', 'recrutement@efficity.com', 'kmerini@efficity.com']
            notifier_adress = ['chris@boombots.co','christophertominaga@hotmail.com']

            #General Attach Image 
            fp_logo = open("img/logo_sem_icone.png", 'rb') #Read image 
            msgImage_logo = MIMEImage(fp_logo.read())
            fp_logo.close()

            #Setup the MIME

            html = open("EmailBodyHtmlSuccess.html", encoding="utf-8")
            message = MIMEMultipart()
            message['From'] = sender_address
            message['To'] = ", ".join(receiver_address)
            message['Subject'] = "[Notification] Exécution effectuée avec succès"   #The subject line
            #The body and the attachments for the mail#
            #message.attach(MIMEText(html.read().replace("TotalElements",TotalElements), 'html'))

            message.attach(MIMEText(html.read().replace("TotalElements",TotalElementsStr), 'html'))

            #Attach Image 
            fp_success = open("img/hifivemarinho-2.png", 'rb') #Read image 
            msgImage_success = MIMEImage(fp_success.read())
            fp_success.close()

            # Define the image's ID as referenced above
            msgImage_logo.add_header('Content-ID', '<logo_sem_icone.png>')
            message.attach(msgImage_logo)
            msgImage_success.add_header('Content-ID', '<hifivemarinho-2.png>')
            message.attach(msgImage_success)


                #Create SMTP session for sending the mail
            session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
            session.starttls() #enable security
            session.login(sender_address, sender_pass) #login with mail_id and password
            text = message.as_string()
            session.sendmail(sender_address, receiver_address, text)
            session.quit() 

        except:
            #The mail addresses and password
            sender_address = 'boombots002@gmail.com'
            #sender_pass = 'Passe123$'
            sender_pass = 'maitxsneemsyiane'
            receiver_address = ['chris@boombots.co', 'cintia.tominaga@boombots.co', 'recrutement@efficity.com', 'kmerini@efficity.com']
            notifier_adress = ['chris@boombots.co','christophertominaga@hotmail.com']

            #General Attach Image 
            fp_logo = open("img/logo_sem_icone.png", 'rb') #Read image 
            msgImage_logo = MIMEImage(fp_logo.read())
            fp_logo.close()

            #Setup the MIME

            html_impossible = open("EmailBodyHtmlImpossible.html", encoding="utf-8")
            message_impossible = MIMEMultipart()
            message_impossible['From'] = sender_address
            message_impossible['To'] = ", ".join(notifier_adress)
            message_impossible['Subject'] = "[Notification] Impossible de démarrer l'exécution"   #The subject line
            #The body and the attachments for the mail#
            message_impossible.attach(MIMEText(html_impossible.read(), 'html'))

            #Attach Image 
            fp_impossible = open("img/hifivetriste.png", 'rb') #Read image 
            msgImage_impossible = MIMEImage(fp_impossible.read())
            fp_impossible.close()

            # Define the image's ID as referenced above
            msgImage_logo.add_header('Content-ID', '<logo_sem_icone.png>')
            message_impossible.attach(msgImage_logo)
            msgImage_impossible.add_header('Content-ID', '<hifivetriste.png>')
            message_impossible.attach(msgImage_impossible)


                #Create SMTP session for sending the mail
            session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
            session.starttls() #enable security
            session.login(sender_address, sender_pass) #login with mail_id and password
            text_impossible = message_impossible.as_string()
            session.sendmail(sender_address, notifier_adress, text_impossible)
            session.quit() 

print(datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
