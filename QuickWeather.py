import json, requests
import glob, os, win32com.client, pythoncom, datetime, time, getpass
import win32com.client as win32
import comtypes, comtypes.client
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
import selenium.webdriver.support.ui as ui
from selenium.webdriver.support.wait import WebDriverWait
import pandas as pd                    
import numpy as np                    
from pandas import DataFrame, Series  
import matplotlib.pyplot as plt      

# Setting browser preferences... (Not used in the office because they don't have Firefox)
#mfp = webdriver.FirefoxProfile()
#mfp.set_preference("browser.download.folderList", 2)
#mfp.set_preference("browser.download.manager.showWhenStarting", False)
#mfp.set_preference("browser.download.dir", os.getcwd())
#mfp.set_preference("browser.helperApps.neverAsk.saveToDisk","application/vnd.ms-excel")
#br = webdriver.Firefox(firefox_profile = mfp)

# Starting Google Chrome
#chromeDriver = "C:\\Users\\davidedwards\\Downloads\\chromedriver_win32\\chromedriver"
chromeDriver = "C:\\Users\\"+getpass.getuser()+"\\Downloads\\chromedriver_win32\\chromedriver"
br = webdriver.Chrome(chromeDriver)

# Creating a folder for these files using today's date...
t = datetime.datetime.today()                           
today = t.strftime('%Y-%m-%d')

                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      #cFolder = 'C:\\Users\\EDWARDS\\Desktop'
#cFolder = 'C:\\User\\'+getpass.getuser()+'\\Desktop'
#cFolder = 'C:\\User\\'+getpass.getuser()+'\\Downloads\\'
cFolder = 'https://atgmedia-my.sharepoint.com/personal//'+getpass.getuser()+'_auctiontechnologygroup_com//Documents//Laptop//Desktop//'
wkbk = os.path.expanduser(cFolder)

# The first part of the API url (the full string is in 3 parts)...
url = 'http://api.openweathermap.org/data/2.5/weather?q='

# Criteria for countries and cities...
location = {'England':'London,uk','Ghana':'Accra,gh','France':'Paris,fr','Afghanistan':'Kabul,af',
            'Albania':'Tirana,al','Algeria':'Algiers,dz','Andorra':'Andorra la Vella,ad','Angola':'Luanda,ao',
            'Antigua and Barbuda':'Saint Johns,ag','Argentina':'Bueonos Aires,ar','Armenia':'Yerevan,am','Australia':'Canberra,au',
            'Austria':'Vienna,at','Azerbaijan':'Baku,az','Bahamas':'Nassau,bs','Bahrain':'Manama,bh','Bangladesh':'Dhaka,bd',
            'Barbados':'Bridgetown,bb','Belarus':'Minsk,by','Belgium':'Brussels,be','Belize':'Belmopan,bz',
            'Benin':'Porto-Novo,bj','Bhutan':'Thimphu,bt','Bolivia':'La Paz,bo','Bosnia and Herzegovina':'Sarajevo,ba',
            'Botswana':'Gaborone,bw','Brazil':'Brasilia,br','Brunei':'Bandar Seri Begawan,bn','Bulgaria':'Sofia,bg',
            'Burkina Faso':'Ouagadougou,bf','Burundi':'Bujumbura,bi','Cape Verde':'Praia,cv','Cambodia':'Phnom Penh,kh',
            'Cameroon':'Yaounde,cm','Canada':'Ottawa,ca','Central African Republic':'Bangui,cf','Chile':'Santiago,cl',
            'China':'Beijing,cn','Colombia':'Bogota,co','Comoros':'Moroni,km','Congo':'Kinshasa,cg',
            'Costa Rica':'San Jose,cr','Cote d\'Ivoire':'Yamoussoukro,ci','Croatia':'Zagreb,hr','Cuba':'Havana,cu',
            'Cyprus':'Nicosia,cy','The Czech Republic':'Prague,cz','Denmark':'Copenhagen,dk','Djibouti':'Djibouti,dj',
            'Dominica':'Roseau,dm','Dominican Republic':'Santo Domingo,do','Ecuador':'Quito,ec','Egypt':'Cairo,eg',
            'El Salvador':'San Salvador,sv','Equatorial Guinea':'Malabo,gq','Eritrea':'Asmara,er','Estonia':'Tallinn,ee',
            'Ethiopia':'Addis Ababa,et','Fiji':'Suva,fj','Finland':'Helsinki,fi','Gabon':'Libreville,ga',
            'Gambia':'Banjul,gm','Georgia':'Tbilisi,ge','Germany':'Berlin,de','Greece':'Athens,gr',
            'Guatemala':'Guatemala City,gt','Guinea':'Conakry,gn','Guinea-Bissau':'Bissau,gw','Guyana':'Georgetown,gy',
            'Haiti':'Port-au-Prince,ht','Honduras':'Tegucigalpa,hn','Iceland':'Reykjavik,is','India':'New Delhi,in',
            'Indonesia':'Jakarta,id','Iran':'Tehran,ir','Iraq':'Baghdad,iq','Ireland':'Dublin,ie','Israel':'Jerusalem,il',
            'Italy':'Rome,it','Jamaica':'Kingston,jm','Japan':'Tokyo,jp','Jordan':'Amman,jo','Kazakhstan':'Astana,kz',
            'Kenya':'Nairobi,ke','Kiribati':'South Tarawa,ki','Kyrgyzstan':'Bishkek,kg','Lao PDR':'Vientiane,la',
            'Latvia':'Riga,lv','Lebanon':'Beirut,lb','Lesotho':'Maseru,ls','Liberia':'Monrovia,lr','Libya':'Tripoli,ly',
            'Liechtenstein':'Vaduz,li','Lithuania':'Vilnius,lt','Luxembourg':'Luxembourg City,lu','Republic of Macedonia':'Skopje,mk',
            'Madagascar':'Antananarivo,mg','Malawi':'Lilongwe,mw','Malaysia':'Kuala Lumpur,my','Maldives':'Male,mv','Mali':'Bamako,ml',
            'Malta':'Valletta,mt','Marshall Islands':'Majuro,mh','Mauritania':'Nouakchott,mr','Mauritius':'Port Louis,mu','Mexico':'Mexico City,mx',
            'Federated States of Micronesia':'Palikir,fm','Moldova':'Chisinau,md','Monaco':'Monte Carlo,mc','Mongolia':'Ulaanbaatar,mn','Montenegro':'Podgorica,me',
            'Morocco':'Rabat,ma','Mozambique':'Maputo,mz','Myanmar':'Naypyidaw,mm','Namibia':'Windhoek,na','Nauru':'Yaren District,nr',
            'Nepal':'Kathmandu,np','Netherlands':'Amsterdam,nl','New Zealand':'Wellington,nz','Nicaragua':'Managua,ni','Niger':'Niamey,ne',
            'Nigeria':'Abuja,ng','North Korea':'Pyongyang,kp','Norway':'Oslo,no','Oman':'Muscat,om','Pakistan':'Islamabad,pk',
            'Palau':'Ngerulmud,pw','Palestinian Territory (Occupied)':'Ramallah,ps','Panama':'Panama City,pa','Papua New Guinea':'Port Moresby,pg','Paraguay':'Asuncion,py',
            'Peru':'Lima,pe','Philippines':'Manila,ph','Poland':'Warsaw,pl','Portugal':'Lisbon,pt','Qatar':'Doha,qa',
            'Romania':'Bucharest,ro','Russia':'Moscow,ru','Rwanda':'Kigali,rw','Saint Kitts and Nevis':'Basseterre,kn','Saint Lucia':'Castries,lc',
            'Saint Vincent and Grenadines':'Kingstown,vc','Samoa':'Apia,ws','San Marino':'San Marino,sm','Sao Tome and Principe':'Sao Tome,st','Saudi Arabia':'Riyadh,sa',
            'Senegal':'Dakar,sn','Serbia':'Belgrade,rs','Seychelles':'Victoria,sc','Sierra Leone':'Freetown,sl','Singapore':'Singapore,sg',
            'Slovakia':'Bratislava,sk','Sierra Leone':'Ljubljana,sl','Solomon Islands':'Honiara,sb','Somalia':'Mogadishu,so',
            'South Africa':'Pretoria,za','South Africa':'Cape Town,za','South Africa':'Bloemfontein,za','Korea':'Seoul,kr','South Sudan':'Juba,ss',
            'Spain':'Madrid,es','Sri Lanka':'Sri Jayawardenepura Kotte,lk','Sudan':'Khartoum,sd','Suriname':'Paramaribo,sr','Swaziland':'Mbabane,sz',
            'Sweden':'Stockholm,se','Switzerland':'Bern,ch','Syria':'Damascus,sy','Taiwan':'Taipei,tw','Tajikistan':'Dushanbe,tj','United Republic of Tanzania':'Dodoma,tz',
            'Thailand':'Bangkok,th','Timor-Leste':'Dili,tl','Togo':'Lome,tg','Tongo':'Nuk\'alofa,to','Trinidad and Tobago':'Port of Spain,tt',
            'Tunisia':'Tunis,tn','Turkey':'Ankara,tr','Turkmenistan':'Ashgabat,tm','Tuvalu':'Funafuti,tv','Uganda':'Kampala,ug','Ukraine':'Kyiv,ua',
            'United Arab Emirates':'Abu Dhabi,ae','United States of America':'Washington D.C.,us','Uruguay':'Montevideo,uy','Uzbekistan':'Tashkent,uz',
            'Vanuatu':'Port Vila,vu','Venezuela':'Caracas,ve','Viet Nam':'Hanoi,vn','Yemen':'Sana\'a,ye','Zimbabwe':'Lusaka,zw',
            'Zimbabwe':'Harare,zw'}

# The second part of the API url string...
appid = '&appid='

# Getting the API key...
try:
    br.get('https://home.openweathermap.org/api_keys')
    email = br.find_element_by_xpath(".//input[@id='user_email' and @name='user[email]']")
    email.send_keys('dcyedwards@yahoo.com')
    password = br.find_element_by_xpath(".//input[@type='password' and @name='user[password]']")
    password.send_keys('jazzMAN2')
    remember_me = br.find_element_by_xpath(".//input[@id='user_remember_me']").click()
    submit = br.find_element_by_xpath(".//input[@name='commit' and @value='Submit']").click()
    key = br.find_element_by_tag_name('pre').text # Yep, that's the key
except:
    print('It\'s a dud')

# Okay, now to store that data in Excel...
#xl = win32com.client.Dispatch("Excel.Application")
# Creating an instance of Excel in memory
xl = win32.gencache.EnsureDispatch('Excel.Application')
xl.DisplayAlerts = False
wb = xl.Workbooks.Add()
xlOpenXMLWorkbookMacroEnabled = 52
xlmodule = wb.VBProject.VBComponents.Add(1)

# ...VBA code...
VBA = '''Sub TidyConverter()
Dim wb As Workbook: Set wb = ThisWorkbook
Dim ws As Worksheet: Set ws = wb.Sheets(1)
Dim i As Long, lasti As Long

i = 1
lasti = ws.Cells(Rows.Count, "A").End(xlUp).Row

'Convert from Kelvin to Celcius
For i = 1 To lasti
    If ws.Range("A" & i).Value = "temp" Or _
    ws.Range("A" & i).Value = "temp_min" Or _
    ws.Range("A" & i).Value = "temp_max" Then
        ws.Range("B" & i).Value = ws.Range("B" & i).Value - 273.15
        ws.Range("C" & i).Value = "°C"
    End If
Next i

'Appending hPa pressure sign
i = 1
For i = 1 To lasti
    If ws.Range("A" & i).Value = "pressure" Then
        ws.Range("C" & i).Value = "hPa"
    End If
Next i

'Appending humidity % sign
i = 1
For i = 1 To lasti
    If ws.Range("A" & i).Value = "humidity" Then
        ws.Range("C" & i).Value = "%"
    End If
Next i
End Sub'''

# Saving my created instance of a Macro-enabled workbook for eventual population with data
wb.SaveAs(wkbk +" Weather Book - "+today+".xlsm", FileFormat=xlOpenXMLWorkbookMacroEnabled)
xl.Visible = True
wb.Worksheets.Add()
sht = wb.Worksheets('Sheet2')
sht2 = wb.Worksheets('Sheet1')
used = sht.UsedRange
nrow = used.Row + used.Rows.Count-1
nrow2 = nrow+1
nrow3 = nrow2+1

# Closing the Firefox browser...
br.close()
sht2.Range('A1').Value = 'List of Countries, their capitals and weather forecasts'
sht2.Range('A3').Value = 'Legend'
sht2.Range('A4').Value = 'Column B = Temperature in Celcius °C, Pressure in Hectopascal hPa, humidity in %'
sht2.Range('A5').Value = 'Column C = Unit of measurement symbol'

# Now for some iterative logic:
for k,v in location.items():                  # For each country(k) and capital(v)...
    response = requests.get(url+v+appid+key)  # The response we want from our API (i.e.: from the weather website) is built up of a URL and API key
    response.raise_for_status()               # Just a check to make sure there's a response from the website (A response code of 200 is good. Anything else, not so good)
    response                                  # All the data for each country and capital is stored in the 'response' variable.
    weatherData = json.loads(response.text)   # Now, as this data is in the json data format, we load it into Python using the loads(loads = load string) into the 'weatherData' variable.
    w = weatherData                           # Using a shorter variable name for ease of writing really.
    print('\n')
    print(w['name'])                          # Now, the data is in a Python dictionary. I select the elements in the dictionary of interest beginning with the 'name' key.
    sht.Range('A'+str(nrow)).Value = w['name']# Writing the data from the dictionary to our Excel workbook
    sht.Range('C'+str(nrow)).Value = 'Country:'
    sht.Range('D'+str(nrow)).Value = k
    print(w['weather'][0]['main'])
    sht.Range('A'+str(nrow2)).Value = w['weather'][0]['main']
    for k,v in w['main'].items():
        print(str(k),str(v))
        sht.Range('A'+str(nrow3)+':A'+str(int(nrow3)+4)).Value = [[k] for k in w['main'].keys()]
        sht.Range('B'+str(nrow3)+':B'+str(int(nrow3)+4)).Value = [[v] for v in w['main'].values()]
        d = list(w['main'].values())          # Creating a list of the temperatures, pressures and humidity values and assigning to a variable 'd'| I don't really need these but I do it to check values in IDLE just in case.
        dlist = d[1:5]                        # Selecting the 2nd to 5th values from that list                                                    |  
        #sht.Range('C'+str(int(nrow3)+1)+':C'+str(int(nrow3)+3)).Value = [[temp-273.15] for temp in dlist] # Converting the temperatures from Kelvin to Degrees Celcius[Deprecated] - I do this using VBA instead.
    used = sht.UsedRange
    nrow = used.Row + used.Rows.Count-1
    nrow = nrow + 2
    nrow2 = nrow+1
    nrow3 = nrow2+1

xlmodule.CodeModule.AddFromString(VBA)     # Now to insert our VBA code into a module
xl.Application.Run("TidyConverter")          # Running that VBA code to clean up the data

wb.Save()                                    # Saving...
xl.DisplayAlerts = True
xl.Quit()                                    # And done
number_of_countries = [v for v in location.values()]
print('\nWeather data downloaded for '+str(len(number_of_countries))+' countries')

