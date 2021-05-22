import pandas as pd
from bs4 import BeautifulSoup, NavigableString
from urllib import request
import requests
import os
assert 'SYSTEMROOT' in os.environ
import socket
import lxml
import lxml.etree
from lxml import etree
from lxml import html
import time
test=['https://www.moneycontrol.com/india/stockpricequote/pharmaceuticals/cadilahealthcare/CHC', 'https://www.moneycontrol.com/india/stockpricequote/pharmaceuticals/shilpamedicare/SM19', 'https://www.moneycontrol.com/india/stockpricequote/miscellaneous/dishmancarbogenamcis/DCA',    
'https://www.moneycontrol.com/india/stockpricequote/pharmaceuticals/aartidrugs/AD',
'https://www.moneycontrol.com/india/stockpricequote/chemicals/aartiindustries/AI45',
'https://www.moneycontrol.com/india/stockpricequote/fertilisers/chambalfertiliserschemicals/CFC',
'https://www.moneycontrol.com/india/stockpricequote/electrodesgraphite/graphiteindia/GI13',
'https://www.moneycontrol.com/india/stockpricequote/diversified/indianenergyexchange/IEE',
'https://www.moneycontrol.com/india/stockpricequote/personal-care/marico/M13',
'https://www.moneycontrol.com/india/stockpricequote/pharmaceuticals/sunpharmaadvancedresearchcompany/SPA',
'https://www.moneycontrol.com/india/stockpricequote/sugar/triveniengineeringindustries/TE10', 
'https://www.moneycontrol.com/india/stockpricequote/diversified/sbilifeinsurancecompany/SLI03', 
'https://www.moneycontrol.com/india/stockpricequote/infrastructure-general/adaniportsspecialeconomiczone/MPS',
'https://www.moneycontrol.com/india/stockpricequote/finance-stock-broking/angelbroking/ABL03',
'https://www.moneycontrol.com/india/stockpricequote/pharmaceuticals/aurobindopharma/AP',
'https://www.moneycontrol.com/india/stockpricequote/banks-private-sector/axisbank/AB16',
'https://www.moneycontrol.com/india/stockpricequote/finance-leasinghire-purchase/bajajfinance/BAF',
'https://www.moneycontrol.com/india/stockpricequote/sugar/balrampurchinimills/BCM',
'https://www.moneycontrol.com/india/stockpricequote/castingsforgings/bharatforge/BF03',
'https://www.moneycontrol.com/india/stockpricequote/telecommunications-service/bhartiairtel/BA08',
'https://www.moneycontrol.com/india/stockpricequote/refineries/bharatpetroleumcorporation/BPC',
'https://www.moneycontrol.com/india/stockpricequote/finance-leasinghire-purchase/cholamandalaminvestmentfinancecompany/CDB',
'https://www.moneycontrol.com/india/stockpricequote/pharmaceuticals/cipla/C',
'https://www.moneycontrol.com/india/stockpricequote/transportlogistics/containercrporationindia/CCI',
'https://www.moneycontrol.com/india/stockpricequote/sugar/dhampursugarmills/DSM',
'https://www.moneycontrol.com/india/stockpricequote/pharmaceuticals/glenmarkpharma/GP08',
'https://www.moneycontrol.com/india/stockpricequote/fertilisers/gujaratnarmadavalleyfertilizerschemicals/GNV',
'https://www.moneycontrol.com/india/stockpricequote/personal-care/godrejconsumerproducts/GCP',
'https://www.moneycontrol.com/india/stockpricequote/pharmaceuticals/granulesindia/GI25',
'https://www.moneycontrol.com/india/stockpricequote/pharmaceuticals/guficbiosciences/GB04',
'https://www.moneycontrol.com/india/stockpricequote/computers-software/happiestmindstechnologiesltd/HMT01',
'https://www.moneycontrol.com/india/stockpricequote/banks-private-sector/hdfcbank/HDF01',
'https://www.moneycontrol.com/india/stockpricequote/miscellaneous/hdfclifeinsurancecompanylimited/HSL01',
'https://www.moneycontrol.com/india/stockpricequote/metals-non-ferrous/hindustanzinc/HZ',
'https://www.moneycontrol.com/india/stockpricequote/finance-general/iciciprudentiallifeinsurancecompany/IPL01',
'https://www.moneycontrol.com/india/stockpricequote/banks-private-sector/indusindbank/IIB',
'https://www.moneycontrol.com/india/stockpricequote/computers-software/infosys/IT',
'https://www.moneycontrol.com/india/stockpricequote/banks-private-sector/kotakmahindrabank/KMB',
'https://www.moneycontrol.com/india/stockpricequote/infrastructure-general/larsentoubro/LT',
'https://www.moneycontrol.com/india/stockpricequote/pharmaceuticals/lupin/L',
'https://www.moneycontrol.com/india/stockpricequote/auto-carsjeeps/mahindramahindra/MM',
'https://www.moneycontrol.com/india/stockpricequote/personal-care/marico/M13',
'https://www.moneycontrol.com/india/stockpricequote/engineering-heavy/prajindustries/PI17',
'https://www.moneycontrol.com/india/stockpricequote/finance-leasinghire-purchase/shriramtransportfinancecorporation/STF',
'https://www.moneycontrol.com/india/stockpricequote/pharmaceuticals/sunpharmaadvancedresearchcompany/SPA',
'https://www.moneycontrol.com/india/stockpricequote/pharmaceuticals/sunpharmaceuticalindustries/SPI',
'https://www.moneycontrol.com/india/stockpricequote/mediaentertainment/suntvnetwork/STN01',
'https://www.moneycontrol.com/india/stockpricequote/plantations-teacoffee/tataconsumerproducts/TT',
'https://www.moneycontrol.com/india/stockpricequote/computers-software/techmahindra/TM4',
'https://www.moneycontrol.com/india/stockpricequote/metals-non-ferrous/tinplatecompanyindia/TCI02',
'https://www.moneycontrol.com/india/stockpricequote/miscellaneous/titancompany/TI01',
'https://www.moneycontrol.com/india/stockpricequote/auto-23-wheelers/tvsmotorcompany/TVS',
'https://www.moneycontrol.com/india/stockpricequote/miningminerals/vedanta/SG',
'https://www.moneycontrol.com/india/stockpricequote/diversified/voltas/V',
'https://www.moneycontrol.com/india/stockpricequote/computers-software/wipro/W',
'https://www.moneycontrol.com/india/stockpricequote/pharmaceuticals/wockhardt/W05',
'https://www.moneycontrol.com/india/stockpricequote/gas-distribution/adanitotalgas/ADG01',
'https://www.moneycontrol.com/india/stockpricequote/paintsvarnishes/asianpaints/AP31',
'https://www.moneycontrol.com/india/stockpricequote/auto-ancillaries/bosch/B05',
'https://www.moneycontrol.com/india/stockpricequote/engines/cumminsindia/CI02',
'https://www.moneycontrol.com/india/stockpricequote/diversified/generalinsurancecorporationindia/GIC12',
'https://www.moneycontrol.com/india/stockpricequote/oil-drillingexploration/gujaratstatepetronet/GSP02',
'https://www.moneycontrol.com/india/stockpricequote/electric-equipment/havellsindia/HI01',
'https://www.moneycontrol.com/india/stockpricequote/computers-software/hcltechnologies/HCL02',
'https://www.moneycontrol.com/india/stockpricequote/personal-care/hindustanunilever/HU',
'https://www.moneycontrol.com/india/stockpricequote/finance-housing/housingdevelopmentfinancecorporation/HDF',
'https://www.moneycontrol.com/india/stockpricequote/banks-private-sector/icicibank/ICI02',
'https://www.moneycontrol.com/india/stockpricequote/diversified/indianenergyexchange/IEE',
'https://www.moneycontrol.com/india/stockpricequote/finance-housing/indiabullshousingfinance/IHF01',
'https://www.moneycontrol.com/india/stockpricequote/oil-drillingexploration/indraprasthagas/IG04',
'https://www.moneycontrol.com/india/stockpricequote/telecommunications-equipment/industowers/BI14',
'https://www.moneycontrol.com/india/stockpricequote/steel-sponge-iron/jindalsteelpower/JSP',
'https://www.moneycontrol.com/india/stockpricequote/miscellaneous/justdial/JD',
'https://www.moneycontrol.com/india/stockpricequote/finance-housing/lichousingfinance/LIC',
'https://www.moneycontrol.com/india/stockpricequote/finance-investments/maxfinancialservices/MI',
'https://www.moneycontrol.com/india/stockpricequote/textiles-woollenworsted/raymond/R',
'https://www.moneycontrol.com/india/stockpricequote/refineries/relianceindustries/RI',
'https://www.moneycontrol.com/india/stockpricequote/computers-software-mediumsmall/sonatasoftware/SS42',
'https://www.moneycontrol.com/india/stockpricequote/banks-public-sector/statebankindia/SBI',
'https://www.moneycontrol.com/india/stockpricequote/auto-lcvshcvs/tatamotors/TM03',
'https://www.moneycontrol.com/india/stockpricequote/miscellaneous/newindiaassurancecompany/NIA',
'https://www.moneycontrol.com/india/stockpricequote/power-generationdistribution/torrentpower/TP14',
'https://www.moneycontrol.com/india/stockpricequote/breweriesdistilleries/unitedspirits/US',
'https://www.moneycontrol.com/india/stockpricequote/chemicals/upl/UP04',
'https://www.moneycontrol.com/india/stockpricequote/steel-pig-iron/tatametaliks/TM','https://www.moneycontrol.com/india/stockpricequote/steel-large/tatasteel/TIS','https://www.moneycontrol.com/india/stockpricequote/computers-software/tataconsultancyservices/TCS','https://www.moneycontrol.com/india/stockpricequote/plantations-teacoffee/tataconsumerproducts/TT']
timestr = time.strftime("%Y%m%d-%H%M%S")
ext="_CPR.xlsx"
timestr += ext
y=0
df=pd.DataFrame()
dff=pd.DataFrame()
for x in test:
    html_doc=requests.get(x)
    pagetext = html_doc.text
    soup = BeautifulSoup(html_doc.text, 'lxml')
    df_list = pd.read_html(html_doc.text)
    
    name=soup.h1
    df.at[y,"name"]=[name.text]
    Low = soup.find('div', attrs = {'id': 'sp_low' ,'class' : 'FL nseLP'}).text
    High = soup.find('div', attrs = {'id': 'sp_high' ,'class' : 'FR nseHP'}).text
    tbl=df_list[2]
    Open=tbl.iloc[0,1]
    Close=tbl.iloc[1,1]
    Vol=tbl.iloc[4,1]
    VWAP=df_list[3].iloc[0,1]   
    
   
    df.at[y,"Open"]=Open
    df.at[y,"High"]=High
    df.at[y,"Low"]=Low
    df.at[y,"Close"]=Close
    df.at[y,"Range"]=(float(High)-float(Low))
    df.at[y,"Vol"]=Vol
    df.at[y,"VWAP"]=VWAP

    a=soup.select('#pcforum > div > div > div.commounity_senti > div.clearfix.senti_flbxg > div > ul')
    for n in a:
        b=n.get_text()
        z=b.splitlines()
        df.at[y,"Community Sentiment"]=z[1]

    a=soup.select('#mc_insight > div.clearfix.mcibx_cnt > div:nth-child(1)')
    for n in a:
        b=n.get_text()
        k=b.splitlines()
        df.at[y, "Momentum"]=k[3]
        df.at[y, "52Low"]=k[4]
        df.at[y, "52High"]=k[5]
        df.at[y, "Performance"]=k[6]
        if k[7] is not None:
            df.at[y, "Buildup"]=k[7]
        try:
            if k[9] is not None:
                p=int("".join(filter(str.isdigit, k[9])))
                df.at[y, "Deals"]=p
        except (IndexError,ValueError):
            pass
        continue
        

    a=soup.select('#mcessential_div > div > div.escnt > div')
    for n in a:
        b=n.get_text()
        df.at[y, "MC Essentials"]=b

    a=soup.select('#insight_class')
    for n in a:
        b=n.get_text()
        z=b.splitlines()
        df.at[y, "MC Insight"]=z[1].lstrip()
        df.at[y, "Insight"]=z[2].lstrip()
    
    a=soup.select('#mc_insight > div.clearfix.mcibx_cnt > div:nth-child(2) > div.fpioi > p')
    for n in a:
        b=n.get_text()
        c=b[10:]
        df.at[y, "Financial Score"]=c
   

    df["Pivot"]=(df["High"].astype(float)+df["Low"].astype(float)+df["Close"].astype(float))/3
    df["BC"]=(df["High"].astype(float)+df["Low"].astype(float))/2
    df["TC"]=df["Pivot"]-df["BC"]+df["Pivot"]
    df.loc[(df['Close'].between((df["BC"]-1.5),(df["TC"])+1.5)) , 'Within CPR'] = 'True'
    y=y+1
    print(y)
    
    df.to_excel(timestr)
