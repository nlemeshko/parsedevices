import os

try:
    from bs4 import BeautifulSoup as BS
except ImportError:
  print ("Trying to Install required module: BeautifulSoup\n")
  os.system('python -m pip install bs4')

try:
    import pandas as pd
except ImportError:
  print ("Trying to Install required module: pandas\n")
  os.system('python -m pip install pandas')

try:
    import lxml.html as LH
except ImportError:
  print ("Trying to Install required module: lxml.html\n")
  os.system('python -m pip install lxml')

try:
    import requests
except ImportError:
  print ("Trying to Install required module: requests\n")
  os.system('python -m pip install requests')

try:
    import re
except ImportError:
  print ("Trying to Install required module: re\n")
  os.system('python -m pip install re')

try:
    import xlwt
except ImportError:
  print ("Trying to Install required module: xlwt\n")
  os.system('python -m pip install xlwt')

try:
    import xlrd
except ImportError:
  print ("Trying to Install required module: xlrd\n")
  os.system('python -m pip install xlrd')

try:
    from tempfile import TemporaryFile
except ImportError:
  print ("Trying to Install required module: TemporaryFile\n")
  os.system('python -m pip install temporaryfile')

max_page = 6
pages = []
titlenew = []
dopnew = []
url='https://getitnew.com'



for x in range(1, max_page + 1):
    #pages.append( requests.get('https://getitnew.com/collections/cisco-catalyst-express-switches'))
    pages.append( requests.get('https://getitnew.com/collections/juniper-networks-switches?page=' + str(x)))

for r in pages:
    html = BS(r.content, 'html.parser')

    for el in html.select('.productitem--info'):
        title = el.select('.productitem--title > a')
        if 'Switch' in str(title):
            titlenew.append(title[0].get('href'))
        else:
            dopnew.append(title[0].get('href'))

book = xlwt.Workbook()

print('Starting parse switches.')
dfnames = pd.read_excel ('template.xlsx')
sheet1 = book.add_sheet('Mirawork')
for i,e in enumerate(dfnames):
    sheet1.write(0,i,e)
wrk=0
for wrk in range(len(titlenew)):
    #url = 'https://getitnew.com/collections/cisco-catalyst-9500-switches/products/c9500-32qc-a'
    #url = 'https://getitnew.com/collections/cisco-catalyst-3650-switches/products/c1-ws3650-24pd-k9'
    try:
        df=pd.read_html(url+titlenew[wrk])
    except Exception:
        pass
    pd.options.display.max_colwidth = 10000
    r = requests.get(url+titlenew[wrk])
    root = LH.fromstring(r.content)



    column = list()

    #Тип оборудования
    type = ['bundle']
    column.append(type[0])

    #Производитель
    vendor = ['Cisco']
    column.append(vendor[0])

    #Серия
    serial = root.xpath('//*[@id="shopify-section-static-product"]/section/article/div[2]/div[1]/div[3]/span/text()')
    try:
        column.append(serial[0][:serial[0].find("-")])
    except Exception:
        pass


    #Подсерия
    try:
        column.append(serial[0][:serial[0].find("-")])
    except Exception:
        pass


    #Партномер
    name = root.xpath('//*[@id="shopify-section-static-product"]/section/article/div[2]/div[1]/div[3]/span/text()')
    column.append(name[0].rstrip())

    #Название устройства
    column.append(name[0].rstrip())

    #Описание
    desc = root.xpath('//*[@id="shopify-section-static-product"]/section/article/div[2]/div[3]/p[3]/text()')
    column.append(desc[0])

    #Описание расширенное
    #bref = root.xpath('//*[@id="shopify-section-static-product"]/section/article/div[2]/div[3]/p[4]/span/text()')
    #column.append(bref[0])
    column.append('')

    #Допы
    column.append('')

    #Price
    column.append('')

    #Ширина
    try:
        height = float(df[0].loc[df[0][0] == 'Height'][1].to_string(index=False)[1:-3])/0.039370
        column.append(height)
    except Exception:
        column.append('')
        pass

    #Глубина
    try:
        depth = float(df[0].loc[df[0][0] == 'Depth'][1].to_string(index=False)[1:-3])/0.039370
        column.append(height)
    except Exception:
        column.append('')
        pass

    #Высота
    try:
        width = float(df[0].loc[df[0][0] == 'Width'][1].to_string(index=False)[1:-3])/0.039370
        column.append(width)
    except Exception:
        column.append('')
        pass

    #Package Height
    column.append('')

    #Package Depth
    column.append('')

    #Package Width
    column.append('')

    #Вес
    try:
        weight = df[0].loc[df[0][0] == 'Weight'][1].to_string(index=False)[1:-3]
        weight = int(re.search(r'\d+', weight).group())
        column.append(df[0].loc[df[0][0] == 'Weight'][1].to_string(index=False)[1:-3])
    except Exception:
        column.append('')
        pass


    #Package Weight
    column.append('')

    #Typical power consumption
    column.append('')

    #Maximum Power Consumption
    column.append('')

    #NET-ACL
    column.append('')

    #NET-ARP-Table-Size
    column.append('')

    #NET-FIBv4
    column.append('')

    #NET-FIBv6
    column.append('')


    #Foraward Perfomance
    try:
        perfomance = df[0].loc[df[0][0] == 'Performance'][1].to_string(index=False)
        forwardperfom = perfomance.find('Forwarding ')
        forwardperfom = perfomance[forwardperfom:]
        forwardperfom = int(re.search(r'\d+', forwardperfom).group())
        check = perfomance.find('Bpps')
        if (check != -1):
            forwardperfom = forwardperfom * 1000
        column.append(forwardperfom)
    except Exception:
        column.append('')
        pass

    #NET-Heat dispassion (BTU/h)
    column.append('')

    #Rack Units
    try:
        rack = df[0].loc[df[0][0] == 'Enclosure Type'][1].to_string(index=False)
        rack = int(re.search(r'\d+', rack).group())
        column.append(rack)
    except Exception:
        column.append('')
        pass

    #"NET-IP SLA (Full)
    column.append('')

    #NET-IP SLA Responder
    column.append('')

    #Routing
    features = df[0].loc[df[0][0] == 'Features'][1].to_string(index=False)
    devicetype = df[0].loc[df[0][0] == 'Device Type'][1].to_string(index=False)
    if (features.find('layer 2') != -1) or (devicetype.find('L3') != -1):
        column.append('+')
        column.append('+')
        column.append('+')
        column.append('+')
    else:
        column.append('')
        column.append('')
        column.append('')
        column.append('')

    #NET-L2 ERSPAN
    column.append('')

    #rspan
    if (features.find('RSPAN') != -1):
        column.append("+")
        column.append("+")
    else:
        column.append('')
        column.append('')

    #LACP
    if (features.find('LACP') != -1):
        column.append("+")
    else:
        column.append('')

    #NET-LAG M-LAG
    column.append('')

    #NET-LAG max groups
    column.append('')

    #NET-LAG Max Links
    column.append('')

    #Jumbo
    try:
        jumbo = df[0].loc[df[0][0] == 'Jumbo Frame Support'][1].to_string(index=False)[1:-6]
        jumbo = int(re.search(r'\d+', jumbo).group())
        column.append(jumbo)
    except Exception:
        column.append('')
        pass

    #JumboSupport
    if jumbo:
        column.append('+')
    else:
        column.append('')

    #LLDP
    if (features.find('LLDP') != -1):
        column.append("+")
    else:
        column.append('')

    #"NET-LAN LLDP-MED
    column.append('')

    #Max active VLAN
    try:
        capacity = df[0].loc[df[0][0] == 'Capacity'][1].to_string(index=False)
        vlan = capacity.find('VLAN')
        vlan = capacity[vlan:]
        vlan = int(re.search(r'\d+', vlan).group())
        column.append(vlan)
    except Exception:
        column.append('')
        pass

    #NET-LAN MST max instances
    column.append('')

    #MSTP
    if (features.find('MSTP') != -1):
        column.append("+")
    else:
        column.append('')

    #STP per Lan
    if (features.find('STP') != -1):
        column.append("+")
    else:
        column.append('')

    #"NET-LAN QinQ
    column.append('')

    #RSTP
    if (features.find('RSTP') != -1):
        column.append("+")
    else:
        column.append('')

    #NET-LAN Selective QinQ
    column.append('')

    #STP
    if (features.find('STP') != -1):
        column.append("+")
    else:
        column.append('')

    #BPDU
    if (features.find('BPDU') != -1):
        column.append("+")
    else:
        column.append('')

    #Portfast
    if (features.find('Portfast') != -1):
        column.append("+")
    else:
        column.append('')

    #"NET-LAN STP RootGuard
    if (features.find('STP') != -1):
        column.append("+")
    else:
        column.append('')

    #"NET-MAC Address Table Size
    column.append('')

    #Max Work temperature
    try:
        maxtemp = df[0].loc[df[0][0] == 'Max Operating Temperature'][1].to_string(index=False)
        maxtemp = int(re.search(r'\d+', maxtemp).group())
        maxtemp = (maxtemp - 32) / 1.8
        column.append(maxtemp)
    except Exception:
        column.append('')
        pass

    #Memory Ram
    try:
        memory = df[0].loc[df[0][0] == 'RAM'][1].to_string(index=False)
        memory = int(re.search(r'\d+', memory).group())*1000
        column.append(memory)
    except Exception:
        column.append('')
        pass

    #Min Work Temperature
    try:
        maxtemp = df[0].loc[df[0][0] == 'Min Operating Temperature'][1].to_string(index=False)
        maxtemp = int(re.search(r'\d+', maxtemp).group())
        maxtemp = (maxtemp - 32) / 1.8
        column.append(maxtemp)
    except Exception:
        column.append('')
        pass

    #CLI
    remote = df[0].loc[df[0][0] == 'Remote Management Protocol'][1].to_string(index=False)
    if (remote.find('CLI') != -1):
        column.append("+")
    else:
        column.append('')

    #"NET-MNG Event Manager(Scripting)
    column.append('')

    #Remote
    if (remote.find('NETCONF') != -1):
        column.append("+")
    else:
        column.append('')

    #SNMPv2
    if (remote.find('SNMP 2') != -1):
        column.append("+")
    else:
        column.append('')

    #SNMPv3
    if (remote.find('SNMP 3') != -1):
        column.append("+")
    else:
        column.append('')

    #SSHv2
    if (remote.find('SSH') != -1):
        column.append("+")
    else:
        column.append('')

    #Telnet
    if (remote.find('Telnet') != -1):
        column.append("+")
    else:
        column.append('')

    #"NET-MNG Web-management
    if (remote.find('WEB') != -1):
        column.append("+")
    else:
        column.append('')

    #"NET-MPLS L2 VPLS
    if (features.find('VPLS') != -1):
        column.append("+")
    else:
        column.append('')

    #MPLS l2 VPN
    if (features.find('VPN') != -1):
        column.append("+")
    else:
        column.append('')

    #"NET-MPLS L2 VPN (Kompela)
    if (features.find('Kompela') != -1):
        column.append("+")
    else:
        column.append('')

    #"NET-MPLS L2 VPN (Martini)
    if (features.find('Martini') != -1):
        column.append("+")
    else:
        column.append('')

    #L3VPN
    if (features.find('L3VPN') != -1):
        column.append("+")
    else:
        column.append('')

    #"NET-MPLS RSPV
    if (features.find('RSPV') != -1):
        column.append("+")
    else:
        column.append('')

    #MPLS
    if (features.find('MPLS') != -1):
        column.append("+")
    else:
        column.append('')

    #"NET-MPLS Traffic Engineering (TE)
    column.append('')

    #MTBF
    try:
        mtbf = df[0].loc[df[0][0] == 'MTBF'][1].to_string(index=False)
        mtbf = mtbf.replace(",","")
        mtbf = int(re.search(r'\d+', mtbf).group())
        column.append(mtbf)
    except Exception:
        column.append('')
        pass


    #IGMP multicast
    routing = df[0].loc[df[0][0] == 'Routing Protocol'][1].to_string(index=False)
    if (routing.find('IGMP') != -1) or (features.find('IGMP') != -1):
        column.append('+')
    else:
        column.append('')

    #IGMP snooping
    if (routing.find('IGMP') != -1) or (features.find('IGMP') != -1):
        column.append('+')
    else:
        column.append('')

    #multicast v4
    column.append('')

    #multicast v6
    column.append('')

    #"NET-Multicast PIM-DM
    if (routing.find('PIM-DM') != -1):
        column.append("+")
    else:
        column.append('')

    #PIM-SM
    if (routing.find('PIM-SM') != -1):
        column.append("+")
    else:
        column.append('')

    #PIM-SSM
    if (routing.find('PIM-SSM') != -1):
        column.append("+")
    else:
        column.append('')

    #"NET-Netflow v5
    if (features.find('Netflow') != -1):
        column.append("+")
    else:
        column.append('')
    #"NET-Netflow v9
    column.append('')

    #"NET-POE power budget
    column.append('')

    #"NET-Primary Power AC
    column.append('')

    #"NET-Primary Power DC
    column.append('')

    #"NET-Primary Power HVDC
    column.append('')

    #"NET-Primary power supply MAX voltage
    column.append('')

    #"NET-Primary power supply MIN voltage
    column.append('')

    #"NET-QOS HQOS Levels
    column.append('')

    #"NET-QOS Policing
    column.append('')

    #"NET-QOS Queue per Port
    column.append('')

    #"NET-QOS Shaping (GTS)
    column.append('')

    #Qos Support
    if (features.find('QoS') != -1):
        column.append("+")
    else:
        column.append('')

    #"NET-QOS Support HQOS
    column.append('')

    #"NET-QOS WRED
    column.append('')

    #"NET-Rack 19"" Moutable
    column.append('')

    #"NET-Routing BFD
    if (routing.find('BFD') != -1):
        column.append("+")
    else:
        column.append('')

    #BGP
    if (routing.find('BGP') != -1):
        column.append("+")
    else:
        column.append('')

    #"NET-Routing BGP-EVPN
    if (routing.find('EVPN') != -1):
        column.append("+")
    else:
        column.append('')

    #"NET-Routing BGP4+
    column.append('')

    #EIGRP
    if (routing.find('EIGRP') != -1):
        column.append("+")
    else:
        column.append('')

    #"NET-Routing EIGRP Stub
    column.append('')

    #ISIS
    if (routing.find('IS-IS') != -1):
        column.append("+")
    else:
        column.append('')

    #"NET-Routing ISISv6
    column.append('')

    #"NET-Routing NAPT
    if (features.find('NAPT') != -1):
        column.append("+")
    else:
        column.append('')

    #Routing Nat
    if (features.find('NAT') != -1):
        column.append("+")
    else:
        column.append('')

    #"NET-Routing NAT-ALG
    if (features.find('ALG') != -1):
        column.append("+")
    else:
        column.append('')

    #OSPF
    if (routing.find('OSPF') != -1):
        column.append("+")
    else:
        column.append('')

    #"NET-Routing OSPF Stub
    column.append('')

    #"NET-Routing OSPFv3
    if (routing.find('OSPFv3') != -1):
        column.append("+")
    else:
        column.append('')

    #"NET-Routing PBR
    if (routing.find('PBR') != -1):
        column.append("+")
    else:
        column.append('')

    #RIP
    if (routing.find('RIP') != -1):
        column.append("+")
    else:
        column.append('')

    #RIPng
    if (routing.find('RIPng') != -1):
        column.append("+")
    else:
        column.append('')

    #"NET-Routing SR
    if (routing.find('SR') != -1):
        column.append("+")
    else:
        column.append('')

    #"NET-Routing SRv6
    if (routing.find('SRv6') != -1):
        column.append("+")
    else:
        column.append('')

    #VRF
    if (features.find('VRF') != -1):
        column.append("+")
    else:
        column.append('')

    #NET-Routing VRF max number
    column.append('')

    #VRRP
    if (features.find('VRRP') != -1):
        column.append("+")
    else:
        column.append('')

    #"NET-Routing VRRPv6
    column.append('')
    #"NET-Secondary Power AC
    column.append('')
    #"NET-Secondary Power DC
    column.append('')
    #"NET-Secondary Power HVDC
    column.append('')
    #"NET-Secondary power supply MAX voltage
    column.append('')
    #"NET-Secondary power supply MIN voltage
    column.append('')

    #802.1x
    standart = df[0].loc[df[0][0] == 'Compliant Standards'][1].to_string(index=False)
    if (standart.find('802.1x')):
        column.append('+')
    else:
        column.append('')

    #Radius
    if (features.find('RADIUS') != -1):
        column.append("+")
    else:
        column.append('')

    #Tacacs
    auth = df[0].loc[df[0][0] == 'Authentication Method'][1].to_string(index=False)
    if (auth.find('TACACS')):
        column.append('+')
    else:
        column.append('')

    #MACsec
    if (features.find('MACsec') != -1):
        column.append('+')
    else:
        column.append('')

    #"NET-Security Port Security
    if (features.find('SPS') != -1):
        column.append('+')
    else:
        column.append('')

    #"NET-Service Length (Month)
    column.append('')
    #"NET-sFlow
    if (features.find('sFlow') != -1):
        column.append('+')
    else:
        column.append('')

    #"NET-Stack Max Bandwidth
    column.append('')

    #"NET-Stack Max Device
    column.append('')

    #"NET-Stack Support
    column.append('')

    #"NET-Stack Support Service Port
    column.append('')
    #"NET-Stack Support Special Port
    column.append('')


    #Storage
    try:
        storage = df[0].loc[df[0][0] == 'Flash Memory'][1].to_string(index=False)
        storage = int(re.search(r'\d+', storage).group())*1000
        column.append(storage)
    except Exception:
        column.append('')
        pass


    #"NET-Subscription Length (Month)
    column.append('')

    #SVI
    try:
        svi = capacity.find('SVI')
        svi = capacity[svi:]
        svi = int(re.search(r'\d+', svi).group())
        if svi:
            column.append(svi)
    except Exception:
        column.append('')
        pass

    #SW FAN number
    column.append('')

    #Switching Capacity
    try:
        swcap = perfomance.find('Switching capacity')
        swcap = perfomance[swcap:]
        swcap = swcap.replace(".","")
        swcap = int(re.search(r'\d+', swcap).group())*100
        column.append(swcap)
    except Exception:
        column.append('')
        pass

    #"NET-Tunnel GRE
    column.append('')
    #"NET-Tunnel GRE Max Number
    column.append('')

    #"NET-VXLAN
    if (features.find('VXLAN') != -1):
        column.append('+')
    else:
        column.append('')

    #"NET-VXLAN-Gateway
    column.append('')
    #"NET-Warranty Length (Month)
    column.append('')
    #"NET-WLAN AC Functionality
    column.append('')
    #"NET-WLAN AP managed
    column.append('')
    #"SPort 1000Base-X
    column.append('')
    #"SPort 10GBase-X
    column.append('')
    #"Port 10/100/1000Base-T
    column.append('')
    #"Port 100/1000Base-T
    column.append('')
    #"Port 10/100/1000Base-T POE
    column.append('')
    #"Port 10/100/1000Base-T POE+
    column.append('')
    #"Port 10/100/1000Base-T POE++
    column.append('')
    #"Port 100/1000Base-X
    column.append('')
    #"Port 100/1000/10GBase-X
    column.append('')
    #"Port 1000/10GBase-X
    column.append('')
    #"CPort 10/100/1000Base-T or 100/1000Base-X
    column.append('')
    #"SPort 40G-QSFP
    column.append('')
    #"SPort 40G-QSFP-s-SPLIT
    column.append('')
    #"SPort 100G-QSFP28
    column.append('')
    #"SPort 100G-QSFP28-s-SPLIT
    column.append('')
    #"Port 40/100G
    column.append('')
    #"Port 40/100G-s-SPLIT
    column.append('')
    #"NET-Outband-Eth-Management Port
    column.append('')
    #"NET-HDD-Storage-Slot
    column.append('')
    #"NET-SW-Fixed-Interface-Slot
    column.append('')
    #"NET-FAN-Slot
    column.append('')
    #"NET-Secondary-Power-Supply-Slot
    column.append('')
    #"NET-Primary-Power-Supply-Slot
    column.append('')
    #"NET-Software-Slot
    column.append('')
    #"NET-Subscription-Slot
    column.append('')
    #"NET-Primary-PS-Socket-C14
    column.append('')
    #"NET-Secondary-PS-Socket-C14
    column.append('')
    #"NET-Primary-PS-Socket-C20
    column.append('')
    #"NET-Secondary-PS-Socket-C20
    column.append('')
    #"NET-Primary-PS-Socket-DC
    column.append('')
    #"NET-Secondary-PS-Socket-DC
    column.append('')


    #USB
    interfaces = df[0].loc[df[0][0] == 'Interfaces'][1].to_string(index=False)
    if (interfaces.find('USB') != -1):
        column.append('+')
    else:
        column.append('')

    #"Port 100/1000/2.5GBase-T POE++
    column.append('')
    #"Port 100/1000/2.5G/5G POE++
    column.append('')
    #"Port 100/1000/2.5G/5G/10G POE++
    column.append('')
    #"Port 1000/10G/25GBase-X
    column.append('')
    #"Port 1000/2.5G/5G/10GBase-T
    column.append('')
    #"Port 10/100/1000Base-X
    column.append('')
    #"NET-Primary-PS-Socket-C16
    column.append('')
    #"NET-Stack-Port
    column.append('')
    #"Stack Power Port
    column.append('')
    #"Console USB
    interfaces = df[0].loc[df[0][0] == 'Interfaces'][1].to_string(index=False)
    if (interfaces.find('USB') != -1):
        column.append('+')
    else:
        column.append('')
    #"RJ45
    interfaces = df[0].loc[df[0][0] == 'Interfaces'][1].to_string(index=False)
    if (interfaces.find('RJ45') != -1):
        column.append('+')
    else:
        column.append('')
    #Port 10/1000Base
    column.append('')
    #Port 10/100 poe
    column.append('')
    #Sport 1000Base
    column.append('')
    #Port 100/1000/10GBase
    column.append('')
    #SW POWER-RPS
    column.append('')
    #10/1000/2.5/5/10
    column.append('')
    #Port 10/100 POE+
    column.append('')
    #RPS-POWER-supply Slot
    column.append('')



  #print(forwardperfom)
    for i,e in enumerate(column):
        sheet1.write(wrk+1,i,e )



#Дополнительные устройства
name = "mira.xls"
print('Switch parsed successfully')
print('Starting parse dops.')
dop=0
dopwrk=wrk
for dop in range(len(dopnew)):
    try:
        df=pd.read_html(url+dopnew[dop])
    except Exception:
        pass
    pd.options.display.max_colwidth = 10000
    r = requests.get(url+dopnew[dop])
    root = LH.fromstring(r.content)

    column1 = list()

    #Тип устройства
    column1.append('dop')

    #Производитель
    column1.append('Cisco')

    #Серия
    for i in range(len(df)):
        if df[i].loc[df[i][0] == 'Designed For'][1].to_string(index=False) != 'Series([], )':
            serial = df[i].loc[df[i][0] == 'Designed For'][1].to_string(index=False)
    column1.append(serial)

    #Подсерия
    column1.append('')


    #Партномер
    namedop = root.xpath('//*[@id="shopify-section-static-product"]/section/article/div[2]/div[1]/div[3]/span/text()')
    column1.append(namedop[0])

    #Название
    column1.append(namedop[0])

    #Описание
    briefdop = root.xpath('//*[@id="shopify-section-static-product"]/section/article/div[2]/div[3]/p[3]/text()')
    column1.append(briefdop[0])

    for i,e in enumerate(column1):
        sheet1.write(dopwrk+2,i,e)
    dopwrk=dopwrk+1
print('Dops parsed successfully')
book.save(name)
book.save(TemporaryFile())