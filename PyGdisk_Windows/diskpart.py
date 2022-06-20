import win32com.client
from re import search
import re

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case,
                                    # the inbox. You can change that number to reference
                                    # any other folder
messages = inbox.Items
message = messages.GetLast()
body_content = message.body.lower()
#print (body_content)                   #отладочный кусок выводит текст письма
match = re.search(re.escape('контактный телефон') + '.*',body_content  ).group() # регулярка для поиска по заданному шаблону
if match:
    #print(match)                  #отладочный кусок выводит совпадение
    result = match
    if result =='null':
        print('ALARM')
    aresult =re.findall('\d{5}',result)
    for index in range(len(aresult)):
        try:
            print(aresult[index])
            b =  int(aresult[index])
            if b > 12000:
                print('its worked')
                if b == 12040:
                    print('oops i did it again')


        except:
            print('не номер телефона')




#else:
#    message.Delete()