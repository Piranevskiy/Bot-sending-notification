import os
from datetime import datetime, timedelta
from selenium import webdriver
import requests
os.chdir(r"C:\Users\Desktop\sqldeveloper\sqldeveloper\instantclient_19_6")
import cx_Oracle
import time
import win32com.client
import telepot

SetProxy = telepot.api.set_proxy('http://116.203.254.70:3129')
bot = telepot.Bot('token')
def mess(mess):
    try:
        bot.sendMessage(chat_id, mess)
    except:
        sheet1.cells(rw, 1).value = str(datetime.now()) + ' - Отправка уведомления не удалась.'
        print(str(datetime.now()) + ' - Отправка уведомления не удалась.')
ex_fil = r'C:\Users\Desktop\IPTV\logs_iptv.xlsx'
# Excel create

Excel = win32com.client.Dispatch("Excel.Application")
wb = Excel.Workbooks.Open(ex_fil)
sheet1 = wb.Sheets('logs')
sheet2 = wb.Sheets('attempts')
rw = 1
while sheet1.cells(rw, 1).value is not None:
    rw += 1

try:
    requests.get('http://www.google.com/')
except:
    sheet1.cells(rw, 1).value = str(datetime.now()) + " - Нет подключения к Интернет сети"
    rw += 1
    print(str(datetime.now()) + " - Нет подключения к Интернет сети")
    wb.Save()  # сохраняем рабочую книгу
    wb.Close()  # закрываем ее
    raise SystemExit(1)

os.environ["NLS_LANG"] = ".AL32UTF8"  # Кодировка для чтения русских символов
log = 'login'
passw = 'pass'
cou = 2
slovar, spisok1, spisok2 = {}, [], []

ORACLE_CONNECT = "login/pass@(DESCRIPTION=(TRANSPORT_CONNECT_TIMEOUT=5)(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=host)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=DBName)))"

# Попытка подключиться к БД
try:
    orcl = cx_Oracle.connect(ORACLE_CONNECT)
except:
    # Шаг для подключения к SSL
    sheet1.cells(rw, 1).value = str(datetime.now()) + " - Не подключен к capsule"
    rw += 1
    print(str(datetime.now()) + " - Не подключен к capsule")
    url = "https://capsule.ru"
    driverChr = r'C:\Users\ssl\chromedriver.exe'
    # Указываем полный путь к geckodriver.exe на ПК.
    driver = webdriver.Chrome(driverChr)
    driver.get(url)
    time.sleep(3)

    # Ввод логина
    login = driver.find_element_by_id("userName")
    login.clear()
    login.send_keys("login")

    # Ввод пароля
    pswd = driver.find_element_by_id("passwordDisplayed")
    pswd.send_keys("pass")

    # Жмем войти
    driver.find_element_by_id("LoginButton").click()

    time.sleep(15)  # Ждет загрузки браузера
    try:
        orcl = cx_Oracle.connect(ORACLE_CONNECT)
    except:
        sheet1.cells(rw, 1).value = str(datetime.now()) + " - 2ая попытка подключения к capsule"
        rw += 1
        print(str(datetime.now()) + " - 2ая попытка подключения к capsule")
        time.sleep(10)  # Ждет загрузки браузера
        try:
            orcl = cx_Oracle.connect(ORACLE_CONNECT)
        except:
            sheet1.cells(rw, 1).value = str(datetime.now()) + " - Не удалось соедениться с capsule"
            rw += 1
            print(str(datetime.now()) + " - Не удалось соедениться с capsule")
            wb.Save()  # сохраняем рабочую книгу
            wb.Close()  # закрываем ее
            raise SystemExit(1)

sql = '''
select * from load_control a
where a.fact_name = 'sdp_log'
order by a.dt desc, a.fact_name, a.change_dttm desc
fetch first 24 rows only
              '''
cursor = orcl.cursor()
cursor.execute(sql)
for row in cursor:
    spisok1.append(row)

now = datetime.now()
nowst = str(datetime.now())[:10]
yesterday = now - timedelta(days=1)
yestdate = str(yesterday)[:10]

# Проверка времени между загрузкой логов SmartSpy
if str(now - spisok1[0][3]) > '1:20:00':
    sheet1.cells(rw, 1).value = str(datetime.now()) + ' - Внимание разница между загрузкой SmartSpy logs составляет ' + str(now - spisok1[0][3])
    rw += 1
    sheet2.cells(2, 2).value = sheet2.cells(2, 2).value + 1
    print(str(datetime.now()) + ' - Внимание разница между загрузкой SmartSpy logs составляет ' + str(now - spisok1[0][3]))
    if sheet2.cells(2, 2).value < 3:
        mess(str(datetime.now()) + ' - Внимание разница между загрузкой SmartSpy logs составляет ' + str(now - spisok1[0][3]))
else:
    sheet1.cells(rw, 1).value = str(datetime.now()) + ' - Разница загрузки менее часа - ' + str(now - spisok1[0][3]) + ' Статус последнего лога: ' + str(
        spisok1[0][2]) + ' - ' + str(spisok1[0][0])
    rw += 1
    if sheet2.cells(2, 2).value > 0:
        print('Загрузка SmartSpy logs восстановлена')
        mess(str(datetime.now()) + ' Загрузка SmartSpy logs восстановлена')
    sheet2.cells(2, 2).value = 0
    print(str(datetime.now()) + ' - Разница загрузки менее часа - ' + str(now - spisok1[0][3]) + ' Статус последнего лога: ' + str(
        spisok1[0][2]) + ' - ' + str(spisok1[0][0]))

# Проверка статуса предпоследнего лога SmartSpy
if spisok1[1][2] == 'HIVE_COMPLETE':
    sheet1.cells(rw, 1).value = str(datetime.now()) + ' - Внимание статус предыдущего лога ' + str(spisok1[1][2]) + ' необходимо перезапустить джоб RIC_SDP_LOG_HOURLY'
    rw += 1
    sheet2.cells(9, 2).value = sheet2.cells(9, 2).value + 1
    print(str(datetime.now()) + ' - Внимание статус предыдущего лога ' + str(spisok1[1][2]) + ' необходимо перезапустить джоб RIC_SDP_LOG_HOURLY')
    if sheet2.cells(9, 2).value < 3:
        mess(str(datetime.now()) + ' - Внимание статус предыдущего лога ' + str(spisok1[1][2]) + ' необходимо перезапустить джоб RIC_SDP_LOG_HOURLY')
else:
    sheet1.cells(rw, 1).value = str(datetime.now()) + ' - Статус предыдущего лога ' + str(spisok1[1][2])
    rw += 1
    print(str(datetime.now()) + ' - Статус предыдущего лога ' + str(spisok1[1][2]))

dict_sec = {'dbview': [yestdate], 'dbview2': [yestdate], 'sdp_log_agg': [yestdate], 'sdp_log_daily': [yestdate],
            'stb_traffic': [yestdate], 'term_state': [yestdate]}

sql = '''
select * from (
select * from load_control a
where a.fact_name <> 'sdp_log'
order by a.dt desc, a.fact_name, a.change_dttm desc
fetch first 6 rows only) sub
where sub.dt >= (select max(dt) from load_control where fact_name <> 'sdp_log')
              '''
cursor = orcl.cursor()
cursor.execute(sql)
counter, c = 0, 0
for row in cursor:
    spisok2.append(row)
    if dict_sec[spisok2[counter][1]][0] == str(spisok2[counter][0])[:10]:

        c += 1
        sheet1.cells(rw, 1).value = str(datetime.now()) + ' - Лог ' + str(spisok2[counter][1]) + ' за ' + str(
        spisok2[counter][0]) + ' успешно загружен. Время загрузки ' + str(spisok2[counter][3]) + ' статус - ' + str(
        spisok2[counter][2])
        rw += 1
        # Обнуление попыток
        if spisok2[counter][1] == 'sdp_log_agg':
            if sheet2.cells(5, 2).value > 0:
                print('Загрузка sdp_log_agg восстановлена')
                mess(str(datetime.now()) + ' Загрузка sdp_log_agg восстановлена')
            sheet2.cells(5, 2).value = 0
        elif spisok2[counter][1] == 'sdp_log_daily':
            if sheet2.cells(6, 2).value > 0:
                print('Загрузка sdp_log_daily восстановлена')
                mess(str(datetime.now()) + ' Загрузка sdp_log_daily восстановлена')
            sheet2.cells(6, 2).value = 0
        elif spisok2[counter][1] == 'stb_traffic':
            if sheet2.cells(7, 2).value > 0:
                print('Загрузка stb_traffic восстановлена')
                mess(str(datetime.now()) + ' Загрузка stb_traffic восстановлена')
            sheet2.cells(7, 2).value = 0
        elif spisok2[counter][1] == 'term_state':
            if sheet2.cells(8, 2).value > 0:
                print('Загрузка term_state восстановлена')
                mess(str(datetime.now()) + ' Загрузка term_state восстановлена')
            sheet2.cells(8, 2).value = 0
        else:
            if sheet2.cells(3, 2).value > 0 or sheet2.cells(4, 2).value > 0:
                print('Загрузка dbview восстановлена')
                mess(str(datetime.now()) + ' Загрузка dbview восстановлена')
            sheet2.cells(3, 2).value = 0
            sheet2.cells(4, 2).value = 0
        print(*spisok2[counter])
    else:
        # Проверяем 'stb_traffic' и 'term_state'
        if str(now) > str(nowst) + ' 02:00:00' and (
                spisok2[counter][1] == 'stb_traffic' or
                spisok2[counter][1] == 'term_state'):
            sheet1.cells(rw, 1).value = str(datetime.now()) + ' Отсутствуют логи по ' + spisok2[counter][1] +' за ' + str(yestdate)
            rw += 1
            # Проставляем попытки
            if spisok2[counter][1] == 'stb_traffic':
                sheet2.cells(7, 2).value = sheet2.cells(7, 2).value + 1
                if sheet2.cells(7, 2).value < 3:
                    mess(str(datetime.now()) + ' Отсутствуют логи по ' + spisok2[counter][1] + ' за ' + str(yestdate))
            elif spisok2[counter][1] == 'term_state':
                sheet2.cells(8, 2).value = sheet2.cells(8, 2).value + 1
                if sheet2.cells(8, 2).value < 3:
                    mess(str(datetime.now()) + ' Отсутствуют логи по ' + spisok2[counter][1] + ' за ' + str(yestdate))
            print(str(datetime.now()) + ' Отсутствуют логи по ' + spisok2[counter][1] +' за ' + str(yestdate))
        # Проверяем 'dbview'
        elif str(now) > str(nowst) + ' 09:00:00' and spisok2[counter][1] == 'dbview':
            cursflag = orcl.cursor()
            cursflag.execute('select * from  iptv_md.etl_flags')
            res = cursflag.fetchall()
            sheet1.cells(rw, 1).value = str(datetime.now()) + ' - Внимание логи dbview не загружены, статус флага - ' + str(res[0][2]) + ' последнее обновление ' + str(res[0][4])
            rw += 1
            sheet2.cells(3, 2).value = sheet2.cells(3, 2).value + 1
            sheet2.cells(4, 2).value = sheet2.cells(4, 2).value + 1
            print(str(datetime.now()) + ' - Внимание логи dbview не загружены, статус флага - '
                  + str(res[0][2]) + ' последнее обновление ' + str(res[0][4]))
            if sheet2.cells(3, 2).value < 3 or sheet2.cells(4, 2).value < 3:
                mess(str(datetime.now()) + ' - Внимание логи dbview не загружены, статус флага - '
                  + str(res[0][2]) + ' последнее обновление ' + str(res[0][4]))
            # Проверяем 'sdp_log_agg' и 'sdp_log_daily'
        elif str(now) > str(nowst) + ' 09:00:00'  and (
                spisok2[counter][1] == 'sdp_log_agg' or spisok2[counter][1] == 'sdp_log_daily'):
            sheet1.cells(rw, 1).value = str(datetime.now()) + ' Отсутствуют логи по ' + spisok2[counter][1] +' за ' + str(yestdate)
            rw += 1
            if spisok2[counter][1] == 'sdp_log_agg':
                sheet2.cells(5, 2).value = sheet2.cells(5, 2).value + 1
                if sheet2.cells(5, 2).value < 3:
                    mess(str(datetime.now()) + ' Отсутствуют логи по ' + spisok2[counter][1] + ' за ' + str(yestdate))
            elif spisok2[counter][1] == 'sdp_log_daily':
                sheet2.cells(6, 2).value = sheet2.cells(6, 2).value + 1
                if sheet2.cells(6, 2).value < 3:
                    mess(str(datetime.now()) + ' Отсутствуют логи по ' + spisok2[counter][1] + ' за ' + str(yestdate))
            print(str(datetime.now()) + ' Отсутствуют логи по ' + spisok2[counter][1] +' за ' + str(yestdate))

    counter += 1
print(c)
# Проверка времени между загрузкой логов SmartSpy


cursor.close()  # закрываем курсор
orcl.close()  # закрываем соединение с БД

wb.Save()  # сохраняем рабочую книгу
wb.Close()  # закрываем ее