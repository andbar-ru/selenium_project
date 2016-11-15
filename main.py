#!/usr/bin/python
# coding: utf-8

# Imports {{{1
from os import path
import os
import shutil
import re
import sys
import getopt
import logging
import csv
from collections import OrderedDict
from time import time as T
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, WebDriverException, UnexpectedAlertPresentException, NoSuchWindowException

#}}}
# Global variables and configuration {{{1
SCRIPT = path.basename(sys.argv[0])
ARGS = sys.argv[1:]
FILE = ARGS[0]
with open(FILE) as F:
    # по мере бронирования этот словарь будет сокращаться
    LINES = OrderedDict()
    for n,line in enumerate(F.readlines()):
        LINES[n] = line
SCENARIO = 'reserve'

# Конфиг
SITE = 'http://ticket.tzar.ru'
LOGLEVEL = logging.INFO
# Данные авторизации
LOGIN = 'david@rusnavigator.ru'
PASSWORD = 'dummy_password'
# Минимальный процент или количество, при котором бронируем частично
MIN_PERCENT = 50
MIN_QUOTA = 50
# Диапазон поиска ближайших времён в минутах, если ничего нет, переходим на другой маршрут
SEARCH_RANGE = 180
# Количество маршрутов
NUM_ROUTES = 3
# Регулярки
DATE_RE = re.compile(r'^\d{2}\.\d{2}\.\d{4}$')
TIME_RE = re.compile(r'^\d{2}:\d{2}$')
# Файлы для регистрации неудач
if not path.isdir('log'):
    os.mkdir('log')
FAILEDFILE = 'log/failed.log'
FORWARDEDFILE = 'log/forwarded.log'
MISSEDFILE = 'log/missed.log'

OVERLIMIT_MESSAGE = u'Превышение доступного кол-ва бронирования на данный месяц!!'

# Текущая дата на сайте
reservationDate = None
# Список заказов (важен порядок при извлечении дат и маршрутов). Сокращается по мере выполнения.
orderlist = []
# Ожидание открытия расписания. Как только расписание откроется, становится False
waitOpening = True

# Информация о неудачах
# Пропущенные заказы
failedOrders = []
# Сроки, которые забронировались на другое время и/или маршршут
forwardedOrders = []

# Настройка журналирования
# LOGFILE = path.realpath(LOGFILE)
# logging.basicConfig(filename=LOGFILE, format='%(levelname)s:%(message)s', filemode='w', level=LOGLEVEL)
logging.basicConfig(format='%(levelname)s:%(message)s', level=LOGLEVEL) # stdout

#}}}
def profile(func): #{{{1
    def wrap(*args, **kwargs):
        t0 = T()
        func(*args, **kwargs)
        t1 = T()
        logging.info("%.6f" % (t1-t0))
    return wrap

#}}}
class ReservationDate(object): #{{{1
    """Объект даты, на которую производится резервирование. Singleton.
    Единственный экземпляр - глобальная переменная reservationDate.
    Хранит маршруты, свойства маршрутов.
    """
    def __new__(cls, Date): #{{{2
        global reservationDate
        if reservationDate is None:
            logging.info(u'Создание нового объекта ReservationDate.')
            reservationDate = super(ReservationDate, cls).__new__(cls)
        return reservationDate

    #}}}
    def __init__(self, Date): #{{{2
        if not hasattr(self, 'date') or self.date != Date:
            logging.info(u'Присваивание новой даты %s объекту ReservationDate.' % Date)
            self.date = Date
            # на новой дате маршруты не определёны
            self.routes = {}
            for i in range(1, NUM_ROUTES+1):
                self.routes[str(i)] = None
            self.route = None # текущий маршрут
            # Перейти на дату в браузере
            datepicker = driver.find_element_by_id("datepicker1")
            datepicker.clear()
            datepicker.send_keys(Date, Keys.RETURN)

    #}}}
    def decide_route(self, route): #{{{2
        """Определяем маршрут согласно доступности и предпочтению route.
        Также парсим таблицу, исследуя доступность билетов по временам.
        """
        global waitOpening
        # Если все маршруты недоступны {{{3
        if not self.routes:
            return False
        #}}}
        # Исключаем маршруты, где не осталось доступных времён {{{3
        for r in self.routes:
            if type(r) is dict and r['timesAvail'] == []:
                del self.routes[r]
        #}}}
        # Если route уже является текущим маршрутом {{{3
        if route in self.routes and route == self.route:
            return self.route
        #}}}
        # Выбор альтернативного маршрута (если маршрут недоступен и уже удалён) {{{3
        if route not in self.routes:
            logging.info(u"Поиск альтернативного маршрута")
            if route == '1': # от 10:00 до 11:45
                if '2' in self.routes:
                    route = '2'
                elif '3' in self.routes:
                    route = '3'
                else: # self.routes == {}
                    return False
            elif route == '2': # от 12:00 до 18:40
                if '3' in self.routes:
                    route = '3'
                elif '1' in self.routes:
                    route = '1'
                else: # self.routes == {}
                    return False
            elif route == '3': # от 12:00 до 17:40
                if '2' in self.routes:
                    route = '2'
                elif '1' in self.routes:
                    route = '1'
                else: # self.routes == {}
                    return False
            else:
                logging.error("Unexpected route: %s" % route)
                return False
        #}}}
        def get_timetableTrs(): # {{{3
            wait_and_do(EC.element_to_be_clickable((By.XPATH, '//select[@id="place"]/option[%s]' % route)), 'click')
            # Нажать кнопку "Найти"
            wait_and_do(EC.element_to_be_clickable((By.ID, 'submitBtn')), 'click')
            # Проверяем, что на текущей странице есть таблица с расписанием.
            wait_and_do(EC.visibility_of_element_located((By.CLASS_NAME, 'timetable')), '')
            timetableTrs = driver.find_elements_by_css_selector('tr.body:not([style])')
            return timetableTrs
        #}}}
        # Переходим на маршрут route {{{3
        logging.info(u'Переход на маршрут %s' % route)
        if waitOpening:
            while True:
                timetableTrs = get_timetableTrs()
                if timetableTrs == []: # Таблицы с расписанием нет, маршрут недоступен
                    sleep(1)
                    logging.info(u'Ждём открытия расписания')
                    continue
                else:
                    logging.info(u'Расписание открылось')
                    waitOpening = False
                    break
        else:
            timetableTrs = get_timetableTrs()
            timetableTrs = driver.find_elements_by_css_selector('tr.body:not([style])')
            if timetableTrs == []: # Таблицы с расписанием нет, маршрут недоступен
                logging.warning(u"На дату %s на маршрут %s нет таблицы с расписанием." % (self.date, route))
                # Удаляем маршрут из списка
                del self.routes[route]
                # Ветка 'Выбор альтернативного маршрута'
                return self.decide_route(route)
        #}}}
        # Формируем словарь и список доступных времён {{{3
        if type(self.routes[route]) is dict: # словарь и список на этот маршрут уже имеются
            logging.info(u'Словарь и список доступных времён на маршрут %s уже имеется.' % route)
            self.route = route
            return self.route

        logging.info(u'Формируем словарь и список доступных времён на маршрут %s.' % route)
        timesAvailDict = {} # {Время:[Кол-во, Бронь],...}
        timesAvail = [] # [(Время, кол-во минут),...]
        for tr in timetableTrs:
            try:
                Time, quota, reservation = tr.text.split()[1:4]
                quota = int(quota)
                reservation = int(reservation)
                # Вносим только те времена, на которые есть возможность ЕЩЁ забронировать
                if quota > 0:
                    timesAvailDict[Time] = [quota, reservation]
                    timesAvail.append( (Time, str2minutes(Time)) )
                else:
                    continue # следующий tr
            except ValueError:
                logging.warning(u'В дате %s на время %s маршрута %s в ячейке "Кол-во" и/или "Бронь" '
                    u'обнаружено нечисло. Это время не будет использоваться.' % (self.date, Time, route))
                continue # следующий tr
        #}}}
        # Если словарь пустой (все билеты забронированы) {{{3
        if not timesAvailDict:
            logging.warning(u'На дату %s на маршрут %s нет доступных билетов. '
                u'Все билеты забронированы?' % (self.date, route))
            # Удаляем маршрут из списка
            del self.routes[route]
            # Ветка 'Выбор альтернативного маршрута'
            return self.decide_route(route)
        #}}}
        # Устанавливаем текущий маршрут {{{3
        self.route = route
        self.routes[route] = {'timesAvailDict':timesAvailDict, 'timesAvail':timesAvail}

        return self.route

    #}}}
    def unreserve_all(self): #{{{2
        """Снять все брони со всех маршрутов."""
        for route in self.routes:
            wait_and_do(EC.element_to_be_clickable((By.XPATH, '//select[@id="place"]/option[%s]' % route)), 'click')
            # Нажать кнопку "Найти"
            wait_and_do(EC.element_to_be_clickable((By.ID, 'submitBtn')), 'click')
            # Проверка того, что загрузилась таблица с расписанием
            wait_and_do(EC.visibility_of_element_located((By.CLASS_NAME, 'timetable')), '')
            timetableTrs = driver.find_elements_by_css_selector('tr.body:not([style])')
            # Таблицы с расписанием нет
            if timetableTrs == []:
                logging.warning(u"На дату %s на маршрут %s нет таблицы с расписанием." % (self.date, route))
                continue

            # Таблица с расписанием загружена. Ищем и удаляем брони по одному.
            # Так как таблица обновляется после каждого удаления, то после каждого удаления
            # надо репарсить таблицу в поисках следующей брони
            while True:
                tds_with_reservation = driver.find_elements_by_xpath('//table[@class="timetable"]//tr[@class="body"][not(@style)]/td[4][text()!="0"]')

                if tds_with_reservation == []:
                    logging.info(u"На дату %s на маршрут %s брони (больше) нет." % (Date, route))
                    # переход на другой маршрут или конец обработки даты
                    if route == '1':
                        logging.info(u"Переход на маршрут 2")
                    else:
                        logging.info(u'Конец обработки даты %s' % Date)
                    break

                # так как после удаления tds устареют, то берём первый и continue
                td = tds_with_reservation[0]
                # чтобы не зависало на окне confirm
                script = 'window.confirm = function() {return true;}'
                driver.execute_script(script)
                td.click()
                # в ячейке появляется div#options_mod
                delete = td.find_element_by_id('delete')
                delete.click()
                # окно confirm должно пропускаться
                # здесь выполняется ajax ...
                window = WebDriverWait(driver, 10, 0.1).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div.ui-dialog')))
                title = window.find_element_by_css_selector('span.ui-dialog-title').text
                window.find_element_by_tag_name('button').click()
    #}}}
#}}}
class Order(object): #{{{1
    def __init__(self, lineNum, Date, route, Time, quantity): #{{{2
        # номер строки в исходном файле
        self.lineNum = lineNum
        # дата заказа
        self.date = Date
        # маршруты
        self.routesPassed = {} # пройденные маршруты: {route:[closestTime, diffInMins],...}
        self.route0 = route # изначальный
        self.route = route # текущий
        # время, относительно которого ищутся ближайшие времена
        self.time0 = Time
        # текущее время бронировании, при размазывании меняется
        self.time = Time
        # количество билетов, которое необходимо забронировать
        self.quantity0 = quantity
        self.quantity = quantity # оставшиеся
        # стиль поведения: 'spread' - размазываться по одному маршруту
        self.style = 'base'
        # статус после обработки
        self.statuses = []

    #}}}
    def reserve(self): #{{{2
        """Инициирование процесса резервирования заказа"""
        # Проверка на исчерпание заказа {{{3
        if self.quantity <= 0:
            return True
        #}}}
        # Определение объекта даты {{{3
        global reservationDate
        reservationDate = ReservationDate(self.date)
        #}}}
        # Определение маршрута {{{3
        if self.route != reservationDate.route:
            result = reservationDate.decide_route(self.route)
            if result is False: # пропуск заказа
                logging.warning(u"Все маршруты на дату %s недоступны, пропуск заказа" % reservationDate.date)
                return False
            # текущий маршрут
            self.route = result
        #}}}
        # timesAvail[Dict] {{{3
        # К этому моменты д.б. готовы. Это ссылки на объекты, принадлежащие reservationDate.
        timesAvailDict = reservationDate.routes[self.route]['timesAvailDict']
        timesAvail = reservationDate.routes[self.route]['timesAvail']
        #}}}
        # Проверка на исчерпание доступных времён {{{3
        if timesAvail == []: # переопределяем маршрут
            del reservationDate.routes[self.route]
            result = reservationDate.decide_route(self.route)
            if result is False:
                logging.warning(u"Все маршруты на дату %s исчерпаны, пропуск заказа" % self.date)
                return False
            self.route = result
            return self.reserve()
        #}}}
        # Если время доступно {{{3
        if self.time in timesAvailDict:
            quota, reservation = timesAvailDict[self.time]
            # Во время бронирования quantity и quota увеличиваются на значение брони
            operQuantity = self.quantity + reservation
            operQuota = quota + reservation
 
            # Доступного количества хватает {{{4
            if quota >= self.quantity:
                # Иначе "Превышено доступное количество" (глюк сайта)
                if operQuantity > quota:
                    self.unreserve_time(self.time)
                # бронируем и заканчиваем обработку заказа на это время
                result = self.reserve_time(self.time, operQuantity)
                if result == True:
                    if operQuantity == operQuota: # время исчерпано, убираем его
                        self.delete_time(timesAvailDict, timesAvail, self.time)
                    else: # время не исчерпано, убавляем у него квоту и прибавляем бронь
                        timesAvailDict[self.time][0] -= self.quantity
                        timesAvailDict[self.time][1] += self.quantity
                    self.statuses.append('OK')
                    return True
                elif result == False:
                    logging.error(u"Не удалось забронировать билет на время %s даты %s." % (self.time, self.date))
                    return False
                elif result == OVERLIMIT_MESSAGE: # Превышение доступного кол-ва бронирования на данный месяц
                    logging.warning(u'%s: %s %s маршрут %s' % (OVERLIMIT_MESSAGE, self.date, self.time, self.route))
                    # Количество делаем равным лимиту
                    limit = int(driver.find_element_by_xpath('//span[@class="limit"]/strong').text)
                    if limit == 0:
                        logging.warning(u"Лимит равен 0")
                        return False
                    self.quantity = limit
                    self.statuses.append('TRUNCATED')
                    return self.reserve()
                else: # result -- текст ошибки
                    # возможные ошибки:
                    # - Превышено доступное количество (квота изменилась в процессе бронирования)
                    # - Не доступно для бронирования
                    logging.warning(u"Произошла ошибка при бронировании на %s %s маршрут %s. "
                        u"Текст ошибки: '%s'." % (self.date, self.time, self.route, result))
                    # действуем, исходя из того, что квота обнулилась и время больше недоступно
                    self.delete_time(timesAvailDict, timesAvail, self.time)
                    return self.reserve()
            #}}}
            # operQuota < operQuantity
            # Если объект заказа "размазывается" {{{4
            elif self.style == 'spread':
                # Иначе "Превышено доступное количество" (глюк сайта)
                if operQuantity > quota:
                    self.unreserve_time(self.time)
                # Бронируем всё доступное и продолжаем "размазываться"
                result = self.reserve_time(self.time, operQuota)
                if result == True:
                    # остаток продолжаем размазывать
                    self.quantity -= quota
                    # убираем время из списка доступных времён
                    self.delete_time(timesAvailDict, timesAvail, self.time)
                    # сортируем доступные времена в порядке удаления от первоначального времени заказа
                    minutes = str2minutes(self.time0)
                    if timesAvail != []:
                        timesAvail.sort(key=lambda t: abs(minutes-t[1]))
                        self.time = timesAvail[0][0]
                    else:
                        # Доступные времена могут закончиться на строке
                        #self.delete_time(timesAvailDict, timesAvail, self.time) выше.
                        # В этом случае должна выполниться ветка elif timesAvail == []. 
                        pass
                    return self.reserve()
                elif result == False:
                    logging.error(u"Не удалось забронировать остатки на время %s %s маршрут %s."
                        % (self.date, self.time, self.route))
                    return False
                elif result == OVERLIMIT_MESSAGE: # Превышение доступного кол-ва бронирования на данный месяц
                    logging.warning(u'%s: %s %s маршрут %s' % (OVERLIMIT_MESSAGE, self.date, self.time, self.route))
                    # Количество делаем равным лимиту
                    limit = int(driver.find_element_by_xpath('//span[@class="limit"]/strong').text)
                    if limit == 0:
                        logging.warning(u"Лимит равен 0")
                        return False
                    self.quantity = limit
                    return self.reserve()
                else: # result -- текст ошибки (возможные ошибки см. выше)
                    logging.warning(u"Произошла ошибка при бронировании остатков на %s %s маршрут %s. "
                        u"Текст ошибки: '%s'." % (self.date, self.time, self.route, result))
                    # действуем, исходя из того, что квота обнулилась и время больше недоступно
                    self.delete_time(timesAvailDict, timesAvail, self.time)
                    return self.reserve()
            #}}}
            # доступного количества хватает, чтобы остаться на этом маршруте и размазаться по нему {{{4
            elif operQuota>=operQuantity*(MIN_PERCENT/100.0) or operQuota>=MIN_QUOTA:
                logging.warning(u"На %s %s маршрут %s нельзя забронировать все %d билетов "
                    u"(доступно %d), бронируем доступное, остаток размазываем по другим временам"
                    % (self.date, self.time, self.route, self.quantity, quota))
                # Иначе "Превышено доступное количество" (глюк сайта)
                if operQuantity > quota:
                    self.unreserve_time(self.time)
                # Бронируем всё доступное
                result = self.reserve_time(self.time, operQuota)
                if result == True:
                    # размазываем остаток по другим временам этого маршрута
                    self.style = 'spread'
                    self.statuses.append('SPREAD')
                    self.quantity -= quota
                    # убираем время из списка доступных времён
                    self.delete_time(timesAvailDict, timesAvail, self.time)
                    # сортируем доступные времена в порядке удаления от первоначального времени заказа
                    minutes = str2minutes(self.time0)
                    timesAvail.sort(key=lambda t: abs(minutes-t[1]))
                    if timesAvail != []:
                        self.time = timesAvail[0][0]
                    else:
                        pass
                    return self.reserve()
                elif result == False:
                    logging.error(u"Не удалось забронировать остатки на время %s даты %s." % (self.time, self.date))
                    return False
                elif result == OVERLIMIT_MESSAGE: # Превышение доступного кол-ва бронирования на данный месяц
                    logging.warning(u'%s: %s %s маршрут %s' % (OVERLIMIT_MESSAGE, self.date, self.time, self.route))
                    # Количество делаем равным лимиту
                    limit = int(driver.find_element_by_xpath('//span[@class="limit"]/strong').text)
                    if limit == 0:
                        logging.warning(u"Лимит равен 0")
                        return False
                    self.quantity = limit
                    return self.reserve()
                else: # result -- текст ошибки (возможные ошибки см. выше)
                    logging.warning(u"Произошла ошибка при бронировании остатков на %s %s маршрут %s. "
                        u"Текст ошибки: '%s'." % (self.date, self.time, self.route, result))
                    # действуем, исходя из того, что квота обнулилась и время больше недоступно
                    self.delete_time(timesAvailDict, timesAvail, self.time)
                    return self.reserve()
            #}}}
            # доступное количество меньше MIN_PERCENT и MIN_QUOTA {{{4
            elif operQuota < operQuantity*(MIN_PERCENT/100.0) and operQuota < MIN_QUOTA and operQuota > 0:
                logging.warning(u"На %s %s маршрут %s доступно всего %d билетов (меньше %d%% и меньше %d)."
                    % (self.date, self.time, self.route, quota, MIN_PERCENT, MIN_QUOTA))

                # Ищем время, подходящее под условие резервирования, среди ближайших времён
                #в порядке удаления от первоначального времени заказа. Диапазон поиска времени ±1час
                #для маршрутов '2' и '3', и все времена в пределах маршрута для маршрута '1'.
                # Сортируем доступные времена в порядке удаления от первоначального времени заказа.
                logging.info(u'Поиск среди ближайших времён')
                minutes = str2minutes(self.time0)
                timesAvail.sort(key=lambda t: abs(minutes-t[1]))
                # ближайшее - это то же самое время, поэтому начинаем искать, начиная со второго.
                for time, mins in timesAvail[1:]:
                    q, r = timesAvailDict[time]
                    if q >= self.quantity*(MIN_PERCENT/100.0) or q >= MIN_QUOTA:
                        if self.route == 1:
                            logging.info(u'Ближайшее время: %s' % time)
                            self.time = time
                            self.statuses.append('OTHERTIME')
                            return self.reserve()
                        else:
                            logging.info(u'Ближайшее время: %s' % time)
                            if abs(minutes-mins) > 60:
                                logging.info(u'Ближайшее время слишком далеко')
                                self.routesPassed[self.route] = [time, abs(minutes-mins)]
                                break
                            else:
                                self.time = time
                                self.statuses.append('OTHERTIME')
                                return self.reserve()

                # Ближайшее время слишком далеко или его нет, переход на другой маршрут
                result = self.change_route()
                if result is False: # Переходим на маршрут, где доступное время наиболее близко к исходному
                    routes = self.routesPassed
                    route = min(routes, key=lambda k: routes[k][1])
                    time = routes[route][0]
                    self.route = route
                    self.time = time
                    self.statuses.append('ANYTHING')
                    return self.reserve()
                # result is True
                result = reservationDate.decide_route(self.route)
                if result is False: # пропуск заказа
                    logging.warning(u"Все маршруты на дату %s недоступны, пропуск заказа" % self.date)
                    return False
                self.route = result
                self.statuses.append('OTHERROUTE')
                return self.reserve()
            #}}}
            else: # if operQuota >= operQuantity {{{4
                logging.critical(u'Непредусмотренный случай. Дата %s' % self.date)
                # удаление времени
                self.delete_time(self.time)
                # попробуем ещё раз
                return self.reserve()
            #}}}
        #}}}
        # на это время уже всё забронировано или данного времени вообще нет {{{3
        else: # if self.time in timesAvailDict
            logging.warning(u"На %s %s маршрут %s нельзя ничего забронировать." %
                (self.date, self.time, self.route))

            # Ищем время, подходящее под условие резервирования, среди ближайших времён
            #в порядке удаления от первоначального времени заказа. Диапазон поиска времени SEARCH_RANGE
            #для маршрутов '2' и '3', и все времена в пределах маршрута для маршрута '1'.
            # Сортируем доступные времена в порядке удаления от первоначального времени заказа.
            logging.info(u'Поиск среди ближайших времён')
            minutes = str2minutes(self.time0)
            timesAvail.sort(key=lambda t: abs(minutes-t[1]))
            # Здесь начинаем искать, начиная с первого
            for time, mins in timesAvail:
                q, r = timesAvailDict[time]
                if q >= self.quantity*(MIN_PERCENT/100.0) or q >= MIN_QUOTA:
                    if self.route == 1:
                        logging.info(u'Ближайшее время: %s' % time)
                        self.time = time
                        self.statuses.append('OTHERTIME')
                        return self.reserve()
                    else:
                        logging.info(u'Ближайшее время: %s' % time)
                        if abs(minutes-mins) > SEARCH_RANGE:
                            logging.info(u'Ближайшее время слишком далеко')
                            self.routesPassed[self.route] = [time, abs(minutes-mins)]
                            break
                        else:
                            self.time = time
                            self.statuses.append('OTHERTIME')
                            return self.reserve()

            # Ближайшее время слишком далеко или его нет, переход на другой маршрут
            result = self.change_route()
            if result is False: # Переходим на маршрут, где доступное время наиболее близко к исходному
                routes = self.routesPassed
                route = min(routes, key=lambda k: routes[k][1])
                time = routes[route][0]
                self.route = route
                self.time = time
                self.statuses.append('ANYTHING')
                return self.reserve()
            result = reservationDate.decide_route(self.route)
            if result is False: # пропуск заказа
                logging.warning(u"Все маршруты на дату %s недоступны, пропуск заказа" % self.date)
                return False
            self.route = result
            self.statuses.append('OTHERROUTE')
            return self.reserve()
        #}}}
    #}}}
    def reserve_time(self, Time, quantity): #{{{2
        # td = driver.find_element_by_xpath('//td[contains(text(), "%s")]' % Time)
        td = WebDriverWait(driver, 10, 0.1).until(EC.element_to_be_clickable((By.XPATH, '//td[contains(text(), "%s")]' % Time)))
        return self.reserve_td_tickets(td, quantity)

    #}}}
    def reserve_td_tickets(self, td, quantity): #{{{2
        """Забронировать билеты на данное время"""
        try:
            td.click()
            # В ячейке появляется div#options_mod
            reserve = td.find_element_by_id("reserve")
            reserve.click()
            # Появляется окно div.window, вводим туда количество
            window = driver.find_element_by_css_selector('div.window')
            field = window.find_element_by_name('quantity')
            field.clear()
            field.send_keys(quantity)
            # Переключаемся на фрейм, тыкаем "Я не робот", уходим из фрейма
            driver.switch_to.frame(driver.find_element_by_tag_name('iframe'))
            checkbox = driver.find_element_by_id('recaptcha-anchor')
            checkbox.click()
            driver.switch_to.default_content()
            # Робот ждёт, пока человек введёт капчу
            WebDriverWait(driver, 100, 0.1).until(EC.visibility_of_element_located((By.CLASS_NAME, 'pls-contentLeft')))
            WebDriverWait(driver, 100, 0.1).until(EC.invisibility_of_element_located((By.CLASS_NAME, 'pls-contentLeft')))
            # Закрываем все диалоговые окна
            window.find_element_by_id("ok").click()
            t0 = T()
            window = WebDriverWait(driver, 10, 0.1).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div.ui-dialog')))
            t1 = T()
            title = window.find_element_by_css_selector('span.ui-dialog-title').text
            if u"ошибка" in title.lower():
                error_text = window.find_element_by_css_selector('div.ui-dialog-content').text
                result = error_text
            window.find_element_by_tag_name('button').click()
        except Exception, e: # NoSuchElementException и другие
            logging.error(u"Ошибка в функции reserve_td_tickets: %s" % e)
            return False
        else:
            return True

    #}}}
    def unreserve_time(self, Time): #{{{2
        td = driver.find_element_by_xpath('//td[contains(text(), "%s")]' % Time)
        return self.delete_td_reservation(td)

    #}}}
    def delete_td_reservation(self, td): #{{{2
        """Удалить бронь на данное время"""
        result = True
        try:
            script = """window.confirm = function() {return true;}"""
            driver.execute_script(script)
            td.click()
            # В ячейке появляется div#options_mod
            delete = td.find_element_by_id("delete")
            delete.click()
            # окно confirm должно пропускаться
            # здесь выполняется ajax ...
            window = WebDriverWait(driver, 10, 0.1).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div.ui-dialog')))
            title = window.find_element_by_css_selector('span.ui-dialog-title').text
            if u"ошибка" in title.lower():
                error_text = window.find_element_by_css_selector('div.ui-dialog-content').text
                result = error_text
            window.find_element_by_tag_name('button').click()
        except Exception, e: # NoSuchElementException и другие
            logging.error(u"Ошибка в функции delete_td_reservation: %s" % e)
            return False
        else:
            return result

    #}}}
    def delete_time(self, timesAvailDict, timesAvail, Time): #{{{2
        """Удалить время из timesAvail[Dict]. timesAvail[Dict] - это ссылки"""
        try:
            del timesAvailDict[Time]
            timesAvail.remove( (Time, str2minutes(Time)) )
        except KeyError:
            logging.error(u"Время %s не найдено в словаре 'timesAvailDict' маршрута %s" % (Time, self.route))
        except ValueError:
            logging.error(u"Время %s не найдено в списке 'timesAvail' маршрута %s" % (Time, self.route))

    #}}}
    def change_route(self): #{{{2
        """Выбрать маршрут, который ещё не обрабатывался заказом.
        Функция вызывается перед вызовом reservationDate.decide_route().
        """
        if len(self.routesPassed) == 3: # все маршруты перебраны
            return False
        # Подобно ReservationDate.decide_route() ветка 'Выбор альтернативного маршрута'
        if self.route == '1':
            if '2' not in self.routesPassed and '2' in reservationDate.routes:
                self.route = '2'
            elif '3' not in self.routesPassed and '3' in reservationDate.routes:
                self.route = '3'
            else: # все доступные маршруты перебраны
                return False
        elif self.route == '2':
            if '3' not in self.routesPassed and '3' in reservationDate.routes:
                self.route = '3'
            elif '1' not in self.routesPassed and '1' in reservationDate.routes:
                self.route = '1'
            else: # все доступные маршруты перебраны
                return False
        elif self.route == '3':
            if '2' not in self.routesPassed and '2' in reservationDate.routes:
                self.route = '2'
            elif '1' not in self.routesPassed and '1' in reservationDate.routes:
                self.route = '1'
            else: # все доступные маршруты перебраны
                return False
        else:
            logging.error("Unexpected self.route: %s" % self.route)
            return False

        return True
    #}}}
#}}}
def print_help(): #{{{1
    print u"%s - автоматическое бронирование билетов на сайте %s" % (SCRIPT, SITE)
    print u"Использование: %s ключи csv-файл(ы)..." % SCRIPT
    print u"Ключи:"
    print u"  -h, --help                  показать эту справку"
    print u"  -d, --delete                вместо бронирования удалять бронь"
    print u""
    print u"Журнал выводится в stdout"
    print u"Файл с пропусками -- log/failed.log."
    print u""
    print u"В файлах csv первая строка считается заголовком и пропускается. Далее:"
    print u"Первое поле -- дата в формате дд.мм.гггг"
    print u"Второе поле -- время в формате чч:мм"
    print u"Третье поле -- количество билетов, которое нужно бронировать (натуральное число)"
    print u"Четвёртое поле -- маршрут, на который предпочтительно бронировать билеты (1 или 2)"
    print u""
    print u"Скрипт исходит из того, что маршрутов два и обращается к ним по порядку,"
    print u"в котором они представлены в выборе \"Место проведения\"."

#}}}
class FormatError(Exception): #{{{1
    """Возникает, когда значение поля не соответствует регулярному выражению"""
    pass

#}}}
def str2minutes(S): #{{{1
    hh, mm = S.split(':')
    return int(hh)*60 + int(mm)

#}}}
def forwardedTimes_add(Date, Time): #{{{1
    if Date in forwardedTimes:
        forwardedTimes[Date].add(Time)
    else:
        forwardedTimes[Date] = set([Time])

#}}}
def wait_and_do(wait, do, trial=1): #{{{1
    """Функция дожидается события wait и делает действие do"""
    if trial > 3:
        return False
    try:
        E = WebDriverWait(driver, 10, 0.1).until(wait)
        if do == 'click':
            E.click()
        else:
            pass
    except WebDriverException:
        trial += 1
        return wait_and_do(wait, do, trial)
    else:
        return True

#}}}
def color(text): #{{{1
    """Обеспечить вывод текста в нужном цвете"""
    colors = {
        'FAILED': '31', # red
        'OK': '32', # green
        'OTHERTIME': '33', # yellow 
        'OTHERROUTE': '33',
        'ANYTHING': '35',
        'SPREAD': '36', # cyan
        'TRUNCATED': '36',
    }
    if text in colors:
        return '\033[1;%sm' % colors[text] + text + '\033[0m'
    else:
        return text

#}}}
def reserve(): #{{{1
    """Операция резервирования билетов"""
    while orderlist:
        order = orderlist.pop(0)
        try:
            result = order.reserve()
        except KeyboardInterrupt:
            logging.error(u'Возникло исключение KeyboardInterrupt!')
            report_all()
            sys.exit(1)
        except NoSuchWindowException:
            logging.error(u'Возникло исключение NoSuchWindowException!')
            report_all()
            sys.exit(1)
        except UnexpectedAlertPresentException:
            driver.switch_to.alert.accept()
            result = False
        except: # Здесь все необработанные исключения
            import traceback
            traceback.print_exc()
            logging.error(u'\033[1;31mНепредусмотренная ошибка\033[0m: пропуск заказа.')
            # При некоторых ошибках (oracle.php) DOM может слететь.
            # Нужно идти на главную и удалять экземпляр ReservationDate.
            global reservationDate
            reservationDate = None
            driver.get(SITE)
            result = False
        if result is True:
            del LINES[order.lineNum]
            save_file()
        if result is False:
            order.statuses.append('FAILED')
        statusText = ', '.join([color(s) for s in order.statuses])
        logging.info('%s;%s;%s;%s: [%s]'
            % (order.date, order.time0, order.quantity0, order.route0, statusText))
        if 'FAILED' in order.statuses:
            failedOrders.append(order)
        elif 'OTHERTIME' in order.statuses or 'OTHERROUTE' in order.statuses:
            forwardedOrders.append(order)

#}}}
def save_file(): #{{{1
    """Обновляем файл с заказами, оставляя только незабронированные."""
    with open(FILE, 'w') as F:
        F.writelines(LINES.values())

#}}}
def report(filename, orders): #{{{1
    """Ротация и создание отчётов об ошибках.
    filename - файл журнала
    orders - список объектов-заказов
    """
    # Ротация
    ends = ('.5', '.4', '.3', '.2', '.1', '')
    l = len(ends)
    # удалить файлы с последними номерами
    lastfile = filename + ends[0]
    if path.exists(lastfile):
        os.remove(lastfile)
    # увеличить номер остальных
    for i in xrange(1,l):
        oldfile = filename + ends[i]
        newfile = filename + ends[i-1]
        if path.exists(oldfile):
            os.rename(oldfile, newfile)

    # Создание файла
    F = open(filename, 'w')
    string = u"Дата;Время;Кол-во;Маршрут\n"
    for order in orders:
        string += '%s;%s;%s;%s\n' % (order.date, order.time0, order.quantity0, order.route0)
    F.write(string.encode("utf-8"))
    F.close()

#}}}
def report_all(): #{{{1
    """Создать отчёты по неудавшимся, перенаправленным и, если есть, пропущенным заказам."""
    if failedOrders:
        logging.warning(u"Есть неудавшиеся заказы. Посмотреть их можно в файле %s" % FAILEDFILE)
        report(FAILEDFILE, failedOrders)
    if forwardedOrders:
        logging.warning(u"Есть перенаправленные заказы. Посмотреть их можно в файле %s" % FORWARDEDFILE)
        report(FORWARDEDFILE, forwardedOrders)
    if orderlist:
        logging.warning(u"Есть невыполненные заказы. Посмотреть их можно в файле %s" % MISSEDFILE)
        report(MISSEDFILE, orderlist)

#}}}
# Начало исполнения. Анализ ключей и аргументов {{{1
try:
    longopts = ['help', 'delete']
    opts, args = getopt.getopt(ARGS, "dh12", longopts)
except getopt.GetoptError, err:
    print_help()
    sys.exit(2)

for o,a in opts:
    if o in ('-h', '--help'):
        print_help()
        sys.exit()
    elif o in ('-d', '--delete'):
        SCENARIO = 'delete'
    else:
        assert False, "unhandled option"

if len(args) == 0 or 'LOGIN' not in vars():
    print_help()
    sys.exit(2)

#}}}
# Анализ файла, создание объектов заказов {{{1
# Бэкапим файл, а из исходного будут удаляться строки по мере резервирования.
shutil.copyfile(FILE, FILE+'.orig')
# Даты должны обрабатываться в том порядке, в котором представлены в файле,
#но нужно избежать лишних перескоков по датам и маршрутам.
#Поэтому вводятся временные обёртывающие упорядоченные словари с ключами-датами
#и ключами-маршрутами, а потом заполняется список заказов.
orderDict =  OrderedDict() # od{date:od{route:[order,...],...},...}
with open(FILE) as F:
    reader = csv.reader(F, delimiter=';')
    for i,row in enumerate(reader):
        if i != 0: # первая строка -- заголовок
            if row == []: # пустая строка игнорируется
                continue
            if row[0].startswith('#'): # комментарий игнорируется
                continue

            try:
                d,t,q,r = row[:4]
                # проверка значений полей на правильность
                if not DATE_RE.match(d):
                    raise FormatError(u"Неправильный формат даты (%s не соответствует дд.мм.гггг) в строке %d файла %s." % (d, i+1, file))
                if not TIME_RE.match(t):
                    raise FormatError(u"Неправильный формат времени (%s не соответствует чч:мм) в строке %d файла %s." % (t, i+1, file))
                if not q.isdigit():
                    raise FormatError(u"Значение количества (%s) не является натуральным числом; строка %d файла %s." % (q, i+1, file))
                if not r.isdigit():
                    raise FormatError(u"Номер маршрута (%s) не является натуральным числом; строка %d файла %s." % (r, i+1, file))
                # проверки пройдены успешно, заполняем словарь
                q = int(q)
                if d in orderDict:
                    if r in orderDict[d]:
                        o = Order(i, d, r, t, q)
                        orderDict[d][r].append(o)
                    else:
                        o = Order(i, d, r, t, q)
                        orderDict[d][r] = [o]
                else:
                    o = Order(i, d, r, t, q)
                    orderDict[d] = OrderedDict([(r, [o])])
            except ValueError: # значений в строке меньше четырёх
                raise FormatError(u"Недостаточно данных в строке %d файла %s." % (i+1, file))

# Заполняем список orderlist
for Date, routes in orderDict.iteritems():
    for route, orders in routes.iteritems():
        for order in orders:
            orderlist.append(order)

# for order in orderlist:
    # print order.date, order.route0, order.time0, order.quantity
# sys.exit(0)

#}}}
# Настройка драйвера {{{1
chromedriver = os.path.expanduser('~/bin/chromedriver')
os.environ['webdriver.chrome.driver'] = chromedriver
CO = webdriver.ChromeOptions()
CO.binary_location = '/opt/google/chrome/google-chrome'

driver = webdriver.Chrome(chromedriver, chrome_options=CO)
# driver.implicitly_wait(10) # сколько максимально ждать появления элементов

#}}}
# Открытие браузера -- страница авторизации {{{1
driver.get(SITE)

driver.find_element_by_id("login").send_keys(LOGIN)
driver.find_element_by_id("password").send_keys(PASSWORD)
driver.find_element_by_id('submitBtn').click()

#}}}
# Начало отсчёта времени {{{1
T1 = T()
#}}}
# Бронирование {{{1
if SCENARIO == 'reserve':
    logging.info(u"Бронируем билеты")
    reserve()
    print "RESERVATION COMPLETED"
    report_all()

#}}}
# Удаление брони {{{1
# Удаляем бронь на все времена на все маршруты на даты, присутствующие в csv-файлах
elif SCENARIO == 'delete':
    logging.info(u'Удаляем бронь')
    datesDone = []

    for order in orderlist:
        d = order.date
        if d not in datesDone:
            dateObj = ReservationDate(d)
            dateObj.unreserve_all()
            datesDone.append(d)

else:
    print u"ОШИБКА. Неизвестный сценарий"
    sys.exit(2)

#}}}
# Конец работы {{{1
T2 = T()
print u"Время работы скрипта %.2f сек." % (T2-T1) 
#}}}
