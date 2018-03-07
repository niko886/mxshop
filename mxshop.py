#!/usr/bin/python
# -*- coding: utf8 -*-


import sys, os, re, shutil, hashlib, urllib, subprocess, json, imaplib, email, base64, requests, sqlite3, copy
import mechanize
import xlrd
import paramiko
from bs4 import BeautifulSoup, element


import conf
from collections import OrderedDict


class WebSearchNotFound(Exception):
    pass
class WebSearchMultipleFound(Exception):
    pass
class WebSearchSkuNotMatched(Exception):
    pass
class WebSearchPriceFound(Exception):
    pass
class WebSearchDuplicatedImage(Exception):
    pass
class WebSearchNoCategory(Exception):
    pass
class AdminFileNotFound(Exception):
    pass
class AdminNeedContinue(Exception):
    pass

if sys.version_info[0] == 3:
    from urllib import request as urllib2
else:
    import urllib2


import logging
import unittest

import tempfile
#from config import _CACHE_PATH

from optparse import OptionParser
from PIL import Image
from datetime import datetime
    
if os.path.exists('mxshop.log'):  # FIXME: proper log config
    os.remove('mxshop.log')

logging.basicConfig(format='%(levelname)s:%(funcName)s:%(lineno)d:%(message)s', level=logging.DEBUG, filename='mxshop.log')
log = logging.getLogger("mxshop")

console = logging.StreamHandler()
console.setLevel(logging.INFO)

log.addHandler(console) 

_CACHE_PATH = os.path.join(tempfile.gettempdir(), 'mxshop', 'cache')


def initCacheFolder(cachePath):

    try:
        os.makedirs(cachePath)
    except OSError as oe:    
        if '[Errno 17] File exists: ' not in str(oe):
            raise oe
    
    
class HttpPageCache():
    
    def __init__(self, site, **kw):
        
        dbFile = 'httpcache.db'
        
        newFile = kw.get('dbFile', 0) 
        if newFile:
            dbFile = newFile
        
        self._c = sqlite3.connect(os.path.join(_CACHE_PATH, dbFile))
        self._c.execute('CREATE TABLE IF NOT EXISTS "%s" (url TEXT PRIMARY KEY, data BLOB)' % site)
        self._site = site
        
        if kw.get('isClear', False):
            self._c.execute('DELETE FROM "%s"' % site)
        
    def put(self, url, data):
        
        self._c.execute('INSERT INTO "%s" VALUES (?, ?)' % self._site, (url, data,))
        self._c.commit()
        
        log.debug('stored to cache "%s" %d bytes [%s]', url, len(data), self._site)
        

    def get(self, url):
        
        cur = self._c.cursor() 
        cur.execute('SELECT data FROM "%s" WHERE url = ?' % self._site, (url,))
        
        res = cur.fetchone()
        
        if res:
            log.debug('fetched from cache "%s" %d bytes [%s]', url, len(res[0]), self._site)
            return res[0]
        else:
            return None
        
    def drop(self, url):
        cur = self._c.cursor()
        cur.execute('DELETE FROM "%s" WHERE url = ?' % self._site, (url,))
        self._c.commit()
        

class FileHlp():
    
    def __init__(self, path, access):
        
        
        self._path = ''
        
        if type(path) == list:
            for p in path:
                self._path = os.path.join(self._path, p)
        elif type(path) == str:            
            self._path = path
        else:
            raise ValueError('invalid type of parameter "path" provided,' \
                             'got "%s", but expected str or list' % (type(path)))
        
        log.debug("open file '%s'..." % self._path)
        
        self._file = open(self._path, access)
        
        
    
    def write(self, data):
        
        log.debug('writing to %s... (%d bytes)', self._path, len(data))
        
        self._file.write(data)
        self._file.close()
        
    def read(self):
        
        data = self._file.read()
        self._file.close()
        
        log.debug('%d bytes red from %s', len(data), self._path)
        
        return data
    

class Xml2003FileStub():
    
    _data = ''
    
    
    def escape(self, s, quote=False):
        
        s = s.replace("&", "&amp;") # Must be done first!
        s = s.replace("<", "&lt;")
        s = s.replace(">", "&gt;")
        if quote:
            s = s.replace('"', "&quot;")
            
        return s
    
    def __init__(self):
        
        pass
    
    def addrow(self, row):
        
        
        dummyRow = ' <Row>\n%s </Row>\n'
        dummyCell = '  <Cell><Data ss:Type="String">%s</Data></Cell>\n'
        
        
        
        cells = ''
        
        for r in row:
            cells += dummyCell % self.escape(str(r))
                    
        self._data += dummyRow % cells
            

    def write(self, filePath):
        
        stubFileData = FileHlp(['templates', 'minimal.xml'], 'r').read()        
        stubFileData = stubFileData.replace('<!--add-->', self._data)
                
        FileHlp(filePath, 'w').write(stubFileData)  
        
        log.info("xml file saved: %s", str(filePath))
      
    def getdata(self):
    
        stubFileData = FileHlp(['templates', 'minimal.xml'], 'r').read()        
        stubFileData = stubFileData.replace('<!--add-->', self._data)
        
        return str(stubFileData)

            

class Xml2003File():
    
    def __init__(self):
        
        stubFileData = FileHlp(['templates', 'minimal.xml'], 'r').read()
        
        self._soup = BeautifulSoup(stubFileData, 'lxml-xml') 

        log.info('---')        
        log.info(self._soup.Row)
        log.info('---')
        
        self._row = copy.copy(self._soup.Row)
        self._cell = copy.copy(self._row.contents[1])
        
        
    def setheaders(self, headers):


        rowOrig = self._soup.Row
        rowOrig.clear()
        
        log.debug('headers set: %s', ' | '.join(headers))

        for hh in headers:
            
            c = copy.copy(self._cell)
            
            c.Data.string = hh
            rowOrig.append(c)
            rowOrig.append('\n')
        
        
        
    def addrow(self, rowParam):
        
        row = copy.copy(self._row)
        cell = copy.copy(self._cell)
            
        row.clear()
        row.append('\n')
        
        numberIdx = [4, 5]
        
        intIdx = [11]
      
        for i in range(0, 45):
            
            try:
                rowEl = rowParam[i]
            except IndexError:
                rowEl = ''
         
            c = copy.copy(cell)
            
            
#             if i + 1 in numberIdx:
#                 if rowEl:              
#                     c.Data['ss:Type'] = "Number"
#                     
#             if i + 1 in intIdx:
#                 if rowEl:
#                     rowEl = str(int(float(rowEl))) # "1.0" case
                
            c.Data.string = rowEl
            row.append(c)
            row.append('\n')
      
                
        self._soup.Table.append(row)
        self._soup.Table.append('\n')

    def write(self, filePath):
                
        FileHlp(filePath, 'w').write(str(self._soup))  
        
        log.info("xml file saved: %s", str(filePath))
      
    def getdata(self):
        
        return str(self._soup)
        
        
reload(sys)
if os.name == 'nt':
    sys.setdefaultencoding('cp1251')
else:
    sys.setdefaultencoding('utf8')

_VERSION_ = '0.0.2'


class MXShopZhovtuha():
    
    # TODO: multiple responses to search requests, for example "sku in sku"

    _d = 'zhovtuha'  # just current dealer 
    
    _seoPrefix = 'zhv-'
        
    _walk = os.walk
    
    _webAdminId = 'zhovtuha-2'    
    _remoteImageDir = 'zhov'    
    
    _redirectByName = {}
    
    _categoryMap = {'Мотошлемы | Кросс': 'Шлема | Кроссовые',
                    'Мотошлемы | Dual': 'Шлема | Туристические',
                    'Мотошлемы | Модуляр': 'Шлема | Туристические',
                    'Мотошлемы | Интеграл': 'Шлема | Дорожные',
                    'Мотошлемы | Открытый': 'Шлема | Дорожные',
                    'Мотошлемы | Аксессуары': 'Шлема | Запчасти к шлему',
                    'МотоОчки | Кроссовые очки': 'Очки | Кроссовые',
                    'МотоОчки | Солнцезащитные': 'Очки | Солнцезащитные',
                    'МотоОчки | Аксессуары': 'Очки | Аксессуары к очкам',
                    'Мотоботинки | Кросс': 'Мотоботы | Кроссовые',
                    'Мотоботинки | Эндуро\АТV': 'Мотоботы | Туристические',
                    'Мотоботинки | Туристические': 'Мотоботы | Туристические',
                    'Мотоботинки | Городские': 'Мотоботы | Дорожные',
                    'Мотоботинки | Спортивные': 'Мотоботы | Дорожные',
                    'Мотоботинки | Аксессуары': 'Мотоботы | Запчасти к мотоботам',
                    'Защита | Защита тела': 'Защита | Груди и спины',
                    'Защита | Защита спины': 'Защита | Груди и спины',
                    'Защита | Наколенники': 'Защита | Коленей',
                    'Защита | Налокотники': 'Защита | Локтей',
                    'Защита | Мотоналокотники': 'Защита | Локтей',
                    'Защита | Защита шеи': 'Защита | Шеи',
                    'Защита | Защитные шорты': 'Защита | Шорты',
                    'Экипировка | Кросс | Джерси': 'Форма | Джерси',
                    'Экипировка | Кросс | Перчатки': 'Форма | Перчатки',
                    'Экипировка | Кросс | Штаны': 'Форма | Штаны',
                    'Экипировка | Подшлемники': 'Дорожная экипировка | Другое',
                    'Экипировка | Термоодежда': 'Форма | Термобелье',
                    'Экипировка | Дорожная | Перчатки': 'Дорожная экипировка | Перчатки',
                    'Экипировка | Дорожная | Курточки': 'Дорожная экипировка | Куртки',
                    'Экипировка | Дорожная | Штаны': 'Дорожная экипировка | Штаны',
                    'Экипировка | Дорожная | Дождевики': 'Дорожная экипировка | Другое',
                    'Расходники и Запчасти | Тормозная система | Тормозные колодки': 'Запчасти | Тормозные колодки',
                    'Расходники и Запчасти | Воздушные фильтры': 'Запчасти | Воздушные фильтры',
                    'Расходники и Запчасти | Тормозная система | Тормозные диски': 'Запчасти | Тормозные диски',
                    'Жидкости, смазки | Моторное масло': 'Химия | Моторное масло',
                    'Жидкости, смазки | Вилочное масло': 'Химия | Масло в подвеску',
                    'Жидкости, смазки | Для цепи': 'Химия | Для цепи',
                    'Жидкости, смазки | Для тормозов': 'Химия | Для тормозов',
                    'Жидкости, смазки | Для воздушного фильтра': 'Химия | Для воздушного фильтра',
                    'Жидкости, смазки | Для чистки шлемов и пластика': 'Химия | Другое',
                    'Жидкости, смазки | Для тросов': 'Химия | Другое',
                    'Жидкости, смазки | Бензиновые присадки': 'Химия | Другое',
                    'Жидкости, смазки | Охлаждающая жидкость': 'Химия | Другое',
                    'Жидкости, смазки | Трансмиссионное масло': 'Химия | Другое',
                    'Жидкости, смазки | Для шин': 'Химия | Другое',
                    'Аксессуары | Аудио': 'Аксессуары | Гарнитуры',
                    'Аксессуары | Сумки': 'Аксессуары | Сумки',
                    'Аксессуары | Разное' :'Аксессуары | Другое',
                    
                    # new style
                    
                    'МотоАксессуары | Разное' :'Аксессуары | Другое',
                    'МотоАксессуары | Сумки': 'Аксессуары | Сумки',
                    'МотоЭкипировка | Дорожная | МотоКурточки': 'Дорожная экипировка | Куртки',
                    'МотоЭкипировка | Дорожная | МотоДождевики': 'Дорожная экипировка | Другое', 
                    'МотоЭкипировка | Дорожная | Мотоперчатки': 'Дорожная экипировка | Перчатки', 
                    'МотоЭкипировка | Дорожная | Мотоштаны': 'Дорожная экипировка | Штаны', 
                    'МотоЭкипировка | Кросс | Джерси': 'Форма | Джерси', 
                    'МотоЭкипировка | Кросс | МотоПерчатки': 'Форма | Перчатки', 
                    'МотоЭкипировка | Кросс | МотоШтаны': 'Форма | Штаны', 
                    'МотоЭкипировка | МотоПодшлемники': 'Дорожная экипировка | Другое', 
                    'МотоЭкипировка | Термоодежда': 'Форма | Термобелье', 
                    'Мотоботинки | МотоАксессуары': 'Мотоботы | Запчасти к мотоботам', 
                    'Мотоботинки | Мотоботинки Городские': 'Мотоботы | Дорожные', 
                    'Мотоботинки | Мотоботинки Кроссовые': 'Мотоботы | Кроссовые', 
                    'Мотоботинки | Мотоботинки Спортивные': 'Мотоботы | Дорожные', 
                    'Мотоботинки | Мотоботинки Туристические': 'Мотоботы | Туристические', 
                    'Мотоботинки | Мотоботинки Эндуро\АТV': 'Мотоботы | Туристические', 
                    'Мотозащита | Мотозащита спины': 'Защита | Груди и спины', 
                    'Мотозащита | Мотозащита тела': 'Защита | Груди и спины', 
                    'Мотозащита | Мотозащита шеи': 'Защита | Шеи', 
                    'Мотозащита | Мотозащитные шорты': 'Защита | Шорты', 
                    'Мотозащита | Мотонаколенники': 'Защита | Коленей', 
                    'Мотозащита | Мотоналокотники': 'Защита | Локтей', 
                    'Мотозащита | Доп вставки мотозащиты': 'Запчасти | Другое',
                    'Мотошлемы | Мотошлем Dual': 'Шлема | Туристические', 
                    'Мотошлемы | Мотошлем Интеграл': 'Шлема | Дорожные', 
                    'Мотошлемы | Мотошлем Кросс': 'Шлема | Кроссовые', 
                    'Мотошлемы | Мотошлем Модуляр': 'Шлема | Туристические', 
                    'Мотошлемы | Мотошлем Открытый': 'Шлема | Дорожные', 
                    'Мотогарнитура': 'Аксессуары | Гарнитуры',
                    'МотоОчки | Мотоочки Кроссовые': 'Очки | Кроссовые',

                    # skipped

                    'Распродажа': '!!! Распродажа !!!',
                    'МотоОчки | Лыжные очки': '!!! МотоОчки | Лыжные очки !!!',                                         
                    
                    # Motostyle
                    
                    'Мотоэкипировка | Мотокуртки': 'Дорожная экипировка | Куртки',
                    'Расходники и Запчасти | Тормозные диски': 'Запчасти | Тормозные диски',
                    'Расходники и Запчасти | Выхлопные системы | Глушители для мотоциклов': 'Запчасти | Другое',
                    'Защита': 'Защита | Другое',
                    'Защита | Моточерепаха и защита спины': 'Защита | Груди и спины',
                    'Защита | Кроссовые маски и очки | Линзы для кроссовых масок': 'Очки | Аксессуары к очкам',
                    'Защита | Защита шеи / плеча': 'Защита | Шеи',
                    'Защита | Мотошлемы': 'Шлема | Дорожные',
                    'Защита | Мотонаколенники': 'Защита | Коленей',
                    'Защита | Моточерепахи': 'Защита | Груди и спины',
                    'Защита | Защитные  шорты': 'Защита | Шорты',
                    'Защита | Защита спины, груди (вставки)': 'Защита | Другое',
                    'Мотоэкипировка | Подшлемники': 'Дорожная экипировка | Другое',
                    'Мотоэкипировка | Дождевики': 'Дорожная экипировка | Другое',
                    'Мотоэкипировка | Термобелье': 'Форма | Термобелье',
                    'Тюнинг | Экстерьер | Замена пластика на кроссовые мотоциклы': 'Запчасти | Другое',
                    

                    }
    
    _categoryMapExt = [
        {'name': 'Защита | Мотоботы', 'mustHave': 'Обувь: Городская', 'target': 'Мотоботы | Дорожные'},
        {'name': 'Защита | Мотоботы', 'mustHave': 'Обувь: Спортивная', 'target': 'Мотоботы | Дорожные'},
        {'name': 'Защита | Мотоботы', 'mustHave': 'Обувь: Внедорожная - кросс, эндуро', 'target': 'Мотоботы | Кроссовые'},
        {'name': 'Защита | Мотоботы', 'mustHave': 'Обувь: Спортивная, Туристическая', 'target': 'Мотоботы | Дорожные'},
        {'name': 'Защита | Мотоботы', 'mustHave': 'Обувь: Туристическая', 'target': 'Мотоботы | Туристические'}, 
        {'name': 'Мотоэкипировка | Мотоперчатки', 'mustHave': 'Направление: Спорт', 'target': 'Дорожная экипировка | Перчатки'},
        {'name': 'Мотоэкипировка | Мотоперчатки', 'mustHave': 'Направление: Стрит/Туризм', 'target': 'Дорожная экипировка | Перчатки'},
        {'name': 'Мотоэкипировка | Мотоперчатки', 'mustHave': 'Направление: Кросс/ATV', 'target': 'Форма | Перчатки'},
        {'name': 'Мотоэкипировка | Мотоштаны', 'mustHave': 'Направление: Текстильные штаны', 'target': 'Дорожная экипировка | Штаны'},
        {'name': 'Мотоэкипировка | Мотоштаны', 'mustHave': 'Направление: Кросс/ATV, Текстильные штаны', 'target': 'Дорожная экипировка | Штаны'},
        ]
    
    _possibleBrands = ['putoline', 'ALIAS', 'ATLAS BRACE', 'Diadora', 'DRIFT', 'FORMA BOOTS', 'HJC', 'Icon', 'INTERPHONE', 'Gaerne',
                       'KNOX', 'MACNA', 'EBC', 'Micron', 'MOBIUS', 'OAKLEY', 'PUTOLINE OIL', 'REVIT', 'RS-TAICHI', 'SHOEI', 'Sidi', 'SPY+', 'SUOMY',
        ]
    
    _possibleBrands = [ps.upper() for ps in _possibleBrands]
    
    
    def assertDirectory(self, dirPath):
        
        try:
            os.makedirs(dirPath)
        except OSError as oe:    
            if '[Errno 17] File exists: ' not in str(oe):
                raise oe

    
    def __init__(self, **kw):
    
        self._pricesOrigDir = os.path.join('prices', 'orig', self._d) 
        self._pricesResutDir = os.path.join('prices', 'result', self._d)
        
        self.assertDirectory(self._pricesOrigDir)
        self.assertDirectory(self._pricesResutDir)
        
        self._priceFileRE = '^Остатки.*?-(\d+).(\d+).(\d+).xls$'  # example: Остатки-29.06.16.xls
        self._priceFileREidx = {'year': 3, 'month': 2, 'day': 1, 'hour': 0, 'minute': 0, 'second': 0}
        
                
        if kw.get('useTestingServer', False):
            log.info('[~] using testing server')
            self._webAdminLoginUrl = conf.MXSHOP_TEST_URL    
            self._remoteDirImageData = 'public_html/newtest/image/data'
            self._remoteUploadDir = 'public_html/newtest/admin/uploads'
        else:
            self._webAdminLoginUrl = conf.MXSHOP_URL      
            self._remoteDirImageData = '/var/www/html/image/data'
            self._remoteUploadDir = '/var/www/html/admin/uploads'
            self._remoteImageCacheDir = '/var/www/html/image/cache/data/kop'

                

    def SetWalkMock(self, mock):
        
        self._walk = mock
        
    def GetPricesOrigDir(self):
        
        return self._pricesOrigDir
    
    def _ReorderPriceListNamesByDate(self, priceNamesList):

        #
        # pre work checks 
        #

        if not priceNamesList:
            log.debug('there is not any prices not reordered')
            return priceNamesList

        if len(priceNamesList) < 2:
            log.debug('there is only one price, nothing can be reordered')
            return priceNamesList

        #
        # get maximum info
        #

        nameDictionary = {}

        for pricePath in priceNamesList:

            priceName = os.path.split(pricePath)[-1]

            s = re.search(self._priceFileRE, priceName)

            assert s

            year = s.group(self._priceFileREidx['year'])
            month = s.group(self._priceFileREidx['month'])
            day = s.group(self._priceFileREidx['day'])

            (hour, minute, second) = (0, 0, 0)

            if self._priceFileREidx['hour']:
                hour = s.group(self._priceFileREidx['hour'])

            if self._priceFileREidx['minute']:
                minute = s.group(self._priceFileREidx['minute'])

            if self._priceFileREidx['second']:
                second = s.group(self._priceFileREidx['second'])

            s = '%s-%s-%s %s:%s:%s' % (year, month, day, hour, minute, second)

            log.debug('datetime extracted from path: %s (%s)', s, pricePath)

            nameDictionary[s] = pricePath

        #
        # do actually reorder
        #

        resultList = []

        log.debug('sorted names: %s' % ' '.join(sorted(nameDictionary.keys())))

        for k in sorted(nameDictionary.keys()):

            resultList.append(nameDictionary[k])

        return resultList 
        
    def GetAllPriceNames(self):
        
        fileNames = []

        r = self._walk(self._pricesOrigDir)
        
        for root, dirs, files in r:

            for f in files:
                                
                log.debug('searching re: %s %s', self._priceFileRE, f)
                
                s = re.search(self._priceFileRE, f)

                if s:
                    path = os.path.join(root, s.group(0))
                    log.debug('found price file:\n %s', path)
                    fileNames.append(path)
                else:

                    if not fileNames:
                        raise NameError('check price file name RE, or directory... %s, %s' % (root, f))

        reordered = self._ReorderPriceListNamesByDate(fileNames)

        return reordered
    
    def ReadPrice(self, xlsPath):
        
        log.debug('reading price %s...' % xlsPath)

        rb = xlrd.open_workbook(xlsPath, formatting_info=True)

        sheet = rb.sheet_by_index(0)

        result = OrderedDict()

        invalidSkus = []
        invalidSkuPrice = []

        log.debug('----------------------------------------- table dump begin -----------------------------------------')

        # assert price format
        
        row6 = sheet.row_values(0)
        
        
        columnProduct = 0
        columnSku = 1
        columnPriceRetail = 2
        columnSaleOff = 0
        columnPriceDealer = 3
        columnBalance = 4
        
            
        if row6[3] == 'Акционная цена со скидкой':
            
            log.info('[!] price with sale off column detected')
            
            columnSaleOff = 3
            columnPriceDealer = 4
            columnBalance = 5
        
        assert(row6[columnProduct] == 'Товар')
        assert(row6[columnSku] == 'Артикул')
        assert(row6[columnPriceRetail] == 'Розничная цена')
        #assert(row6[columnPriceDealer] == 'Оптовая цена') FIXME: 
        assert('Оптовая цена' in row6[columnPriceDealer] or 'Розничная цена' in row6[columnPriceDealer])
        #assert(row6[columnBalance] == 'Склад (оптовый)') FIXME: 
                    
        saleOffCount = 0
        
        currentCategory = ''

        for rownum in range(3, sheet.nrows):
            
            row = sheet.row_values(rownum)

            product = row[columnProduct].strip()
            sku = str(row[columnSku]).strip()
            priceRetail = str(row[columnPriceRetail]).strip()
            saleOff = ''
            if columnSaleOff:
                saleOff = str(row[columnSaleOff]).strip()
            priceDealer = str(row[columnPriceDealer]).strip()
            balance = str(row[columnBalance]).strip()
                        
            #log.debug('%s %s %s %s %s %s', product, sku, priceRetail, saleOff, priceDealer, balance)
            
            invalidStr = ''
            
            priceIdx = str(rownum + 1)
                            
            if product and not sku and not priceRetail and not '(шт.)' in product:
                currentCategory = product
                log.debug('set current category: %s' % currentCategory)
                
                continue 

            if len(sku) <= 3:
                invalidStr = 'too poor sku "%s" (%s)' % (sku, priceIdx)
                invalidSkus.append(sku)

            if saleOff:
                try:
                    if float(saleOff) < float(priceRetail):
                        saleOffCount += 1
                    else:
                        saleOff = ''
                except ValueError as v:
                    invalidStr = 'can not handle saleOff price'
                    saleOff = ''

            if not product:
                invalidStr = 'empty product' 
                invalidSkus.append(sku)
                

            if not priceRetail:
                
                
                if priceDealer and 'GAERNE' in currentCategory.upper():
                    
                    newPrice = str(float(priceDealer) * 1.2)
                    priceRetail = newPrice
                    log.info("[!] price for Gaerne decuted from hardcode %s %s -> %s" %(
                        sku, priceDealer, priceRetail))
                
                else:
                    
                    invalidStr = 'empty priceRetail'
                    invalidSkuPrice.append(sku)

            if not priceDealer:
                invalidStr = 'empty priceDealer'
                invalidSkuPrice.append(sku)
            
            if invalidStr:
                log.warn('[!] %s; %s; %s; %s; %s; %s; %s; [invalid: %s]' % (priceIdx, sku, 
                          priceRetail, priceDealer, currentCategory, product, saleOff, invalidStr))
                continue

            assert(balance == '1.0')
            balance = '2'

            
            element = {'sku': sku, 
                       'priceRetail': priceRetail,
                       'priceDealer': priceDealer,
                       'priceSale': saleOff,
                       'balance': balance,
                       'categoryFromPrice': currentCategory,
                       'productFromPrice': product}

            if saleOff:
                log.debug('%s; %s; %s; %s; [SALE: %s]; %s; %s;' % (priceIdx, sku, priceRetail, priceDealer, saleOff, currentCategory, product))
            else:
                log.debug('%s; %s; %s; %s; %s; %s;' % (priceIdx, sku, priceRetail, priceDealer, currentCategory, product,))

            if not sku in result.keys():
            
                result[sku] = element
            else:
                raise ValueError('dupicate sku %s' % sku)

        log.debug('----------------------------------------- table dump end   -----------------------------------------')

        log.debug('%s price processed\n total rows: %d, skus: %d, invalid skus: %d, invalid price in sku %d' % (
            xlsPath, sheet.nrows, len(result.keys()), len(invalidSkus), len(invalidSkuPrice)))
        
        if saleOffCount:
            log.debug('[*] sale off count = %d', saleOffCount)

        log.info('price successfully read: %s', xlsPath)

        return result 

    def GetInfoMotostyleComUa(self, element):
        
        cache = HttpPageCache('motostyle-com-ua')
        
        cacheJson = HttpPageCache('motostyle-com-ua-json', dbFile='values-json.db')
        jsonResult = cacheJson.get(element['sku'])
        if jsonResult:
            return json.loads(jsonResult)

        
        baseUrl = 'http://motostyle.ua'
        
        searchUrl = baseUrl + '/index.php?route=product/search'
        
        s = requests.Session()
        
        searchReq = element['sku'].replace('/', ' ')
        
        params = {'keyword': searchReq}
        req = requests.Request('GET', searchUrl, params=params)
        prepped = s.prepare_request(req)        
        
        cached = cache.get(prepped.url)
        if not cached:
            resp = s.send(prepped)
            
            log.debug('%s -> %s', prepped.url, resp.status_code)
            
            cached = resp.text
            cache.put(prepped.url, cached)
        
        soup = BeautifulSoup(cached, 'lxml')
        
        productsFound = soup.findAll("div", { "class" : "product_info"})
        
        if not productsFound:
            s = 'not found information for "%s" on "motostyle.ua" (%s)' % (
                (searchReq, element['productFromPrice']))
            s += ' orig sku = "%s"' % (element['sku'])
            raise WebSearchNotFound(s)
        
        if len(productsFound) > 1:
            log.warn('[motostyle] multiple (%d) results found for %s', len(productsFound), searchReq)
        
        
        while len(productsFound):
            
            product = productsFound.pop()
        
            url = product.a['href'] 
            seoUrl = product.a['href'].replace('http://motostyle.ua/', '')
            assert(seoUrl)
                    
            cached = cache.get(url)
            if not cached:    
                resp = s.get(url)
                
                log.debug('%s -> %s', url, resp.status_code)
                
                cached = resp.text
                cache.put(url, cached)
            
            soup = BeautifulSoup(cached, 'lxml')
            
            webElement = {}
            webElement['category'] = ''
            webElement['sku'] = ''
            webElement['saleOffPercent'] = ''
            webElement['product'] = ''
            webElement['description'] = ''
            webElement['images'] = ''
            webElement['option'] = ''
            webElement['options'] = ''
            webElement['seoUrl'] = seoUrl
            webElement['brand'] = ''
    
            # get sku
            
            divs = soup.findAll("div", { "class" : "param_list item"})
            assert(len(divs) == 1)
            webSku = divs[0].em.text.strip()
            assert(webSku)
                            
            # find category
            
            categoryList = []
            
            divs = soup.findAll("div", { "class" : "breadcrumb "})
            assert(len(divs) == 1)
            
            for span in divs[0].findAll('span')[1:-1]:
                txt = span.text.strip()
                if txt:
                    categoryList.append(txt)
                        
            category = ' | '.join(categoryList)
            webElement['category'] = category
            
    
            # find saleOff (%)
    
            divs = soup.findAll("div", { "class" : "brand_discount"})
            if divs:
                assert(len(divs) == 1)
                saleOffText = divs[0].text.strip()
                saleOff = re.search('(\d+)%', saleOffText).group(1)
                assert(saleOff)
                webElement['saleOffPercent'] = saleOff
    
            # find name
    
            hs = soup.findAll("h1", { "itemprop" : "name"})
            assert(len(hs) == 1)
            name = hs[0].text.strip()
            webElement['product'] = name
            assert(name)
            
            origSku = element['sku'].replace('  ', ' ').upper()
            webSku = webSku.upper().replace('  ', ' ').upper()
            if webSku != origSku:
                # it can be that sku is right in name string:
#                 if origSku in name:
#                     webElement['sku'] = element['sku']
#                 else:
                    
                if webSku in origSku:
                    webElement['sku'] = webSku
                else:
                    if len(productsFound):
                        log.debug('[motostyle] page analyzed, sku not matched "%s" != "%s"' % (origSku, webSku))
                        continue
                    else:
                        raise WebSearchSkuNotMatched('[motostyle] sku not matched "%s" != "%s"' % (origSku, webSku)) 
            else:
                webElement['sku'] = webSku

            # assert category list
            
            if not categoryList:
                log.error('[motostyle] there is no category list for sku = %s' % element['sku'])
                assert(categoryList)
                continue

            
            # find description
            
            divs = soup.findAll("div", { "class" : "desc", "itemprop": "description"})
            assert(len(divs) == 1)
            o = ''
            for c in divs[0].contents:
                s = str(c).strip().replace('\n', '').replace('\r', '')
                o += s
            description = o
            webElement['description'] = description
            #assert(description)
            
            # find images
            
            imgList = []        
            
            divs = soup.findAll("div", { "class" : "image_block"})
            assert(len(divs) == 1)
            for a in divs[0].findAll('a'):
                if a.get('href', 0):
                    imgList.append(a['href'])
                    
            assert(imgList)
                 
            webElement['images'] = imgList         
             
            # find current option
            
            al = soup.findAll("a", {"class": "active"})
            if al:
                assert(len(al) == 1)
                option = al[0].text.strip()
                webElement['option'] = option
                assert(option)
                
            # find brand
            
            #<div class="param_list item">
            divs = soup.findAll("div", {"class": "param_list item"})
        
            assert(len(divs) == 1)
            
            s = re.search('<li>Brand: <em>(.*?) -', str(divs[0]), re.MULTILINE)
            webElement['brand'] = s.group(1).strip()
            
            
            
            tt = divs[0].ul.text
            
            extInfo = {}
            
            for line in tt.split('\n'):
                
                line = line.strip()
                
                if line:

                    
                    left, right = line.split(':')
                    left = left.strip()
                    right = right.strip()
                    
                    extInfo[left] = right
                    
            webElement['extInfo'] = extInfo
            webElement['extInfoTxt'] = str(divs[0].ul.text)
    
            cacheJson.put(element['sku'], json.dumps(webElement))
            return webElement


    def GetInfoMotocrazytownComUa(self, element):
        
        cache = HttpPageCache('motocrazytown-com-ua')
        
        cacheJson = HttpPageCache('motocrazytown-com-ua-json', dbFile='values-json.db')
        jsonResult = cacheJson.get(element['sku'])
        if jsonResult:
            return json.loads(jsonResult)

         
        baseUrl = conf.MOTOCRAZY_HOST
        searchUrl = baseUrl + '/search'
        
        session = requests.Session()
        
        params = {'q': element['sku']}
        req = requests.Request('GET', searchUrl, params=params)
        prepped = session.prepare_request(req)
        
        cached = cache.get(prepped.url)         
        if not cached:
            resp = session.send(prepped)
            
            log.debug('%s -> %s', prepped.url, resp.status_code)
            
            cached = resp.text
            cache.put(prepped.url, cached)
        
        soup = BeautifulSoup(cached, 'lxml')
    
        productsFound = soup.findAll("li", { "class" : "product-item"})
        
        if not productsFound:
            raise WebSearchNotFound('can not find any results (%s)' % element['sku'])

        if len(productsFound) > 1:
            msg = '[motocrazytown] more than one result found (%d items) for "%s"' % (len(productsFound), 
                                                                      element['sku'])
            log.warn(msg)
        
        while len(productsFound):
            
            product = productsFound.pop(0)        
                
            seoUrl = product.a['href'].replace('product/view/', '')
            assert(seoUrl)
    
            url = baseUrl + '/' + product.a['href']
    
            cached = cache.get(url)         
            if not cached:    
                resp = session.get(url)
                
                log.debug('%s -> %s', url, resp.status_code)
                
                cached = resp.text
                cache.put(url, cached)
            
            soup = BeautifulSoup(cached, 'lxml')
            
            webElement = {}
            webElement['category'] = ''
            webElement['sku'] = ''
            webElement['saleOffPercent'] = ''
            webElement['product'] = ''
            webElement['description'] = ''
            webElement['images'] = ''
            webElement['option'] = ''
            webElement['seoUrl'] = seoUrl
            webElement['options'] = '' # optional
            webElement['brand'] = ''
            
            # find category
            
            uls = soup.findAll("ul", { "class" : "breadcrumbs"})
            categoryList = []
            for li in uls[0].findAll('li')[:-1]:
                txt = li.text.strip()
                if txt:
                    categoryList.append(txt)
                    
            category = ' | '.join(categoryList)
            webElement['category'] = category
            assert(category)
            
            # find sku
            
            ids = soup.findAll("span", { "itemprop" : "identifier"})
            assert(len(ids) == 1)
            webSku = ids[0].text.strip()
            assert(webSku)
            
            webElement['sku'] = webSku
            
            # find saleOff (%)
    
            ids = soup.findAll("div", { "class" : "discount-info"})
            if ids:
                assert(len(ids) == 1)
                saleOffText = ids[0].text.strip()
                saleOff = re.search('(\d+)%', saleOffText).group(1)    
                webElement['saleOffPercent'] = saleOff
    
            # find name
    
            divs = soup.findAll("div", { "class" : "center"})
            assert(len(divs) == 1)
            name = divs[0].div.contents[0].strip()
            assert(name)
            webElement['product'] = name
            
            # find description
    
            divs = soup.findAll("div", { "class" : "text-block", "itemprop": "description"})
            assert(len(divs) == 1)
            o = ''
            for c in divs[0].contents:
                s = str(c).strip().replace('\n', '').replace('\r', '')
                o += s

            description = o
            #assert(description)
            webElement['description'] = description
            
            # find images
            
            imgList = []
            
            divs = soup.findAll("div", { "class" : "left-column"})
            assert(len(divs) == 1)
            imgList.append(baseUrl + divs[0].div.a['href'])
                    
            uls = soup.findAll("ul", { "class" : "product-images"})
            for ul in uls:
                imgList.append(baseUrl + ul.li.div.a['href'])
            assert(imgList)
            webElement['images'] = imgList
            
             
            # find options
    
            options = {}
            selects = soup.findAll("select", {"id": "product_variant_select"})
            if selects:
                for option in selects[0].findAll("option"):
                    options[option['data-code']] = option.text.strip()
    
            
            if options:
                currentOpt = options.get(element['sku'], 0)
                
                if not currentOpt:
                    
                    if len(productsFound):
                        log.debug('[crazytown] page analyzed, sku not matched "%s" != "%s"' % 
                              (webSku, element['sku']))
                        continue
                    else:
                        raise WebSearchSkuNotMatched('[crazytown] sku not matched "%s" != "%s"' % (element['sku'], str(options)))
            
                webElement['option'] = currentOpt
                webElement['options'] = options
                  
                  
            cacheJson.put(element['sku'], json.dumps(webElement))
            
            return webElement
        
    def GrabWebData(self, priceData):
        
        assert(priceData)
        
        webData = {}
        
        i = 0
        notFoundCount = 0
        notMatchedCount = 0
        notFoundCountOrig = 0
        noDescriptionOrig = 0
        noDescription = 0
        m = len(priceData.keys())
        for sku in priceData.keys():
             
            i += 1             
            
            sys.stdout.write('%d/%d\r' % (i, m))
            sys.stdout.flush()
            
            notFound = False
             
            try:
                
                element = self.GetInfoMotocrazytownComUa(priceData[sku])
                
            except (WebSearchNotFound, WebSearchSkuNotMatched) as ex:
                
                notFound = True
                notFoundCountOrig += 1
                
                log.debug('[!] not found information for %s on "motocrazy...ua" %s' % (sku, ex))
                
                try:
                    element = self.GetInfoMotostyleComUa(priceData[sku])
                    
                except WebSearchNotFound as notFound:
                    
                    notFoundCount += 1
                    
                    log.warn('[!] %s (%s)' % (str(notFound), priceData[sku]['productFromPrice']))
                    continue                
                
                except WebSearchSkuNotMatched as notMatched:
                    
                    notMatchedCount += 1
                    
                    log.warn('[!] %s (%s)' % (str(notMatched), priceData[sku]['productFromPrice']))
                    continue

            assert(element)
            assert(not sku in webData.keys())
            
            if not element['description'] and not notFound:
                
                noDescriptionOrig += 1
                
                log.debug('[~] no description found for "%s" on motocrazytown', sku)
                
                try:
                    
                    element2 = self.GetInfoMotostyleComUa(priceData[sku])
                    
                    if element2['brand'] and not element['brand']:
                        element['brand'] = element2['brand']
                        

                    if element2['description']:
            
                        element['description'] = element2['description']
                        
                    else:
                        log.warn('[~] no description found for "%s" in all sites', sku)
                        
                        noDescription += 1
                    
                except (WebSearchNotFound, WebSearchSkuNotMatched) as ex:
                    
                    noDescription += 1
            
            assert(element['images'])
            
            webData[sku] = element
            
        log.info('web grab for zhovtuha completed')
        log.info('searched for %d items, not found orig - %d, not found total %d, ' \
                 'not matched %d, no description orig %d, no description total %s', 
                 i, notFoundCountOrig, notFoundCount, notMatchedCount, noDescriptionOrig,
                 noDescription)
        
        return webData


    def Transliterate(self, text):

        symMap = (u"абвгдеёжзийклмнопрстуфхцчшщъыьэюя",
                  u"abvgdeezzijklmnoprstufhccss'y'eyy")

        dict2 = {ord(a):ord(b) for a, b in zip(*symMap)}

        text = text.lower()
        out = ''

        for t in text:
            if ord(t) in dict2.keys():
                out += chr(dict2[ord(t)])    
            else:
                out += t

        return out
    
    def WebAdminLogin(self):
        
        br = mechanize.Browser()
         
        br.addheaders = [("User-agent", "Mozilla/5.0 (X11; Linux x86_64; rv:45.0)")]
    
        br.set_handle_referer(True)
        br.set_handle_robots(False)
           
        log.info('logging to admin...')
           
        url1 = self._webAdminLoginUrl 
           
        r = br.open(url1)
        
        FileHlp([_CACHE_PATH, 'web-admin-login-1.html'], 'w').write(BeautifulSoup(r.read(), 'lxml').prettify())
        
        br.select_form(nr=0)
        br['username'] = conf.WEB_ADMIN_LOGIN
        br['password'] = conf.WEB_ADMIN_PASS
        
        r = br.submit()
        
        return br, r
    
    
    def WebAdminFixCategories(self):
        
        br, r = self.WebAdminLogin()
        
        soup = BeautifulSoup(r.read(), 'lxml')
        
        FileHlp([_CACHE_PATH, 'web-admin-fix-cat-2.html'], 'w').write(soup.prettify())
        
        nextLink = ''
        for a in soup.find_all('a'):
            href = a.get('href', '')
            if 'catalog/category' in href:
                nextLink = href
        
        if not nextLink:
            raise ValueError('can not find dealers link')
        
        r = br.open(nextLink)
        soup = BeautifulSoup(r.read(), 'lxml')
        FileHlp([_CACHE_PATH, 'web-admin-fix-cat-3.html'], 'w').write(soup.prettify())
    
        tbody = soup.tbody
        nextLink = ''
        for row in tbody.find_all('tr'):
            
            tds = [td for td in row.find_all('td')]
            
            name = tds[1]
            edit = tds[3]
            
            seoName = self.Transliterate(name.text.replace(' > ', '-').replace(' ', ''))

            nextLink = edit.a['href']

            print("%s %s" % (seoName, nextLink))
            
            r = br.open(nextLink)
            soup = BeautifulSoup(r.read(), 'lxml')
            FileHlp([_CACHE_PATH, 'web-admin-fix-cat-4-%s.html' % seoName], 'w').write(soup.prettify())
            
            
            br.select_form(nr=0)            
            br['keyword'] = seoName
            r = br.submit()
            
            log.info('category "%s" seo url fixed', seoName)
    
    def WebAdminGetRemoteXmlName(self):
        
        br, r = self.WebAdminLogin()
        
        soup = BeautifulSoup(r.read(), 'lxml')
        
        FileHlp([_CACHE_PATH, 'web-admin-login-2.html'], 'w').write(soup.prettify())
        
        nextLink = ''
        for a in soup.find_all('a'):
            href = a.get('href', '')
            if 'catalog/suppler' in href:
                nextLink = href
        
        if not nextLink:
            raise ValueError('can not find dealers link')
        
        r = br.open(nextLink)
        soup = BeautifulSoup(r.read(), 'lxml')
        FileHlp([_CACHE_PATH, 'web-admin-login-3.html'], 'w').write(soup.prettify())
    
        tbody = soup.tbody
        nextLink = ''
        
        for row in tbody.find_all('tr'):
            
            tds = [td for td in row.find_all('td')]
            
            name = tds[1]
            run = tds[3]
            #edit = tds[4]
            
            if self._webAdminId in name.text:
                
                s = re.search('\d+.xml', run.text.strip())
                
                if s:
                    xmlName = s.group(0)
                
                    log.info('found dealer xml name %s -> %s', name.text.strip(), xmlName)
                    return xmlName
            
        raise ValueError('can not find "%s" id in dealers, please create it manually first' % self._webAdminId) 
    
    def WebAdminRunPrice(self):
        
        br, r = self.WebAdminLogin()
        
        soup = BeautifulSoup(r.read(), 'lxml')
        
        FileHlp([_CACHE_PATH, 'web-admin-login-2.html'], 'w').write(soup.prettify())
        
        nextLink = ''
        for a in soup.find_all('a'):
            href = a.get('href', '')
            if 'catalog/suppler' in href:
                nextLink = href
        
        if not nextLink:
            raise ValueError('can not find dealers link')
        
        r = br.open(nextLink)
        soup = BeautifulSoup(r.read(), 'lxml')
        FileHlp([_CACHE_PATH, 'web-admin-login-3.html'], 'w').write(soup.prettify())
    
        tbody = soup.tbody
        nextLink = ''
        
        for row in tbody.find_all('tr'):
            
            tds = [td for td in row.find_all('td')]
            
            name = tds[1]
            run = tds[3]
            #edit = tds[4]
            
            if self._webAdminId in name.text:
                nextLink = run.a['href']
                log.info('found dealer entry: %s -> %s', name.text.strip(), nextLink)
            
        if not nextLink:
            raise ValueError('can not find "%s" id in dealers, please create it manually first' % self._webAdminId)

        log.info('running price for [%s]...', self._webAdminId)

        try:
            r = br.open(nextLink)
        except mechanize.HTTPError as he:
            
            log.info('[~] http error %s', str(he))
            
            if he.code == 404:
                raise AdminNeedContinue('please make request once more time')

            if he.code == 503:
                raise AdminNeedContinue('please make request once more time')

        dd = r.read()
        
        log.info('[+] read %d bytes', len(dd))
        
        soup = BeautifulSoup(dd, 'lxml')
        FileHlp([_CACHE_PATH, 'web-admin-login-4.html'], 'w').write(soup.prettify())
        
        warns = soup.find_all('div', class_='warning')
        for w in warns:
            log.error("admin error: %s", w.text)
            raise RuntimeError(("admin error: %s" % w.text))
        
        oks = soup.find_all('div', class_='success')
        for w in oks:
            log.info("price done: %s", w.text)
            return
        
        raise RuntimeError("unknown page result, please check last .html files")


    def DownloadCurrentPriceFromWeb(self):

        from email.header import decode_header
        
        log.info('logging to google (imap)...')
        
        imapSsl = imaplib.IMAP4_SSL('imap.gmail.com', 993)
        imapSsl.login(conf.GMAIL_LOGIN, conf.GMAIL_PASS)
        imapSsl.select()
        
        typ, data = imapSsl.search(None, '(TO "zhovtuha")') #  % conf.GMAIL_SEARCH_FROM
        
        # fetch last message
        typ, dd = imapSsl.fetch(data[0].split()[-1], '(RFC822)') 
        
        msg = email.message_from_string(dd[0][1])
        
        fileName = ''
        fileData = ''
        messageText = ''
        
        for part in msg.walk():
            
            if part.get_content_type() == 'text/plain':
                
                log.debug('text/plain:')
                
                cont = part.get_content_charset()
                
                if cont:
                    
                    messageText = str(part.get_payload(decode=True).decode(cont))
                    messageText = messageText.strip()
                    
                else:
                    
                    messageText = str(part.get_payload())
                    messageText = messageText.strip()
                
                log.debug(messageText) 

                 
            if part.get_content_type() == 'application/vnd.ms-excel':
                
                fileName = part.get_filename()
                if decode_header(fileName)[0][1] is not None:
                    fileName = decode_header(fileName)[0][0].decode(decode_header(fileName)[0][1])
                
                fileName = str(fileName)            
                log.debug(fileName)
                
                
                fileData = part.get_payload(decode=True)
                    
                    
                            
        imapSsl.close()
        imapSsl.logout()
    
        assert(fileName)
        assert(fileData)   
        assert(messageText)

        # get currency rate
        
        currencyRate = 0
        s = re.search('Курс.*?(\d+[,\.]*\d+)', messageText, re.MULTILINE)
        if s:
            currencyRate = float(s.group(1).replace(',', '.'))
            assert(currencyRate > 20 and currencyRate < 35)
            log.info('currency = %f', currencyRate)
        else:
            raise ValueError('currency not found, text: """%s"""' % messageText)             
            
        
        FileHlp([_CACHE_PATH, fileName], 'w').write(fileData)
        
        return fileData, str(fileName), currencyRate
        

  
    def DownloadCurrentPriceFromWeb0(self):
    
        br = mechanize.Browser()
         
        br.addheaders = [("User-agent", "Mozilla/5.0 (X11; Linux x86_64; rv:45.0)")]
    
        br.set_handle_referer(True)
        br.set_handle_robots(False)
           
        log.info('logging to google...')
           
        url1 = 'https://accounts.google.com/ServiceLogin?service=mail&passive=true&rm=false&continue=https://mail.google.com/mail/&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1' 
           
        r = br.open(url1)
           
        log.debug('google first response code: %d', r.code)
           
        FileHlp([_CACHE_PATH, 'google1.html'], 'w').write(BeautifulSoup(r.read(), 'lxml').prettify())
           
        br.select_form(nr=0)
        br['Email'] = conf.GMAIL_LOGIN
        r = br.submit()
           
        FileHlp([_CACHE_PATH, 'google2.html'], 'w').write(BeautifulSoup(r.read(), 'lxml').prettify())
           
        log.debug('google second response code: %d', r.code)
           
           
        br.select_form(nr=0)
        br['Passwd'] = conf.GMAIL_PASS
        r = br.submit()
           
        data = r.get_data()
   
        FileHlp([_CACHE_PATH, 'google3.html'], 'w').write(data)
           
        log.debug('google third response code: %d', r.code)
   
        soup = BeautifulSoup(r.get_data(), "lxml")
        pretty = soup.prettify()
           
        FileHlp([_CACHE_PATH, 'google3-pretty.html'], 'w').write(pretty)
           
        pretty = re.sub('<noscript>.*?</body>', '</body>', pretty, flags=re.DOTALL + re.MULTILINE)
                   
        FileHlp([_CACHE_PATH, 'google3-pretty-cut.html'], 'w').write(pretty)
           
        r.set_data(pretty)
        br.set_response(r)
   
        # switch to basic html view
           
        br.select_form(nr=0)
        r = br.submit()
                   
        FileHlp([_CACHE_PATH, 'google4.html'], 'w').write(BeautifulSoup(r.read(), 'lxml').prettify())
        log.debug('google fourth response code: %d', r.code)
           
        # search Джон
   
        br.select_form(nr=0)
        br['q'] = 'Жовтуха'
        r = br.submit()
           
        soup = BeautifulSoup(r.read(), 'lxml')
        FileHlp([_CACHE_PATH, 'google-жовтуха.html'], 'w').write(soup.prettify())
        log.debug('google fifth response code: %d', r.code)

        #soup = BeautifulSoup(FileHlp([_CACHE_PATH, 'google-Джон.html'], 'r').read(), 'lxml')
 
        messages = []
        
        for tr in soup.find_all('tr'):
            
            if 'bgcolor' in tr.attrs and tr['bgcolor'] == "#E8EEF7":
                
                td = tr.find_all('td')
                
                message = {}
                message['from'] = td[1].text.strip()
                message['link'] = td[2].a['href']
                message['short'] = td[2].a.span.text.replace('\n', '').replace('  ', '')
                
                messages.append(message)
                
                log.debug('message found: %s %s', message['from'], message['short'])

        # open first message
        
        log.debug('total message found: %d', len(messages))
        log.debug('open message: %s %s %s...', messages[0]['from'],  messages[0]['link'], messages[0]['short'][:15])
        
        r = br.open(messages[0]['link'])
        
        log.debug('google sixth response code: %d', r.code)
        
        soup = BeautifulSoup(r.read(), 'lxml')
        FileHlp([_CACHE_PATH, 'google-6.html'], 'w').write(soup.prettify())
        
        # get currency rate
        
        currencyRate = 0
        s = re.search('Курс (\d+[,\.]*\d+)', str(soup.text), re.MULTILINE)
        if s:
            currencyRate = float(s.group(1).replace(',', '.'))
            assert(currencyRate > 20 and currencyRate < 35)
            log.info('currency = %f', currencyRate)
        else:
            raise ValueError('currency not found, text: """%s"""' % str(soup.text))        
        
        nextUrl = ''
        fileName = ''
        for table in soup.find_all('table', class_='att'):
            
            if '.xls' in table.text:
                
                log.debug('table found')
                                
                for a in table.find_all('a'):

                    log.debug('a found %s' % a)
                
                    if a.text.strip() == 'Scan and download' or \
                       a.text.strip() == 'Download':
                        
                        nextUrl = a['href']
                        
                        s = re.search('^.*?\.xls', table.text, re.MULTILINE)
                        if s:
                            fileName = str(s.group(0))
                        else:
                            raise AttributeError('RE not matched for filename')
                        
        if not nextUrl:
            raise AttributeError('no xls attachment found')
        
        log.info('requesting: %s', nextUrl)
         
        r = br.open(nextUrl)        
        log.debug('google seventh response code: %d', r.code)
        data = r.read()
        
        FileHlp([_CACHE_PATH, fileName], 'w').write(data)
        
        return data, fileName, currencyRate
        
    def RedirectCategoryByName(self, name):
        
        if not self._redirectByName:
            return None

        newCategory = None

        for k in self._redirectByName.keys():
            
            if k in name: 
                newCategory = self._redirectByName[k]
                break
                
            if k.upper() in name.upper():
                newCategory = self._redirectByName[k]
                break


        
        return newCategory

    def ConvertTo97Xls(self, fromPath, toPath):
        
        assert(os.name == 'posix')
        
        tempDir = tempfile.gettempdir()
                    
        params = ['soffice', '--headless', 
                              '--convert-to', 'xls:MS Excel 97', '--outdir', tempDir,
                              fromPath]
        log.debug("running %s", ' '.join(params))
                
        sub = subprocess.Popen(params)        
        code = sub.wait()
        assert(code == 0)
        
        name = os.path.split(fromPath)[-1]
        tmpFile = os.path.join(tempDir, name)
        
        code = subprocess.call(['mv', tmpFile, toPath])
        assert(code == 0)
        
        log.info("price converted %s -> %s" % (fromPath, toPath))
        
        return toPath
                
    def CreateXmlFile(self, priceData, webData, xmlFilePath, currencyRate):
        
        
        assert(priceData)
        assert(webData)
        
        failedCategories = []

        # create base
        xml = Xml2003FileStub()
        
        # price format
        xml.addrow(['orig sku', 
                        'category', 
                        'product',
                        'priceSaleUah (%.2f)' % currencyRate,
                        'priceUah', 
                        'dealer usd/uah/salary', 
                        'option', 
                        'options', 
                        'seoUrl + dealerMark', 
                        'description',
                        'balance',
                        'brand'])
        
        # add to elements          
        
        sortedKeys = priceData.keys()
        sortedKeys.sort()
              
        for sku in sortedKeys:  
            
            if not sku in webData.keys():
                continue
                                        
            # priceData
            # ['sku', 'priceSale', 'priceRetail', 'categoryFromPrice', 'priceDealer', 
            # 'productFromPrice', 'balance']
            
            # webData
            # [u'category', u'sku', u'product', u'description', u'saleOffPercent', 
            # u'seoUrl', u'images', u'options', u'option']            
                
            rr = ['' for i in range(0, 45)]
            
            options = ''
            
            if webData[sku]['options']:
                options = ', '.join(webData[sku]['options'].values())

            rr[0] = priceData[sku]['sku']
            
            log.debug('mapping: "%s" [%s]', webData[sku]['category'], priceData[sku]['sku'])
            
            
            cat = '!!! fail !!!'
            
            try:
                cat = self.MapCategory(webData[sku]['category'])
            except KeyError as ke:
                                
                try:
                    cat = self.MapCategoryExt(webData[sku]['category'], str(webData[sku]['extInfoTxt']))
                except KeyError as ke:
                    
                    if not webData[sku]['category'] in failedCategories:
                        failedCategories.append(webData[sku]['category'])
                        
                    log.error('"%s" %s' % (webData[sku]['category'], ke))
                    
            redirectedCat = self.RedirectCategoryByName(webData[sku]['product'])
            if redirectedCat:
                log.debug("redirect product by name: '%s': '%s' -> '%s'",
                         webData[sku]['product'], cat, redirectedCat)
                cat = redirectedCat
                                
            brand = webData[sku]['brand']
            cfp = priceData[sku].get('categoryFromPrice', '')
            
            if not brand:
                if cfp and cfp.strip().upper() in self._possibleBrands:
                    brand = cfp

            if not brand:
                log.debug('still no brand for [%s] %s (categoryFromPrice = %s)', 
                         webData[sku]['sku'], webData[sku]['product'], cfp)

            priceUah = '%.2f' % (float(priceData[sku]['priceRetail']) * currencyRate)
            if priceData[sku]['priceSale']:
                priceUahSale = '%.2f' % (float(priceData[sku]['priceSale']) * currencyRate)
            else:
                priceUahSale = ''
                
            if priceUahSale:
                lowerPrice = float(priceData[sku]['priceSale']) * currencyRate
            else:
                lowerPrice = float(priceData[sku]['priceRetail']) * currencyRate
            
            lowerPrice = float(priceData[sku]['priceSale']) * currencyRate if priceUahSale else float(priceData[sku]['priceRetail']) * currencyRate
                
            pricesOther = "$%s / ₴%.2f / ₴%.2f" % (priceData[sku]['priceDealer'], 
                float(priceData[sku]['priceDealer']) * currencyRate,
                lowerPrice - float(priceData[sku]['priceDealer']) * currencyRate)
            
            seoUrl = self._seoPrefix + webData[sku]['seoUrl'] 
            
            h = hashlib.sha1()
            h.update(seoUrl)
            skuForWeb = h.hexdigest() 
            
            rr[1] = cat 
            rr[2] = webData[sku]['product']
            rr[3] = priceUahSale
            rr[4] = priceUah
            rr[5] = pricesOther
            rr[6] = webData[sku]['option']
            rr[7] = options
            rr[8] = self._seoPrefix + webData[sku]['seoUrl']
            rr[9] = webData[sku]['description']
            rr[10] = int(priceData[sku]['balance'])
            rr[11] = brand
            rr[12] = skuForWeb
            idx = 20 - 1
            assert(webData[sku]['images'])
            for image in webData[sku]['images']:
               
                rr[idx] = image
               
                idx += 1

            log.debug('add row: %s', str(rr))
            xml.addrow(rr)
        
        
        if failedCategories:
            
            failedCategories.sort()
        
            log.error("categories not found: %s", '\n'.join(failedCategories))
            
            raise RuntimeError("categories not found: %s", '; '.join(failedCategories))
        
                            
        # save to file
        xml.write(xmlFilePath)
            
    def MapCategory(self, categoryName):
        
        return self._categoryMap[str(categoryName).replace('  ', ' ').strip()]

    def MapCategoryExt(self, categoryName, extInfo):
        
        cat = str(categoryName).replace('  ', ' ')
        
        for row in self._categoryMapExt:
            if row['name'] == cat and row['mustHave'] in extInfo:
                return row['target']
                
        
        raise KeyError('can not map category ext: %s\n%s' % (categoryName, extInfo))
        
    def ConnectToServer(self):
        
        ssh = paramiko.SSHClient() 
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())                        
        
        ssh.connect(conf.SSH_ADDR, port=conf.SSH_PORT, username=conf.SSH_LOGIN)
        
        
        stdin_, stdout_, stderr_ = ssh.exec_command("docker inspect opencart | grep merged | sed -r -e 's/.*?: \"//' -e 's/\",.*//'" )
        assert(not stdout_.channel.recv_exit_status())
        
        out = stdout_.read().strip()
        
        log.debug("[+] path to docker volume: %s", out)
        
        ssh.dockerVolumePath = out
        
        return ssh

        
    def UploadToServer(self, localFile, remoteFile, addDockerPrefix=False):
        
        ssh = self.ConnectToServer() 
                        
        sftp = ssh.open_sftp()
        
        if addDockerPrefix:
            remoteFile = ssh.dockerVolumePath + remoteFile
        
        log.info('uploading %s -> %s' % (localFile, remoteFile))
        
        sftp.put(localFile, remoteFile)
        sftp.close()
        ssh.close()
                
    def DownloadFromServer(self, remoteFile, localFile, addDockerPrefix=False):

        ssh = self.ConnectToServer()
        sftp = ssh.open_sftp()

        if addDockerPrefix:
            remoteFile = ssh.dockerVolumePath + remoteFile

        
        log.info('downloading %s <- %s' % (localFile, remoteFile))
        
        try:
            sftp.get(remoteFile, localFile)
        except IOError as io:
            
            log.error("download IOError: %s", str(io))
            
#             if 'IOError: [Errno 2] No such file' in str(io):
#                 raise AdminFileNotFound(remoteFile) 
            
        sftp.close()
        ssh.close()
        
    def AddWaterMarkToImage(self, inputFile, outputFile):

        dealer = self._d
        assert dealer

        im = Image.open(inputFile)

        mark = Image.open(self._waterMark)

        layer = Image.new('RGBA', im.size, (0, 0, 0, 0))

        position = (im.size[0] - mark.size[0], im.size[1] - mark.size[1])

        layer.paste(mark, position)

        newImg = Image.composite(layer, im, layer)

        newImg.save(outputFile)

    def DoWatermark(self, imagesRoot):

        log.info('do watermarks... (%s)' % self._d)

        for root, dirs, files in os.walk(imagesRoot):

            for f in files:

                f2 = os.path.join(root, f)
                log.debug('watermark add... %s %s\n' % (root, f))
                self.AddWaterMarkToImage(f2, f2)
                
    def AddWaterMarkToAllImages(self):
    
        remoteDirImageData = self._remoteDirImageData
        remoteImageDir = self._remoteImageDir
        remoteImageCacheDir = self._remoteImageCacheDir

        log.info('connecting to ssh...')
        
        ssh = self.ConnectToServer()
        
        docPath = ssh.dockerVolumePath
         
        log.info('archiving...')
        stdin_, stdout_, stderr_ = ssh.exec_command("cd %s; tar cfvj %s.tar.bz2 %s" % (
            docPath + remoteDirImageData, remoteImageDir, remoteImageDir))
        assert(not stdout_.channel.recv_exit_status())
 
        log.info('downloading...')
        sftp = ssh.open_sftp()
        sftp.get('%s/%s.tar.bz2' % (docPath + remoteDirImageData, remoteImageDir), 
                 os.path.join(_CACHE_PATH, '%s.tar.bz2' % remoteImageDir))
         
        log.info('removing from server...')
        stdin_, stdout_, stderr_ = ssh.exec_command('cd %s; rm -rf %s.tar.bz2' %(
            docPath + remoteDirImageData, remoteImageDir))
        assert(not stdout_.channel.recv_exit_status())
         
        log.info('extracting...')
        subprocess.check_call('cd %s; tar xvf %s.tar.bz2 > /dev/null; rm -rf %s.tar.bz2' % (
            _CACHE_PATH, remoteImageDir, remoteImageDir), shell=True)
         
        self.DoWatermark(os.path.join(_CACHE_PATH, remoteImageDir))
          
        log.info('archiving again...')
        subprocess.check_call('cd %s; tar cfvj %s-w.tar.bz2 %s > /dev/null' % (
            _CACHE_PATH, remoteImageDir, remoteImageDir), shell=True)
         
        log.info('uploading to server')
        sftp.put(os.path.join(_CACHE_PATH, '%s-w.tar.bz2' % remoteImageDir), 
                 '%s/%s-w.tar.bz2' % (docPath + remoteDirImageData, remoteImageDir))
        
        log.info('removing localfile')
        subprocess.check_call('cd %s; rm -rf %s-w.tar.bz2 %s' % (
            _CACHE_PATH, remoteImageDir, remoteImageDir), shell=True)
        
        log.info('extracting on server...')
        log.info("cd %s; tar xvf %s-w.tar.bz2; chown -R 33:33 %s; rm -rf %s-w.tar.bz2" % (
            docPath + remoteDirImageData, remoteImageDir, remoteImageDir, remoteImageDir))        
        stdin_, stdout_, stderr_ = ssh.exec_command(
                 "cd %s; tar xvf %s-w.tar.bz2; chown -R 33:33 %s; rm -rf %s-w.tar.bz2" % (
            docPath + remoteDirImageData, remoteImageDir, remoteImageDir, remoteImageDir))
        
        log.debug("stdout:\n")
        log.debug(stdout_.read())
        log.debug("stderr:\n")
        log.debug(stdout_.read())
        
        #log.debug("stdout:\n", stdout_.readlines())
        #log.debug("stderr:\n", stderr_.readlines())
        
        assert(not stdout_.channel.recv_exit_status())
        
        log.info('clear cache...')        
        stdin_, stdout_, stderr_ = ssh.exec_command("cd %s; rm -rf *" % (
            docPath + remoteImageCacheDir))

        assert(not stdout_.channel.recv_exit_status())


    def AnalyzeErrorsTmpLines(self, data):
        
        categories = []
        categoriesMargin = []
        jsVideos = []
        
        lines = data.splitlines()
        
        totalCount = 0
        notFoundCategory = 0
        notFoundCategoryMargin = 0
        zeroManufacturer = 0
        jsVideo = 0
        invalidPrice = 0
        raiseStr = ''


        for l in lines:
            
            if not l.strip():
                continue
            
            totalCount += 1
                          
            s = re.search("The Product has not been added: .*? Category: '(.*?)' not found in your settings \(see page 'Category and margin'\)", l)
            if s:
                if s.group(1) not in categories:
                    categories.append(s.group(1))  
                notFoundCategory += 1
                continue
            
            s = re.search("Please, set folder for photo on page 'Category and margin'  for Category '(.*?)' Row ~= \d+", l)
            if s:
                if s.group(1) not in categoriesMargin:
                    categoriesMargin.append(s.group(1))  
                notFoundCategoryMargin += 1
                continue
                            
            s = re.search("Warning. Row ~= (\d*?) SKU = (.*?) Manufacturer: '0' not found", l)
            if s:
                zeroManufacturer += 1
                continue
                 
            
            s = re.search('curl .*? = Could not resolve host: moto.*?.ua#video_code_(\d+)', l)
            if s:
                if not s.group(1) in jsVideos:
                    jsVideo += 1
                    jsVideos.append(s.group(1))
                continue
            
                                      
            s = re.search('Download.*?photo fails.*? Url.*? http://moto.*?ua#video_code_(\d+)', l)
            if s: 
                if not s.group(1) in jsVideos:
                    jsVideo += 1
                    jsVideos.append(s.group(1))
                continue
            
            s = re.search('The Product passed: Row ~= 1 SKU = seoUrl + dealerMark Invalid price of product = 0', l)
            if s:
                # just skip the header
                continue
        
            s = re.search('The Product passed: Row ~= \d+ SKU = (.*?) Invalid price of product = ', l)
            if s:
                log.info('[!] invalid price of product %s', s.group(1))
                invalidPrice += 1
                continue
            
            raiseStr += 'unknown error.txt line: %s\n' % l
        
        if raiseStr:
            raise ValueError(raiseStr)
                
        categories.sort()
        
        log.info('errors.txt analyzed, lines: %d, categories: %d, manufacturers: %d, js videos: %d, invalid prices: %d, category not set: %d',
                 totalCount, notFoundCategory, zeroManufacturer, len(jsVideos), invalidPrice, notFoundCategoryMargin)
                
        if categories:
            log.error("%d categories does not exists (%s)", len(categories), '\n'.join(categories))
            
        if categoriesMargin:
            log.error("%d categories not set: (%s)", len(categoriesMargin), '\n'.join(categoriesMargin))
            
        return categories


    def AnalyzeErrorsTmp(self, errorsFilePath):
        
        
        fileData = FileHlp(errorsFilePath, 'r').read()
        
        categories = self.AnalyzeErrorsTmpLines(fileData)
        
        if categories:
            
            xml = Xml2003FileStub()
            idx = 1
            
            for cc in categories:
                
                if '!!!' in cc:
                    continue
                
                row = ['' for i in range(0, 20)]
            
                try:
                    main, sub = cc.split(' | ')
                except ValueError:
                    main = cc
                
                row[0] = main
                row[1] = sub
                row[19] = str(idx)
                
                xml.addrow(row)
                
                idx += 1
        
        
            xml.write([_CACHE_PATH, '1.xml'])
            
    def AnalyzeReportTxt(self, reportFilePath):
        
        fileData = FileHlp(reportFilePath, 'r').read()
        
        lines = fileData.splitlines()
        
        totalCount = 0
        addCount = 0
        updatedCount = 0
        
        for l in lines:
            
            if not l.split():
                continue
            
            totalCount += 1

            
            s = re.search('Row =~ (\d*?) SKU = (.*?) .*?Price updated', l) 
            if s:
                log.debug('product updated: line = %s, sku = %s', s.group(1), s.group(2))
                updatedCount += 1
                continue
                
            s = re.search('Row =~ (\d*?) SKU = (.*?) .*?Product added', l) 
            if s:
                log.debug('product added: line = %s, sku = %s',s.group(1), s.group(2))
                addCount += 1
                continue
            
            
            
            raise ValueError('invalid line format: %s' % l)
                
        log.info('[+] report.txt analyzed, total:%d updated: %d added: %d' % (
            totalCount, updatedCount, addCount))
        
        

class MXShopKopyl(MXShopZhovtuha):
    
    _d = 'kopyl'
    
    _seoPrefix = 'kop-'        
    
    _webAdminId = 'kopyl-2'    
    _remoteImageDir = 'kop'
    
    _waterMark = 'image/kopyl-watermark.png'
    
    _categoryMap = {
        "Аксессуары | LED оптика": '!!! null !!!',
        "Аксессуары | Наклейки": "Аксессуары | Наклейки",
        "Аксессуары | Сопутствующее": "Аксессуары | Другое",
        "Аксессуары | Защита рук": "Запчасти | Другое",
        "Запчасти и расходники | Грипсы": "Запчасти | Грипсы",
        "Запчасти и расходники | Звезды задние": "Запчасти | Цепи и звезды",
        "Запчасти и расходники | Звезды передние": "Запчасти | Цепи и звезды",
        "Запчасти и расходники | Комплекты пластика": "Запчасти | Другое",
        "Запчасти и расходники | Мото цепи": "Запчасти | Цепи и звезды",
        "Запчасти и расходники | Рули/Clip-Ons": "Запчасти | Рули",
        "Запчасти и расходники | Слайдера/Ловушки/Ролики": "Запчасти | Другое",
        "Запчасти и расходники | Тормозные колодки": "Запчасти | Тормозные колодки",
        "Масла и химия | Антифризы": "Химия | Другое",
        "Масла и химия | Масла гидравлические": "Химия | Масло в подвеску",
        "Масла и химия | Масла моторные 2T": "Химия | Моторное масло",
        "Масла и химия | Масла моторные 4T": "Химия | Моторное масло",
        "Масла и химия | Масла моторные ATV": "Химия | Моторное масло",
        "Масла и химия | Масла моторные V-Twin": "Химия | Моторное масло",
        "Масла и химия | Масла трансмиссионные": "Химия | Другое",
        "Масла и химия | Масла фильтровые": "Химия | Для воздушного фильтра",
        "Масла и химия | Сервисные продукты": "Химия | Другое",
        "Масла и химия | Смазки цепей": "Химия | Для цепи",
        "Масла и химия | Тормозные жидкости": "Химия | Другое",
        "Мото защита | Боты": "Мотоботы | Кроссовые",
        "Мото защита | Защита тела": "Защита | Груди и спины",
        "Мото защита | Защита шеи": "Защита | Шеи",
        "Мото защита | Наколенники": "Защита | Коленей",
        "Мото защита | Налокотники": "Защита | Локтей",
        "Мото защита | Очки": "Очки | Кроссовые",
        "Мото защита | Пояса": "Защита | Пояса",
        "Мото защита | Шлемы": "Шлема | Кроссовые",
        "Мото защита | Шорты": "Защита | Шорты",
        "Мото экипировка | Джерси": "Форма | Джерси",
        "Мото экипировка | Куртки": "Дорожная экипировка | Куртки",
        "Мото экипировка | Носки": "Форма | Носки",
        "Мото экипировка | Перчатки": "Форма | Перчатки",
        "Мото экипировка | Сумки": "Аксессуары | Сумки",
        "Мото экипировка | Штаны": "Форма | Штаны",
        "Брюки/Джинсы": "Одежда | Штаны",
        "Брюки/Джинсы/Штаны": "Одежда | Штаны",
        "Кепки": "Одежда | Кепки",
        "Куртки": "Одежда | Куртки",
        "Носки": "Одежда | Другое",
        "Обувь": "Одежда | Обувь",
        "Очки": "Очки | Солнцезащитные",
        "Полотенца": "Одежда | Другое",
        "Пояса/Кошельки": "Одежда | Другое",
        "Рубахи": "Одежда | Рубахи",
        "Сумки/Рюкзаки": "Аксессуары | Сумки",
        "Толстовки/Свитера": "Одежда | Толстовки",
        "Футболки": "Одежда | Футболки",
        "Шапки": "Одежда | Шапки",
        "Шорты": "Одежда | Шорты",
        "Аксессуары | Защита двигателя": "Запчасти | Другое",
        "Аксессуары | Сервисные продукты": "Аксессуары | Другое",
        "Мото защита | Очки - расходники": "Очки | Аксессуары к очкам",
        "Запчасти и расходники | Пластик/Комплекты пластика": "Запчасти | Другое",
        "Зонты": "Аксессуары | Другое",
        }
    
    _redirectByName = {
        'Грязевая система': 'Очки | Аксессуары к очкам',
        'Срывки': 'Очки | Аксессуары к очкам',
        'Линза к очкам': 'Очки | Аксессуары к очкам',
        'Грязевая сменная': 'Очки | Аксессуары к очкам',
        'Сменная линза': 'Очки | Аксессуары к очкам',
        'Термобелье': 'Форма | Термобелье',
        'Вкладыши для мото': 'Шлема | Запчасти к шлему',
        'Козырек для мото': 'Шлема | Запчасти к шлему',
        'buckle': 'Мотоботы | Запчасти к мотоботам',
        'boot strap': 'Мотоботы | Запчасти к мотоботам',
        'strap kit': 'Мотоботы | Запчасти к мотоботам',
        'Защитный бандаж': 'Защита | Другое',
        'Ligament Set': 'Защита | Другое',
        'Patella Guard': 'Защита | Другое',
        'Refurb Set': 'Защита | Другое',
        'Охлаждающая футболка': 'Форма | Термобелье',
        'Мото перчатки FOX Bomber Glove': 'Дорожная экипировка | Перчатки',
        'Мотоботы FOX BOMBER BOOT': 'Мотоботы | Дорожные',
        'Мотоперчатки SHIFT Hybrid Delta': 'Дорожная экипировка | Перчатки',
        'Mотоперчатки SHIFT Super Street': 'Дорожная экипировка | Перчатки',
        'Мото куртка SHIFT Super Street Textile Jacket': "Дорожная экипировка | Куртки",
        'Вело перчатки FOX Static Wrist Wrap': 'Дорожная экипировка | Перчатки',
        'Мото куртка SHIFT Moto R Textile Jacket': 'Дорожная экипировка | Куртки',
        'Мото штаны SHIFT Squadron Pant': 'Дорожная экипировка | Штаны',
        'Мото штаны FOX NOMAD Pants': 'Дорожная экипировка | Штаны',
        'Защита крышки': "Запчасти | Другое",
        'Защита двигателя': "Запчасти | Другое",
        'Защита свингарма': "Запчасти | Другое",
    }
    
    def __init__(self, **kw):

        MXShopZhovtuha.__init__(self, **kw)
        
        self._priceFileRE = '^dealer_price_(\d+)-(\d+)-(\d+) (\d+)\.(\d+)\.(\d+)\.xls$'  # example: dealer_price_2017-02-04 22.49.00.xls
        self._priceFileREidx = {'year': 1, 'month': 2, 'day': 3, 'hour': 4, 'minute': 5, 'second': 6}


    def DownloadCurrentPriceFromWeb(self):
        
        br = mechanize.Browser()
        
        br.addheaders = [("User-agent", "Mozilla/5.0 (compatible;)")]
 
        br.set_handle_referer(True)
        br.set_handle_robots(False)
        
        rootSite = conf.KOPYL_HOST
        
        log.debug('login to %s...', rootSite)
        
        r = br.open(rootSite)
        
        FileHlp([_CACHE_PATH, '%s-root.html' % self._d], 'w').write(r.read())
        
        log.debug('root response code: %d', r.code)
        
        br.select_form(nr=0)
        br.form['login'] = conf.KOPYL_LOGIN
        br.form['pass'] = conf.KOPYL_PASS
        
        r = br.submit()
        
        data = r.read()
        FileHlp([_CACHE_PATH, '%s-login.html' % self._d], 'w').write(data)
                
        log.debug('login response code: %d', r.code)
        
        soup = BeautifulSoup(data, "html.parser")
        
        # USD = 27.1 грн.
        currencyRate = 0
        s = re.search('USD = (\d+[\.]*\d+) грн\.', str(soup.text), re.MULTILINE)
        if s:
            currencyRate = float(s.group(1))
            log.info('price currencyRate = %.2f' % currencyRate)
        else:
            raise ValueError('currency not found, text: """%s"""' % str(soup.text))
        
        for link in soup.find_all('a'):
            if 'Текущий прайс' == link.text: 
                r = br.open(link.get('href'))        
        
        # "dealer_price_2017-02-11 13.37.28.xls"
                
        s = re.search('filename="(.*?\.xls)"', str(r.info()), re.DOTALL)
        fileName = s.group(1)
                
        log.debug('price response code: %d', r.code)
        
        data = r.read()
        FileHlp([_CACHE_PATH, fileName], 'w').write(data)
        
        log.info('new kopyl price downloaded: %s', fileName)
        
        return data, fileName, currencyRate

    def ReadPrice(self, xlsPath):
        
        log.debug('reading price %s...' % xlsPath)
        

        rb = xlrd.open_workbook(xlsPath, formatting_info=True)

        sheet = rb.sheet_by_index(0)

        log.debug('----------------------------------------- table dump begin -----------------------------------------')
 
        # assert price format
         
        row1 = sheet.row_values(0)
         
        assert(row1[0] == 'Товарный код')
        assert(row1[1] == 'Наименование')
        assert(row1[2] == 'Наименование')
        assert(row1[3] == 'Кол-во')
        assert(row1[4] == 'Заказ')
        assert(row1[5] == 'ОПТ')
        assert(row1[6] == 'РРЦ')
        assert(row1[7] == 'Цена РОЗН старая')
        assert(row1[8] == 'Итого')
        
        invalidPricesCount = 0
        zeroCellsCount = 0
        salesCount = 0
        duplicateCount = 0
        
        result = {}
        
        for rownum in range(1, sheet.nrows):
            
            priceIdx = str(rownum + 1)
            
            row = sheet.row_values(rownum)
            
            if type(row[0]) == float:
                sku = '%d' % row[0]
            else:
                sku = str(row[0]).strip()
            category = str(row[1]).strip()
            product = str(row[2]).strip()
            
            b = str(row[3]).strip()
            if b:
                b = int(float(b))
            balance = b
            
            priceDealer = str(row[5]).strip()            
            priceRetail = str(row[6]).strip()
            saleOff = str(row[7]).strip()
            
            isInvalid = ''
            
            if not sku or not category or not priceRetail or not priceDealer or not saleOff:
                      
                zeroCellsCount += 1
                isInvalid = 'one of essential cells is empty'
        
            if priceRetail == priceDealer:
                invalidPricesCount += 1
                isInvalid = 'price retail = price dealer'
                                
            if isInvalid:
                log.debug('[!] %s; %s; %s; %s; %s; %s; [! invalid line]' % (
                    priceIdx, sku, priceRetail, priceDealer, category, product,))
                continue
            
            assert(len(sku) > 3)
            
            if float(saleOff) == float(priceRetail):
                saleOff = '0'
                log.warn("sku [%s] sale price = retail price = %s", sku, priceRetail)

            if float(saleOff): 
                salesCount += 1 
                priceRetail, saleOff = saleOff, priceRetail   

            element = {'sku': sku, 
                       'priceRetail': priceRetail,
                       'priceDealer': priceDealer,
                       'priceSale': saleOff if float(saleOff) else '',
                       'balance': balance,
                       'categoryFromPrice': category,
                       'productFromPrice': product}
                 
            if float(saleOff):
                log.debug('%s; %s; %s; %s; [SALE: %s]; %s; %s;' % (priceIdx, sku, priceRetail, priceDealer, saleOff, category, product))
            else:
                log.debug('%s; %s; %s; %s; %s; %s;' % (priceIdx, sku, priceRetail, priceDealer, category, product,))
 
            if not sku in result.keys():            
                result[sku] = element
            else:
                 
                if result[sku] == element:
                    duplicateCount += 1
                    log.debug('duplicated line in price %s', sku)
                else:
                    raise ValueError('dupicate sku %s, but not equal data' % sku)
                
            result[sku] = element
                
        log.debug('----------------------------------------- table dump end   -----------------------------------------')

        log.debug('%s price processed\n total rows: %d, skus: %d, zero cells: %d, invalid price in sku %d, duplicate sku: %d' % (
             xlsPath, sheet.nrows, len(result.keys()), zeroCellsCount, invalidPricesCount, duplicateCount))
        
        if salesCount:
            log.debug('[*] sale off count = %d', salesCount)

        return result 

    def GetInfoMotoKopylbrosCom(self, element, **kw):
        
        cache = HttpPageCache('moto-kopylbros-com')
        
        cacheJson = HttpPageCache('moto-kopylbros-com-json', dbFile='values-json.db')

        jsonResult = cacheJson.get(element['sku'])
        if kw.get('noCache', False):
            cacheJson.drop(element['sku'])
            jsonResult = ''
        
        if jsonResult:
            return json.loads(jsonResult)
        
        baseUrl = conf.KOPYL_HOST
        
        searchUrl = baseUrl + '/search'
        
        s = requests.Session()
        
        if kw.get('searchWithName', False):
            params = {'search_text': element['productFromPrice']}
        else:
            params = {'search_text': element['sku']}
        req = requests.Request('POST', searchUrl, data=params)
        prepped = s.prepare_request(req)        
                   
        cached = cache.get(prepped.url + '/' + prepped.body)

        if kw.get('noCache', False):
            cache.drop(prepped.url + '/' + prepped.body)
            cached = ''
        
        if not cached:
            resp = s.send(prepped)
            
            log.debug('%s [POST: %s] -> %s', prepped.url, prepped.body, resp.status_code)
            
            cached = resp.text
            cache.put(prepped.url + '/' + prepped.body, cached)
        
        soup = BeautifulSoup(cached, 'lxml')
        
        divs = soup.findAll("div", { "class" : "item"})
        
        if len(divs) == 0:
            raise WebSearchNotFound('can not find any results %s' % element['sku'])
        elif len(divs) > 1:
            log.info('multiple results found for "%s"', element['sku'])
        
        gotSku = []
        
        for div in divs:
        
            seoUrl = div.a['href'].replace('/products/', '')
            assert(seoUrl)
            
            url = baseUrl + div.a['href']
            cached = cache.get(url)
            
            if kw.get('noCache', False):
                cache.drop(url)
                cached = ''
                     
            if not cached:    
                resp = s.get(url)
                
                log.debug('%s -> %s', url, resp.status_code)
                
                cached = resp.text
                cache.put(url, cached)
            
            soup = BeautifulSoup(cached, 'lxml')
            
            # assure sku
            
            sku = ''
            option = ''
            price = ''
            options = {}
            
            tag = soup.findAll('table', {'cellpadding': '0', 'cellspacing': '2'})
            assert(len(tag) == 1)
            
            for tr in tag[0].findAll('tr'):
                            
                tds = tr.findAll('td')
                
                if len(tds) != 5:
                    continue
                
                possibleSku = tds[0].text.strip()
                possibleBalance = tds[1].text
                possibleOption = tds[2].text
                possiblePrice = tds[3].text
                
                gotSku.append(possibleSku)
                
                if not '---' in possibleOption and not possibleOption in options:
                    options[possibleSku] = possibleOption
                
                if possibleSku == element['sku']:
                    sku = possibleSku
                elif possibleSku.replace('_', '-') == element['sku']:
                    log.warning("[!] sku deduction: %s -> %s" % (possibleSku, element['sku']))
                    sku = possibleSku

                elif possibleSku[1:] == element['sku'] and possibleSku[0] == '0': # drop first zero
                    log.warning('first zero from sku is dropped')
                    sku = possibleSku 
                else:
                    continue
                     
                if not '---' in possibleOption:
                    option = possibleOption
                    
                price = possiblePrice
                    
                if 'В наличии' in possibleBalance:
                    if element['balance'] == '0':
                        raise WebSearchPriceFound('found product on web, but in price it seems to be sold out')
                
            if not sku:

                skuParts = gotSku[0].split('-') 
                
                skuElementParts = element['sku'].split('-')
                
                if skuParts[0] == skuElementParts[0] and skuParts[1] == skuElementParts[1]:
                    
                    log.info('[!] sku mismatch, but size deducted successfully (%s - %s)' %
                             (element['sku'], ' '.join(gotSku)))
                    
                    sku = element['sku']
                    
                else:
                
                
                    continue
    
            webElement = {}
            webElement['category'] = ''
            webElement['sku'] = sku
            webElement['saleOffPercent'] = ''
            webElement['product'] = ''
            webElement['description'] = ''
            webElement['images'] = ''
            webElement['option'] = option
            webElement['options'] = options
            webElement['seoUrl'] = seoUrl
            webElement['price'] = price # optional
            webElement['brand'] = '' # optional
            
            # find category
            
            tags = soup.findAll("li", { "class" : "active"})
            categoryList = [tag.a.text.strip() for tag in tags]
                    
            webElement['category'] = ' | '.join(categoryList)
            if not webElement['category']:
                raise WebSearchNoCategory("no category for sku = %s, url = %s/products/%s" % (
                    sku, baseUrl, seoUrl))
            assert(webElement['category'])
    
            # find saleOff (%)
            # do really we need it?
    
            # find name
    
            tags = soup.findAll("h1")
            assert(len(tags) == 1)
            webElement['product'] = tags[0].b.text
            assert(webElement['product'])
            
            # find description
    
            # <div class="good_description">
            
            divs = soup.findAll("div", { "class" : "good_description"})
    
            # one in one bug:
            
            # <h2 id="description">Описание Товара</h2>
            #  <div class="good_description">
            #   <div class="good_description">
            #    <div class="good_description">
            #     <div style="text-align: justify;"><span style="font-size: small;">Мото руль Renthal Fatbar является лидиром на рынке рулей без перемычек. Благодаря коническому профилю стенки, конструкция позволяет использовать диаметр зажима (1-1/8"| 28.6mm) что является большим, чем у стандартных рулей (7/8" | 22.2mm).</span></div>
            #     <div><span style="font-size: small;"><br /></span></div>
            #    </div>
            #   </div>
            #  </div>
            
            while True:
                
                n = divs[0].findAll("div", { "class" : "good_description"})
                if n:
                    log.debug('found "one in one" bug in good_description tag')
                    divs = n
                else:
                    break
                  
            
            assert(len(divs) == 1)
            
            description = ''.join([str(tags).strip() for tags in divs[0].contents])
            description = description.replace('\n', '').replace('\r', '')
    
            webElement['description'] = description
            assert(webElement['description'])
            
            # find images
    
            imgList = []
            
            tags = soup.findAll('div', {'class': 'available_colors'})
            
            if tags:        
            
                for aa in tags[0].findAll('a'):
                    assert(aa['href'])
                    imageUrl = baseUrl + aa['href'] 
                    if not imageUrl in imgList:        
                        imgList.append(imageUrl)
                    else:
                        raise WebSearchDuplicatedImage("duplicate image found %s [%s]" % (imageUrl, str(imgList)))
                        
            if not imgList:
                tags = soup.findAll('div', {'class': 'big_pic'})
                assert(len(tags) == 1)
                
                for aa in tags[0].findAll('a'):
    
                    assert(aa['href'])
                    imageUrl = baseUrl + aa['href']
                    if not imageUrl in imgList:        
                        imgList.append(imageUrl)
            
            webElement['images'] = imgList                  
            
            # find brand
            
            #<div class="small_img">
            
            tags = soup.findAll('div', {'class': 'small_img'})
            if tags:
                assert(len(tags) == 1)
                brand = tags[0].img['alt'].strip()
                webElement['brand'] = brand
    
            cacheJson.put(element['sku'], json.dumps(webElement))
                  
            return webElement
        
        raise WebSearchSkuNotMatched('[kopyl] sku seems invalid "%s", values (%s)' % (
                                     element['sku'], ' '.join(gotSku)))

        
    def GrabWebData(self, priceData):
        
        webData = {}
        i = 0
        notFoundCount = 0
        skuNotMatched = 0
        skuNoCategory = 0
        
        keysLen = len(priceData.keys())
        
        for sku in priceData.keys():
             
            i += 1
            
            sys.stdout.write('%d/%d\r' % (i, keysLen))
            sys.stdout.flush()
             
            #log.info('%d/%d', i, keysLen)
             
            try:
                element = self.GetInfoMotoKopylbrosCom(priceData[sku])
            except WebSearchNotFound as ex:
                
                try:
                    element = self.GetInfoMotoKopylbrosCom(priceData[sku], searchWithName=1)
                except WebSearchNotFound as ex2:
                    ex = ex2
                    
                notFoundCount += 1
                log.warning('[!] %s', (str(ex)))
                raise ex
                
            except WebSearchSkuNotMatched as ex:
                
                log.warning('[~] %s, retry with no cache...', str(ex))
                
                try:
                    element = self.GetInfoMotoKopylbrosCom(priceData[sku],
                                                           noCache=1)
                except WebSearchSkuNotMatched as ex2:
                    skuNotMatched += 1
                    log.warning('[!] %s', (str(ex)))
                    raise ex
                

                            
            except WebSearchNoCategory as ex:
                
                skuNoCategory += 1
                
                log.error("[!] skipped: %s", str(ex))
                continue
             
            assert(element)
            assert(element['images'])
                
            webData[sku] = element
            
        log.info('web grab for kopyl completed')
        log.info('searched for %d items, sku not matched - %d, not found - %d, no category - %d', 
                 i, skuNotMatched, notFoundCount, skuNoCategory)

        return webData

                
    

class testDirectoriesZhovtuha(unittest.TestCase):
    
    def runTest(self):
        
        if sys.version_info[0] == 3:     
            from unittest.mock import MagicMock
        else:
            from mock import MagicMock
        
        def walkDummy(root, files):                 
            yield (root, None, files)

        
        #
        # sort test
        #

        log.info("zhovtuha test")
        
        zhov = MXShopZhovtuha()
        root = zhov.GetPricesOrigDir()
        
        walkMock = MagicMock()
        files = ['Остатки-26.01.11.xls', 
                 'Остатки-26.02.12.xls',
                 'Остатки-25.01.15.xls', 
                 'Остатки-25.03.14.xls',
                 'Остатки-23.01.15.xls',
                 'Остатки-24.01.15.xls']
        
        attrs = {'return_value': walkDummy(root, files)}
        walkMock.configure_mock(**attrs)
                    
        zhov.SetWalkMock(walkMock)
        res = zhov.GetAllPriceNames()

        self.assertEqual(res[-1], os.path.join(root, 'Остатки-25.01.15.xls')) 
        
        walkMock.assert_called_with(root)
        self.assertEqual(len(walkMock.mock_calls), 1)
        
        #
        # invalid param test 1
        #  
        
        zhov = MXShopZhovtuha()               
        root = zhov.GetPricesOrigDir()
        
        walkMock = MagicMock()
        files = ['Остатки-26.01.11.xls']
        attrs = {'return_value': walkDummy(root, files)}
        walkMock.configure_mock(**attrs)
                    
        zhov.SetWalkMock(walkMock)
        res = zhov.GetAllPriceNames()

        self.assertEqual(res[-1], os.path.join(root, 'Остатки-26.01.11.xls')) 
        
        walkMock.assert_called_with(root)
        self.assertEqual(len(walkMock.mock_calls), 1)


        #
        # invalid param test 2
        #  

        zhov = MXShopZhovtuha()
        root = zhov.GetPricesOrigDir()
              
        walkMock = MagicMock()        
        files = []
        attrs = {'return_value': walkDummy(root, files)}
        walkMock.configure_mock(**attrs)
            
        zhov.SetWalkMock(walkMock)
        res = zhov.GetAllPriceNames()

        self.assertFalse(res)

        walkMock.assert_called_with(root)
        self.assertEqual(len(walkMock.mock_calls), 1)
        
        #
        # invalid param test 3
        #  
        
        zhov = MXShopZhovtuha()               
        root = zhov.GetPricesOrigDir()
        
        walkMock = MagicMock()        
        files = ['Остати-26.01.11.xls']
        attrs = {'return_value': walkDummy(root, files)}
        walkMock.configure_mock(**attrs)
        zhov.SetWalkMock(walkMock)
                
        isNameError = False        
        try:
            res = zhov.GetAllPriceNames()
        except NameError:
            isNameError = True

        self.assertTrue(isNameError)
        walkMock.assert_called_with(root)
        self.assertEqual(len(walkMock.mock_calls), 1)
        
        
class testDirectoriesKopyl(unittest.TestCase):

    def runTest(self):
        
        if sys.version_info[0] == 3:     
            from unittest.mock import MagicMock
        else:
            from mock import MagicMock
       
        def walkDummy(root, files):                 
            yield (root, None, files) 
        
        log.info("kopyl test")

        kop = MXShopKopyl()
        root = kop.GetPricesOrigDir()

        walkMock = MagicMock()        
        files = ['dealer_price_2017-02-04 22.49.00.xls',
                 'dealer_price_2017-02-04 22.49.01.xls',
                 'dealer_price_2016-02-04 22.49.00.xls',
                 'dealer_price_2017-01-04 22.49.00.xls',
                 'dealer_price_2017-02-03 22.49.00.xls']
        attrs = {'return_value': walkDummy(root, files)}
        walkMock.configure_mock(**attrs)
            
        kop.SetWalkMock(walkMock)
        res = kop.GetAllPriceNames()
                
        self.assertEqual(res[-1], os.path.join(root, 'dealer_price_2017-02-04 22.49.01.xls'))

        walkMock.assert_called_with(root)
        self.assertEqual(len(walkMock.mock_calls), 1)

        
class OFF_testPriceDownloadKopyl(unittest.TestCase):
    
    def runTest(self):
        
        
        log.info('kopyl download test')
        
        kop = MXShopKopyl()
        
        r = kop.DownloadCurrentPriceFromWeb()
        
        self.assertTrue(r)
        
class OFF_testPriceDownloadZhovtuha(unittest.TestCase):
    
    def runTest(self):
        
        log.info('zhovtuha download test')
        
        zhov = MXShopZhovtuha()
        r = zhov.DownloadCurrentPriceFromWeb()
            
        self.assertTrue(r)
        
class OFF_testReadPriceFromFileZhovtuha(unittest.TestCase):
    
    def runTest(self):
        
        log.info('zhovtuha read test')
        
        testPricePath = os.path.join('prices', 'test', 'zhovtuha', 'Остатки-24.01.17.xls')
        
        zhov = MXShopZhovtuha()

        priceData = zhov.ReadPrice(testPricePath)
        
        # TODO: add more precise assertions 
        
        self.assertTrue(priceData)


class OFF_testReadPriceFromFileKopyl(unittest.TestCase):
    
    def runTest(self):
        
        log.info('kopyl read test')
        
        testPricePath = os.path.join('prices', 'test', 'kopyl', 'dealer_price_2017-02-04 22.49.00-xls2003.xls')
        
        kop = MXShopKopyl()

        priceData = kop.ReadPrice(testPricePath)
        
        self.assertTrue(priceData)


class testProcessCycleKopyl(unittest.TestCase):
    
    def runTest(self):
        

        kop = MXShopZhovtuha()
        
        log.info('%s cycle test' % kop._d)        
        
        remoteXmlFile = kop.WebAdminGetRemoteXmlName()
                
        data, fileName, currencyRate = kop.DownloadCurrentPriceFromWeb()
                
        log.info('price downloaded: %s', fileName)
         
        xlsFilePath = os.path.join('prices', 'orig', kop._d, fileName)
        xmlFilePath = os.path.join('prices', 'result', kop._d, 
                                      os.path.splitext(fileName)[0] + '.xml')
         
        newFilePath = os.path.join(_CACHE_PATH, os.path.splitext(fileName)[0] + '.xml')
                 
        FileHlp(xlsFilePath, 'w').write(data)
         
        r = kop.ConvertTo97Xls(xlsFilePath, newFilePath)
        self.assertTrue(r)
         
        priceData = kop.ReadPrice(newFilePath)
        self.assertTrue(priceData)
        os.remove(newFilePath)
          
        webData = kop.GrabWebData(priceData)
        self.assertTrue(webData)
          
        kop.CreateXmlFile(priceData, webData, xmlFilePath, currencyRate)
          
        kop.UploadToServer(xmlFilePath, '%s/%s' % (kop._remoteUploadDir, remoteXmlFile))
           
        for idx in range(0, 9):
               
            try:  
                kop.WebAdminRunPrice()
                break
            except AdminNeedContinue:                
                log.info('not all data processed, make another request...')
                   
                idx += 1
                continue
            
        kop.DownloadFromServer('%s/errors.tmp' % kop._remoteUploadDir,
                                os.path.join(_CACHE_PATH, 'errors.txt'))
      
        kop.DownloadFromServer('%s/report.tmp' % kop._remoteUploadDir,
                                os.path.join(_CACHE_PATH, 'report.txt'))
           
        kop.AnalyzeErrorsTmp(os.path.join(_CACHE_PATH, 'errors.txt'))        
        kop.AnalyzeReportTxt(os.path.join(_CACHE_PATH, 'report.txt'))
        

class OFF_testProcessCycleZhovtuha(unittest.TestCase):
    
    def runTest(self):
        
        log.info('Zhovtuha cycle test...')
        
        zhov = MXShopZhovtuha()
        
        priceData, fileName, currencyRate = zhov.DownloadCurrentPriceFromWeb() 
        
        log.info('price downloaded: %s', fileName)
        
        xlsFilePath = os.path.join('prices', 'orig', zhov._d, fileName)
        xmlFilePath = os.path.join('prices', 'result', zhov._d, 
                                      os.path.splitext(fileName)[0] + '.xml')
        
        if os.path.exists(xlsFilePath):
            log.info('file already existed: %s, it seems there is no updates so far', 
                     xlsFilePath)
            #return
        else:            
            FileHlp(xlsFilePath, 'w').write(priceData)
            log.info('new price write ok: %s', xlsFilePath) 
        
        priceData = zhov.ReadPrice(xlsFilePath)
        
        webData = zhov.GrabWebData(priceData)
        self.assertTrue(webData)        
        
        zhov.CreateXmlFile(priceData, webData, xmlFilePath, currencyRate)
        zhov.CreateXmlFile(priceData, webData, '/home/vasya/Dropbox/mxshop/out-zhov.xml', currencyRate)
        
        zhov.UploadToServer(xmlFilePath, './public_html/newtest/admin/uploads/4.xml')    
        zhov.WebAdminRunPrice()
         
        zhov.DownloadFromServer('./public_html/newtest/admin/uploads/errors.tmp',
                                os.path.join(_CACHE_PATH, 'errors.txt'))
  
   
        zhov.DownloadFromServer('./public_html/newtest/admin/uploads/report.tmp',
                                os.path.join(_CACHE_PATH, 'report.txt'))
         
         
        zhov.AnalyzeErrorsTmp(os.path.join(_CACHE_PATH, 'errors.txt'))
        
        zhov.AnalyzeReportTxt(os.path.join(_CACHE_PATH, 'report.txt'))
        
        

class OFF_testEmptyDescriptionZhovtuha(unittest.TestCase):
    
    def runTest(self):
        
        zhov = MXShopZhovtuha()
        element = {'sku': '2058-350-036'}
        
        r = zhov.GetInfoMotocrazytownComUa(element)        
        
        self.assertFalse(repr(r['description']))

        
class OFF_testMultipleFoundZhovtuha(unittest.TestCase):
    
    def runTest(self):
        
        zhov = MXShopZhovtuha()
        element = {'sku': '70002'}
        
        r = zhov.GetInfoMotocrazytownComUa(element)
        
        self.assertTrue(r)

class OFF_testMultipleNotFoundZhovtuha(unittest.TestCase):
    
    def runTest(self):
        
        zhov = MXShopZhovtuha()
        element = {'sku': '7000'}
        
        isException = False
        
        try:
            zhov.GetInfoMotocrazytownComUa(element)
        except WebSearchSkuNotMatched:
            isException = True
        
        self.assertTrue(isException)
        
        
class OFF_testMultipleFoundMotostyle(unittest.TestCase):
    
    def runTest(self):
        
        zhov = MXShopZhovtuha()
        element = {'sku': '70002'}
        
        r = zhov.GetInfoMotostyleComUa(element)
        
        self.assertTrue(r)

class OFF_testMultipleNotFoundMotostyle(unittest.TestCase):
    
    def runTest(self):
        
        zhov = MXShopZhovtuha()
        element = {'sku': '7000'}
        
        isException = False
        
        try:
            zhov.GetInfoMotostyleComUa(element)
        except WebSearchSkuNotMatched:
            isException = True
        
        self.assertTrue(isException)
               
        
            
class OFF_testGetFromWebMotostyleComUa(unittest.TestCase):
    
    def runTest(self):
        
        testPricePath = os.path.join('prices', 'test', 'zhovtuha', 'Остатки-24.01.17.xls')
        
        zhov = MXShopZhovtuha()
        priceData = zhov.ReadPrice(testPricePath)

        notFoundSku = ['101019000084',
            'RSF3-AXIUM-L',
            '165 3329 XL 101',
            'FA185TT',
            '165 2051 XL 101',
            'MD6017D',
            '8117732120844 BLACK 45',
            'FORV180-98 white 46',
            '190 6183 XL 101',
            'RSJ2519900XXL',
            'BTLCRY BLK LG',
            '165 3605 XL 180',
            '310157',
            '1017807041040',
            'YZF600R6',
            'BLACK\RED',
            'MD6255C',
            'FORT71W-9998800 black/white/camo 46',
            'FORT71W-9998800 black/white/camo 45',
            'FORT71W-9998800 black/white/camo 44',
            'FORT71W-9998800 black/white/camo 43',
            'FORT71W-9998800 black/white/camo 42',
            'PLV078S',
            'DIAC590 XXL',
            'RSJ266150003XL',
            '1011536060020',
            'FORV150-10 red 42',
            '1012862020030',
            '165 3246 XS 810',
            '1012864010030',
            'TRV0379900',
            '17 13 001 0',
            'Merc',
            'FA185X',
            '1010800013320',
            'FORT750-99 S\M',
            'DIAC530 XXL',
            '2810-1465',
            '190 6202 XL 101',
            'VR-2 WIZARD M',
            '1012866020020',
            'KSHA00W3.4',
            'FORC29W-99 black 42',
            'FORC29W-99 black 41',
            'FORC29W-99 black 40',
            '2840-0042',
            'RSJ2669900XL',
            '1231-0222',
            '2855-0072',
            '1012881020030',
            '2839-007-008',
            '1657012180',
            '1010800013310',
            '1015129010020',
            'RSJ2579994    3XL',
            '165 3248  S 870',
            '1011536010030',
            '2855-0037',
            'KSHA0006.4',
            '14 06 171 3',
            '1017805041010',
            'FGS051-2250 blue/black L',
            '2152-304-063',
            '091095000211',
            'FGS051-2000 red/black L',
            'MD6037C',
            '190 6203  L 101',
            '1010105-XL',
            '2063-301-034',
            '1012864010020',
            'Overlord',
            '190 6204 XS 101',
            '1657016180',
            '11 11 110 5',
            '1014103362042',
            '20-262-02',
            '165 1303 XL 331',
            'FORC420-981011 white/red/blue 45',
            '330706',
            '2706-0070',
            '14 06 178 5',
            '1014106363042',
            '1231-0273',
            'MCT-XL',
            'KSCT00W6.6',
            '2839-007-005',
            '2839-007-007',
            '2839-007-006',
            'Taichi Dyna BLACK M',
            'SM6264C',
            '2833-001-008',
            '165 3329 XXL 101',
            '190 6203 XL 101',
            'DIAV210-10 RED 42',
            'FORC370-98 white 37',
            '80279',
            '1011536060040',
            '1017805040010',
            '2840-0043',
            'DIAC540 S',
            'FORT64W-99 black 42',
            'FGS046-0010 black L',
            '165 3246  M 110',
            '203088',
            '2839-010-008',
            '2839-010-007',
            '2839-010-006',
            'FORT87W-99 black 44',
            'FORT87W-99 black 45',
            'FORT87W-99 black 42',
            'FORT87W-99 black 43',
            'DIAC530 S',
            '2831-0045',
            '2810-005-011',
            'RSJ821-LAMB BROWN-L',
            'RSJ2709926XXL',
            'MERCSTAGE2',
            'Revit GT BLACK 48',
            'SILVER\BLACK',
            'FORV150-10 red 38',
            '1011536060030',
            '1012862020020',
            '2000000000626 BLACK 43',
            'FORW160-99 black 47',
            '42100-27820',
            '1015129010040',
            'MD6246C',
            'RSJ8259915 XL',
            '1011536010040',
            '8117732070842 BLACK 42',
            'MCT-L',
            '1011536230060',
            'Superduty',
            '165 3246  L 810',
            'FJT083 3520 XXL',
            '93562187',
            '1017806040040',
            'RSJ2759900M',
            '165 3246  S 110',
            'Icon Titan M',
            'RSB264 Black',
            '050042000001',
            'RSI EDEN XL',
            '8117732100846 BLACK 44']

        for sku in notFoundSku:
            
            try:
                motocrzyElement = zhov.GetInfoMotostyleComUa(priceData[sku])
            except WebSearchNotFound as notFound:
                log.warn('[!] %s' % str(notFound))
                continue
            except WebSearchSkuNotMatched as notMatched:
                log.warn('[!] %s' % str(notMatched))
                continue
                
            

class OFF_testGetFromWebMotoKopylbrosCom(unittest.TestCase):
    
    def runTest(self):
        
        log.info('get info from kop...com test')
        
        testPricePath = os.path.join('prices', 'test', 'kopyl', 'dealer_price_2017-02-04 22.49.00-xls2003.xls')
        
        kop = MXShopKopyl()
        priceData = kop.ReadPrice(testPricePath)

        i = 0
        notFoundCount = 0
        skuNotMatched = 0
        noImages = 0
        
        keysLen = len(priceData.keys())
        
        for sku in priceData.keys():
             
            i += 1
             
            log.debug('%d/%d', i, keysLen)
             
            try:
                element = kop.GetInfoMotoKopylbrosCom(priceData[sku])
            except WebSearchNotFound as ex:
                
                try:
                    element = kop.GetInfoMotoKopylbrosCom(priceData[sku], searchWithName=1)
                except WebSearchNotFound as ex2:
                    ex = ex2
                    
                notFoundCount += 1
                log.warning('[!] %s', (str(ex)))
                raise ex
                
            except WebSearchSkuNotMatched as ex:
                skuNotMatched += 1
                log.warning('[!] %s', (str(ex)))
                raise ex                
             
            self.assertTrue(element)
            if not element['images']:
                noImages += 1
                assert(False)
            
        log.info('web grab for kopyl completed')
        log.info('searched for %d items, sku not matched - %d, not found - %d', i, skuNotMatched, notFoundCount)

class OFF_testMakeUpXmlZhovtuha(unittest.TestCase):
    
    def runTest(self):
        
        xmlFilePath = os.path.join('prices', 'test', 'zhovtuha', 'makeup-test.xls')
        
        zhov = MXShopZhovtuha()
        priceData = zhov.ReadPrice(xmlFilePath)
        
        webData = zhov.GrabWebData(priceData)
        self.assertTrue(webData)
        
        zhov.CreateXmlFile(priceData, webData, '/home/vasya/Dropbox/out.xml')

class OFF_testMakeUpXmlZhovtuha1(unittest.TestCase):
    
    def runTest(self):
        
        zhov = MXShopZhovtuha()
        
        priceData, fileName, currencyRate = zhov.DownloadCurrentPriceFromWeb() 
        
        pp = os.path.join('prices', 'orig', 'zhovtuha', fileName)
        
        FileHlp(pp, 'w').write(priceData)
        
        log.info('file write ok: %s, currency: %f', pp, currencyRate)
                

class testMakeUpXmlKopyl2(unittest.TestCase):
    
    def runTest(self):
        
        zhov = MXShopKopyl()
        
        #priceData, fileName, currencyRate = zhov.DownloadCurrentPriceFromWeb() 
        fileName = 'dealer_price_2017-06-02 14.11.40-2.xls'#'Остатки-23.02.17.xls'
        
        
        pp = os.path.join('prices', 'orig', 'kopyl', fileName)
        
        #FileHlp(pp, 'w').write(priceData)
        
        log.info('file write ok: %s', pp)
        
        priceData = zhov.ReadPrice(pp)
        
        webData = zhov.GrabWebData(priceData)
        self.assertTrue(webData)
        
        outFile = '/media/samba/tmp/out.xml'
        
        zhov.CreateXmlFile(priceData, webData, outFile, 27.5)



class OFF_testMakeUpXmlZhovtuha2(unittest.TestCase):
    
    def runTest(self):
        
        zhov = MXShopZhovtuha()
        
        #priceData, fileName, currencyRate = zhov.DownloadCurrentPriceFromWeb() 
        fileName = 'Остатки-15.03.17.xls'#'Остатки-23.02.17.xls'
        
        
        pp = os.path.join('prices', 'orig', 'zhovtuha', fileName)
        
        #FileHlp(pp, 'w').write(priceData)
        
        log.info('file write ok: %s', pp)
        
        priceData = zhov.ReadPrice(pp)
        
        webData = zhov.GrabWebData(priceData)
        self.assertTrue(webData)
        
        outFile = '/home/vasya/Dropbox/out.xml'
        
        zhov.CreateXmlFile(priceData, webData, outFile, 27.5)
        

class OFF_testZhovtuhaIteractWithAdmin(unittest.TestCase):
    
    def runTest(self):
        
        zhov = MXShopZhovtuha()

        outFile = '/home/vasya/Dropbox/out.xml'
          
        zhov.UploadToServer(outFile, './public_html/newtest/admin/uploads/4.xml')         
        zhov.WebAdminRunPrice()
         
        zhov.DownloadFromServer('./public_html/newtest/admin/uploads/errors.tmp',
                                os.path.join(_CACHE_PATH, 'errors.txt'))
  
   
        zhov.DownloadFromServer('./public_html/newtest/admin/uploads/report.tmp',
                                os.path.join(_CACHE_PATH, 'report.txt'))
         
         
        zhov.AnalyzeErrorsTmp(os.path.join(_CACHE_PATH, 'errors.txt'))
        
        zhov.AnalyzeReportTxt(os.path.join(_CACHE_PATH, 'report.txt'))
        
        
class OFF_testZhovtuhaFixCategories(unittest.TestCase):
    
    def runTest(self):
        
        zhov = MXShopZhovtuha()
        
        zhov.WebAdminFixCategories()
        
class testDoWaterMarkKopyl(unittest.TestCase):
    
    def runTest(self):
        
        log.info('watermarking...')
        
        kop = MXShopKopyl()
        
        kop.AddWaterMarkToAllImages()        
            
class testHttpCache(unittest.TestCase):
    
    def runTest(self):
        
        hc = HttpPageCache('testHttpCache', isClear=True)
        
        hc.put('url', 'data')
        
        self.assertEqual(hc.get('url'), 'data')
        
        isException = False
        try:
            hc.put('url', 'data2')
        except sqlite3.IntegrityError as i:
            isException = True
            
        self.assertTrue(isException)
        
class testXml(unittest.TestCase):
    
    def runTest(self):
        
                # create base
        xml = Xml2003FileStub()
        
        # price format
        xml.addrow(['orig sku', 
                        'category', 
                        'product',
                        'priceSaleUah (%.2f)' % 0,
                        'priceUah', 
                        'dealer usd/uah/salary', 
                        'option', 
                        'options', 
                        'seoUrl + dealerMark', 
                        'description',
                        'balance',
                        'brand'])
        
        xml.addrow(['1', '2', '3', '4'])

        log.info(xml.getdata())
        
        
class OFF_testConvertToXls2003(unittest.TestCase):
    
    def runTest(self):
        
        pathFrom = os.path.join('prices', 'test', 'kopyl', 'dealer_price_2017-02-04 22.49.00.xls')
        pathTo = os.path.join('prices', 'test', 'kopyl', 'dealer_price_2017-02-04 22.49.00-xls2003.xls')
        
        kop = MXShopKopyl()
        
        kop.ConvertTo97Xls(pathFrom, pathTo)
        
        #TODO: try read new file


def ProcessMain(dealer, **kw):
    
    log.info('%s processing...' % dealer._d)        
    
    remoteXmlFile = dealer.WebAdminGetRemoteXmlName()
            
    data, fileName, currencyRate = dealer.DownloadCurrentPriceFromWeb()
    
    log.info('price downloaded: %s', fileName)
     
    xlsFilePath = os.path.join(dealer._pricesOrigDir, fileName)
    xmlFilePath = os.path.join(dealer._pricesResutDir, os.path.splitext(fileName)[0] + '.xml')
     
    newFilePath = os.path.join(_CACHE_PATH, os.path.splitext(fileName)[0] + '.xml')
             
    FileHlp(xlsFilePath, 'w').write(data)
    
    cachedXml = kw.get('cachedXml', None)
    if not cachedXml:        
     
        r = dealer.ConvertTo97Xls(xlsFilePath, newFilePath)
        assert(r)
         
        priceData = dealer.ReadPrice(newFilePath)
        assert(priceData)
        os.remove(newFilePath)
          
        webData = dealer.GrabWebData(priceData)
        assert(webData)
          
        dealer.CreateXmlFile(priceData, webData, xmlFilePath, currencyRate)
        
    else:
        log.info("using cached xml: %s", cachedXml)
        xmlFilePath = cachedXml
    
    if kw.get('noUploadToAdmin', None):
        log.info('[-] no upload option specified, exiting now')
        return  
    
    dealer.UploadToServer(xmlFilePath, '%s/%s' % (dealer._remoteUploadDir, remoteXmlFile),
                          addDockerPrefix=True)
       
    for idx in range(0, 9):
           
        try:  
            dealer.WebAdminRunPrice()
            break
        except AdminNeedContinue:                
            log.info('not all data processed, make another request...')
               
            idx += 1
            continue
        
    dealer.DownloadFromServer('%s/errors.tmp' % dealer._remoteUploadDir,
                            os.path.join(_CACHE_PATH, 'errors.txt'),
                            addDockerPrefix=True)
    
    dealer.DownloadFromServer('%s/report.tmp' % dealer._remoteUploadDir,
                            os.path.join(_CACHE_PATH, 'report.txt'),
                            addDockerPrefix=True)
       
    dealer.AnalyzeErrorsTmp(os.path.join(_CACHE_PATH, 'errors.txt'))        
    dealer.AnalyzeReportTxt(os.path.join(_CACHE_PATH, 'report.txt'))



if __name__ == "__main__":
    

    parser = OptionParser(usage="usage: %prog [options]", version="%prog " + _VERSION_)    
    
    parser.add_option("-v", "--verbose",
                      action="store_true", dest="verbose", default=False,
                      help="Show debug messages")
    
    parser.add_option("-a", "--pricesAll",
                      action="store_true", dest="pricesAll", default=False,
                      help="Build all prices")

    parser.add_option("-p", "--prices=USER1,USER2,...", type="string",
                      action="store", dest="buildPrices", default="none",
                      help="Build prices for specified user. Use `-p all` to build all")

    parser.add_option("-t", "--useTestingServer",
                      action="store_true", dest="useTestingServer", default=False,
                      help="Use testing server instead of production")

    parser.add_option("-n", "--noUploadToAdmin",
                      action="store_true", dest="noUploadToAdmin", default=False,
                      help="Do not iteract with admin panel")

    parser.add_option("-c", "--cachePath", type="string",
                      action="store", dest="cachePath", default='',
                      help="change default program cache path")
    
    parser.add_option("-x", "--cachedXml", type="string",
                      action="store", dest="cachedXml", default='',
                      help="use cached xml instead of generating new")

# TODO:
#     parser.add_option("-w", "--watermark=USER1,USER2,...", type="string",
#                       action="store", dest="watermark", default=False,
#                       help="Build prices for specified user. Use -wa to buid")
#         
#     parser.add_option("-l", "--watermarksAll",
#                       action="store_true", dest="watermarksAll", default=False,
#                       help="Make watermark for all users")
    

    (options, args) = parser.parse_args()
    
    if options.cachePath:
        _CACHE_PATH = options.cachePath
        
    initCacheFolder(_CACHE_PATH)
    
    if options.buildPrices == 'none':
        parser.print_help()
        exit(1)
        
    if options.verbose:
        console.setLevel(logging.DEBUG)   
                 
    if options.buildPrices:
        
        if 'zhov' in options.buildPrices.split(','):
            dealer = MXShopZhovtuha(useTestingServer=options.useTestingServer)
            ProcessMain(dealer, noUploadToAdmin=options.noUploadToAdmin,
                        cachedXml=options.cachedXml)
            
            del(dealer)

        if 'kop' in options.buildPrices.split(','):
            
            dealer = MXShopKopyl(useTestingServer=options.useTestingServer)
            
            ProcessMain(dealer, noUploadToAdmin=options.noUploadToAdmin,
                         cachedXml=options.cachedXml)
            
            if not options.noUploadToAdmin:
            
                dealer.AddWaterMarkToAllImages()
                        
            del(dealer)

        
    
 
