#!/usr/bin/env python
# encoding: utf8

import re
import json
import cjson
import os.path
from random import random
from urlparse import urlparse as parse_url
from hashlib import sha1
from urllib import quote
from collections import defaultdict, namedtuple, Counter
from subprocess import check_call
from random import sample
from datetime import datetime
import xml.etree.ElementTree as ET
from math import ceil

import pandas as pd
from bs4 import BeautifulSoup

import seaborn as sns
import matplotlib as mpl
from matplotlib import pyplot as plt
from matplotlib import rc
# For cyrillic labels
rc('font', family='Verdana', weight='normal')

import requests
requests.packages.urllib3.disable_warnings()

from jinja2 import Environment, Template

import pyphen
HYPHEN = pyphen.Pyphen(lang='ru_RU')

import site; site.addsitedir("/usr/local/lib/python2.7/site-packages")
import cv2


HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/37.0.2049.0 Safari/537.36'
}
DATA_DIR = 'data'
HTML_DIR = os.path.join(DATA_DIR, 'html')
HTML_LIST = os.path.join(HTML_DIR, 'list.txt')

CHECKS_DIR = os.path.join(DATA_DIR, 'check')

EDUOFFICE_IDS_URL = 'http://map.development.mskobr.ru/api/eduoffice/points.json'
EDUOFFICE_DIR = os.path.join(DATA_DIR, 'data.mskobr', 'accounting')
EDUOFFICE_REPORT_DIR = os.path.join(DATA_DIR, 'data.mskobr', 'export')
EDUOFFICES_CHECK = os.path.join(CHECKS_DIR, 'eduoffices.xlsx')
SAD_CHECKS = os.path.join(CHECKS_DIR, 'sads.xlsx')

MSKOBR_DIR = os.path.join(DATA_DIR, 'mskobr')
MSKOBR_URL_MENUS = os.path.join(MSKOBR_DIR, 'url_menus.json')
MSKOBR_URL_STAFF = os.path.join(MSKOBR_DIR, 'url_staff.json')
MSKOBR_SAD_URL_STAFF = os.path.join(MSKOBR_DIR, 'sad_url_staff.json')
MSKOBR_URL_TEACHER_URLS = os.path.join(MSKOBR_DIR, 'url_teacher_urls.json')
MSKOBR_TEACHERS = os.path.join(MSKOBR_DIR, 'teachers.json')
MSKOBR_SAD_TEACHERS = os.path.join(MSKOBR_DIR, 'sad_teachers.json')
MSKOBR_REVIEWS = os.path.join(MSKOBR_DIR, 'reviews.json')

PODGON = os.path.join(DATA_DIR, 'podgon.xlsx')
PODGON_HOST_ALIASES = {
    'couz117.mskobr.ru': 'sch117.mskobr.ru',
    'lyc1568sv-new.mskobr.ru': 'lyc1568.mskobr.ru',
    'sch519uv.mskobr.ru': 'spectr.mskobr.ru',
    'sch1847s.mskobr.ru': 'itschool.mskobr.ru',
    'gym1599uv.mskobr.ru': 'lgkuv.mskobr.ru'
}

ADDRESSES_DIR = os.path.join(DATA_DIR, 'geocode')
ADDRESSES_LIST = os.path.join(ADDRESSES_DIR, 'list.txt')
ADDRESSES_CHECK = os.path.join(CHECKS_DIR, 'addresses.xlsx')
SAD_ADDRESSES_CHECK = os.path.join(CHECKS_DIR, 'sad_addresses.xlsx')

MSKOBR_URL_PENNANTS = os.path.join(MSKOBR_DIR, 'url_pennants.json')
RATING_2015_URLS = [
    'http://dogm.mos.ru/rating/',
    'http://dogm.mos.ru/napdeyat/obdet/ranking-301-to-500.php'
]
RATING_2014_URL = 'http://dogm.mos.ru/rating/r2013_2014.php'
RATING_2013_URL = 'http://dogm.mos.ru/a_news/28.07.2014.php'
RATINGS_CHECK = os.path.join(CHECKS_DIR, 'ratings.xlsx')

EGE = os.path.join(DATA_DIR, 'ege.xlsx')
EGE_TITLE_HOST_CHECK = os.path.join(DATA_DIR, 'ege_title_host_check.xlsx')
EGE_DISTRIBUTIONS_CHECK = os.path.join(CHECKS_DIR, 'ege.xlsx')
OLYMPIADS_CHECK = os.path.join(CHECKS_DIR, 'olympiads.xlsx')

# https://oauth.vk.com/authorize?client_id=5006136&redirect_uri=https://oauth.vk.com/blank.html&display=page&response_type=token 
VK_TOKEN = 'c81d246133949a8514d52b21b7c9c9804f07f0f457ad687228edf41c5ea0152d536948f376ea5d58409cb'
VK_DIR = os.path.join(DATA_DIR, 'vk')
VK_SCHOOL_IDS = os.path.join(VK_DIR, 'school_ids.xlsx')
VK_PUPILS_DIR = os.path.join(VK_DIR, 'pupils')
URL_UNIVERSITIES = os.path.join(VK_DIR, 'url_universities.json')
UNIVERSITIES_CHECK = os.path.join(CHECKS_DIR, 'universities.xlsx')

CHECK_CONTACTS = os.path.join(CHECKS_DIR, 'contacts.xlsx')
CHECK_TEACHERS = os.path.join(CHECKS_DIR, 'teachers.xlsx')
TEACHERS_STAFF_CHECK = os.path.join(CHECKS_DIR, 'teachers_staff.xlsx')
SAD_TEACHERS_CHECK = os.path.join(CHECKS_DIR, 'sad_teachers.xlsx')

GALLERY_IMAGES = os.path.join(DATA_DIR, 'gallery_images.json')
SAD_GALLERY_IMAGES = os.path.join(DATA_DIR, 'sad_gallery_images.json')
CHECK_IMAGES_DIR = os.path.join(CHECKS_DIR, 'images')
CHECK_IMAGES = os.path.join(CHECKS_DIR, 'images.xlsx')
CHECK_SAD_IMAGES = os.path.join(CHECKS_DIR, 'sad_images.xlsx')

VIZ_DIR = 'viz'
SCHOOL_TEMPLATE = os.path.join(VIZ_DIR, 'school.html')
INDEX_TEMPLATE = os.path.join(VIZ_DIR, 'index.html')
LIST_TEMPLATE = os.path.join(VIZ_DIR, 'list.html')
SAD_TEMPLATE = os.path.join(VIZ_DIR, 'sad.html')
SITE_DIR = os.path.join(VIZ_DIR, 'site')
INDEX = os.path.join(SITE_DIR, 'index.html')
LIST = os.path.join(SITE_DIR, 'list.html')

IMAGES_DIR = os.path.join(SITE_DIR, 'i')
RAW_IMAGES_DIR = os.path.join(DATA_DIR, 'images')
THUMB_IMAGE_DIR = os.path.join(IMAGES_DIR, 'thumb')

SCHOOLOTZYV_REGIONS = [
    u'ЦАО',
    u'СЗАО',
    u'САО',
    u'СВАО',
    u'ВАО',
    u'ЮВАО',
    u'ЮЗАО',
    u'ЮАО',
    u'ЗАО'
]
SCHOOLOTZYV_REVIEWS = os.path.join(DATA_DIR, 'schoolotzyv_reviews.json')
SCHOOLOTZYV_SITE_ALIASES = {
    'gym1579u.mskobr.ru': 'http://www.1579.ru',
    'sch1190sz.mskobr.ru': 'http://www.schoolotzyv.ru/schools/5-moskva/25-srednie/1465-shkola-1190',
    'gym1522.mskobr.ru': 'http://www.schoolotzyv.ru/schools/5-moskva/23-gimnazii/1642-gimnaziya-1522',
    'sch1387.mskobr.ru': 'http://schu1387.mskobr.ru',
    'sch368.mskobr.ru': 'http://schv368.mskobr.ru/',
    'sch1280uz.mskobr.ru': 'http://www.schoolotzyv.ru/schools/5-moskva/25-srednie/1556-shkola-1280',
    'gym1409s-new.mskobr.ru': 'http://www.schoolotzyv.ru/schools/5-moskva/23-gimnazii/1620-gimnaziya-1409',
    'schv1947.mskobr.ru': 'http://www.schoolotzyv.ru/schools/5-moskva/25-srednie/1759-shkola-1947',
    'sch547u.mskobr.ru': 'http://www.schoolotzyv.ru/schools/5-moskva/25-srednie/930-shkola-547',
    'sch597s.mskobr.ru': 'http://schs597.mskobr.ru/',
    'sch484uv-new.mskobr.ru': 'http://sch484.uvao.mskobr.ru/',
    'sch1592.mskobr.ru': 'http://gyms1592.mskobr.ru',
    'sch1413sv.mskobr.ru': 'http://www.schoolotzyv.ru/schools/5-moskva/25-srednie/1624-shkola-1413',
    'sch2006uz.mskobr.ru': 'http://school2006.msk.ru/',
    'sch2055c.mskobr.ru': 'http://www.2055shkola.ru/',
    'sch1874sz.mskobr.ru': 'http://www.schoolotzyv.ru/schools/5-moskva/25-srednie/1719-czentr-obrazovaniya-1874',
    'sch2017u.mskobr.ru': 'http://www.sch2017.ru/',
    'sch1430sv-new.mskobr.ru': 'http://sch1430.lianet.ru/',
    'sch1246v.mskobr.ru': 'http://schools.keldysh.ru/sch1246/',
    'sch1909uv.mskobr.ru': 'http://www.schoolotzyv.ru/schools/5-moskva/25-srednie/1730-shkola-1909',
    'schu1929.mskobr.ru': 'http://www.sch1929.edusite.ru/',
    'lyc507u.mskobr.ru': 'http://licey507.ru',
    'gym1543.mskobr.ru': 'http://www.1543.ru',
    'sch1210sz.mskobr.ru': 'http://sch1210sz.narod.ru/',
    'sch1256uv.mskobr.ru': 'http://www.sch1256.uvuo.ru/',
    'sch1265uz.mskobr.ru': 'http://www.sch1265.ru/',
    'sch1512.mskobr.ru': 'http://1512.ru/',
    'sch1245u.mskobr.ru': 'http://sch1245.com.ru/',
    'schu1302.mskobr.ru': 'http://www.school-1302.ru/',
    'sch1078.mskobr.ru': 'http://schools.keldysh.ru/sch1078/index.html',
    'sch1307.mskobr.ru': 'http://schu1307.mskzapad.ru/',
    'sch1252.mskobr.ru': 'http://www.schoolotzyv.ru/schools/5-moskva/25-srednie/1527-shkola-1252',
    'schv362.mskobr.ru': 'http://www.schoolotzyv.ru/schools/5-moskva/25-srednie/760-shkola-362',
    'cos686.mskobr.ru': 'http://www.schoolotzyv.ru/schools/5-moskva/25-srednie/2045-shkola-686',
    'sch2012uv.mskobr.ru': 'http://www.schoolotzyv.ru/schools/5-moskva/25-srednie/1809-shkola-2012',
    'sch1284.mskobr.ru': 'http://xn---1284-3ve3ei0a2h.xn--p1ai/',
    'co1679s.mskobr.ru': 'http://www.schoolotzyv.ru/schools/5-moskva/25-srednie/1672-shkola-1679',
    'gym1507uz.mskobr.ru': 'http://www.gym1507.mosuzedu.ru /',
    'sch633v.mskobr.ru': 'http://school633.mhost.ru/',
    'sch1935uv.mskobr.ru': 'http://1935.moy.su/',
    'lycc1589.mskobr.ru': 'http://www.schoolotzyv.ru/schools/5-moskva/24-licei/1825-licey-1589',
    'sch402.mskobr.ru': 'http://www.s402.edusite.ru/',
    'sch875.mskobr.ru': 'http://www.schoolotzyv.ru/schools/5-moskva/25-srednie/1209-shkola-875',
    'sch668s.mskobr.ru': 'http://school668.sk6.ru',
    'sch1212.mskobr.ru': 'http://sch1212sz.mskobr.ru',
    'gym201s.mskobr.ru': 'http://www.gimnazia-201.ru/',
    'sch1429c.mskobr.ru': 'http://school1429.ru',
    'sch1432.mskobr.ru': 'http://co1432.mskzapad.ru/',
    'sch975u.mskobr.ru': 'http://www.schoolotzyv.ru/schools/5-moskva/25-srednie/1303-shkola-975',
    'sch1362.mskobr.ru': 'http://school1362.ru',
    'sch1251s.mskobr.ru': 'http://www.school1251.ru/',
    'sch587.mskobr.ru': 'http://www.schoolotzyv.ru/schools/5-moskva/25-srednie/963-shkola-587',
    'sch1368uz.mskobr.ru': 'http://www.schoolotzyv.ru/schools/5-moskva/25-srednie/1827-shkola-1368',
}

SEARCH_HTML = os.path.join(HTML_DIR, 'mel_search.html')
MEL_REVIEWS = os.path.join(DATA_DIR, 'mel_reviews.json')

REVIEWS_CHECK = os.path.join(CHECKS_DIR, 'reviews.xlsx')
REVIEW_MONTHS = {
    1: u'январь',
    2: u'февраль',
    3: u'марnт',
    4: u'апрель',
    5: u'май',
    6: u'июнь',
    7: u'июль',
    8: u'август',
    9: u'сентябрь',
    10: u'октябрь',
    11: u'ноябрь',
    12: u'декабрь',
}
MSKOBR_REVIEW_MONTHS = {
    u'января': 1,
    u'февраля': 2,
    u'марта': 3,
    u'апреля': 4,
    u'мая': 5,
    u'июня': 6,
    u'июля': 7,
    u'августа': 8,
    u'сентября': 9,
    u'октября': 10,
    u'ноября': 11,
    u'декабря': 12,
}
MAIL_LOGIN = 'ak@obr.msk.ru'
MAIL_PASSWORD = os.environ.get('AK_MAIL_PASSWORD')

EMAIL_SOURCE = 'email'
URL_SOURCE = 'url'

EDUOFFICE_REPORT_CHECK = os.path.join(CHECKS_DIR, 'eduoffice_reports.xlsx')

BUS_DIR = os.path.join(DATA_DIR, 'bus')
BUS_SEARCH_DIR = os.path.join(BUS_DIR, 'search')
BUS_REPORT_DIR = os.path.join(BUS_DIR, 'report')

BIN_DIR = 'bin'
TOMITA_BIN = os.path.join(BIN_DIR, 'tomita-mac')
ALGFIO_DIR = os.path.join(BIN_DIR, 'algfio')
ALGFIO_TEXT = os.path.join(ALGFIO_DIR, 'text.txt')
ALGFIO_FACTS = os.path.join(ALGFIO_DIR, 'facts.xml')
ALGFIO_CONFIG = os.path.join(ALGFIO_DIR, 'config.proto')
TOMITA_DATA_DIR = os.path.join(DATA_DIR, 'tomita')
ALGFIO_DATA_DIR = TOMITA_DATA_DIR
ALGFIO_LIST = os.path.join(ALGFIO_DATA_DIR, 'list.txt')

CHECK_SAD_ALGFIO_REVIEWS = os.path.join(CHECKS_DIR, 'sad_algfio_reviews.xml')

PHOTO_DIR = os.path.join(IMAGES_DIR, 'photo')

HAAR_CASCADE = os.path.join(DATA_DIR, 'haarcascade_frontalface_default.xml')


Coordinates = namedtuple('Coordinates', ['latitude', 'longitude'])
Address = namedtuple('Address', ['description', 'coordinates'])
Contacts = namedtuple(
    'Contacts',
    ['name', 'mobile', 'work', 'email']
)
Eduoffice = namedtuple(
    'Eduoffice',
    ['id', 'inn', 'state', 'founder', 'url',
     'title', 'full_title', 'programs',
     'main_address', 'other_addresses',
     'director']
)
CheckEduofficeRecord = namedtuple(
    'CheckEduofficeRecord',
    ['url', 'title', 'short', 'no']
)
CheckAddressRecord = namedtuple(
    'CheckAddressRecord',
    ['program', 'address', 'url']
)
CheckRatingRecord = namedtuple(
    'CheckRatingRecord',
    ['url', 'year_2013', 'year_2014', 'year_2015']
)
EgeRecord = namedtuple(
    'EgeRecord',
    ['title', 'short', 'result']
)
EgeTitleHostRecord = namedtuple(
    'EgeTitleHostRecord',
    ['title', 'host']
)
CheckEgeRecord = namedtuple(
    'CheckEgeRecord',
    ['host', 'less_150', 'between_150_220', 'more_220', 'a', 'b', 'source']
)
CheckOlympiadsRecord = namedtuple(
    'CheckOlympiadsRecord',
    ['host', 'subject', 'olympiad', 'place', 'count']
)
CheckContactsRecord = namedtuple(
    'CheckContactsRecord',
    ['url', 'position', 'name', 'work', 'mobile', 'email']
)
CheckUniversitiesRecord = namedtuple(
    'CheckUniversitiesRecord',
    ['url', 'university', 'count', 'source']
)

TEACHER_SUBJECTS_ORDER = [
    u'Русский и литература',
    u'Математика',
    u'История и обществознание',
    u'Физика',
    u'Информатика',
    u'Английский',
    u'Иностранный язык',
    u'Биология',
    u'Химия',
    u'География',
    u'ИЗО',
    u'ОБЖ',
    u'Физкультура',
    u'Технология',
    u'Музыка',
    u'Начальные классы'
]
CheckTeachersRecord = namedtuple(
    'CheckTeachersRecord',
    ['url', 'subject', 'name', 'category']
)
CheckTeachersStaffRecord = namedtuple(
    'CheckTeachersStaffRecord',
    ['url', 'name', 'position']
)
CheckSadTeachersRecord = namedtuple(
    'CheckSadTeachersRecord',
    ['url', 'name', 'position', 'image']
)

MskobrMenuItem = namedtuple(
    'MskobrMenuItem',
    ['program', 'title', 'url', 'address']
)
MskobrTeacherRecord = namedtuple(
    'MskobrTeacherRecord',
    ['url', 'name', 'image', 'position', 'discipline', 'level',
     'experience', 'total_experience', 'university', 'category']
)
GiaResult = namedtuple('GiaResult', ['pupils', 'grade'])
EgeResults = namedtuple(
    'EgeResult',
    ['total', 'more_150', 'more_220']
)
OlympiadRecord = namedtuple(
    'OlympiadRecord',
    ['type', 'level', 'class_', 'place']
)
PodgonRecord = namedtuple(
    'PodgonRecord',
    ['id', 'inn', 'name', 'number', 'host', 'gia', 'olympiads', 'ege']
)
EgeIntervalsRecord = namedtuple(
    'EgeIntervalsRecord',
    ['host', 'less_150', 'between_150_220', 'more_220']
)
Size = namedtuple('Size', ['width', 'height'])
Image = namedtuple('Image', ['url', 'filename', 'raw', 'thumb'])

VizSchoolTitles = namedtuple(
    'VizSchoolTitles',
    ['full', 'short', 'label', 'link', 'no']
)
VizSchoolRating = namedtuple(
    'VizSchoolRating',
    ['year_2013', 'year_2014', 'year_2015']
)

PROGRAM_0_1 = 'program_0_1'
PROGRAM_1_4 = 'program_1_4'
PROGRAM_1_7 = 'program_1_7'
PROGRAM_1_9 = 'program_1_9'
PROGRAM_1_11 = 'program_1_11'
PROGRAM_5_11 = 'program_5_11'
PROGRAM_6_11 = 'program_6_11'
PROGRAM_8_11 = 'program_8_11'
PROGRAM_10_11 = 'program_10_11'
VizSchoolAdress = namedtuple(
    'VizSchoolAdress',
    ['address', 'latitude', 'longitude', 'program']
)

VizSchoolEge = namedtuple(
    'VizSchoolEge',
    ['less_150', 'between_150_220', 'more_220', 'a', 'b', 'source']
)

OlimpiadaSerpRecord = namedtuple(
    'OlympiadaSerpRecord',
    ['login', 'type', 'name1', 'number', 'name2',
     'name3', 'name4', 'region1', 'region2', 'region3', 'id']
)
OlimpiadaSerp = namedtuple(
    'OlimpiadaSerp',
    ['query', 'results']
)
OlimpiadaLoginRecord = namedtuple(
    'OlimpiadaLoginRecord',
    ['host', 'login']
)
OlimpiadaResultRecord = namedtuple(
    'OlimpiadaResultRecord',
    ['host', 'surname', 'name', 'middle', 'year', 'year2',
     'place', 'type', 'subject', 'year3', 'stage']
)

OLIMPIADA_LOGINS = os.path.join(DATA_DIR, 'olimpiada_logins.xlsx')

OLYMPIAD_MOSCOW_1 = 'olympiad_moscow_1'
OLYMPIAD_MOSCOW_2 = 'olympiad_moscow_2'
OLYMPIAD_RUSSIA_1 = 'olympiad_russia_1'
OLYMPIAD_RUSSIA_2 = 'olympiad_russia_2'
OLYMPIAD_TYPES_ORDER = [
    OLYMPIAD_RUSSIA_1,
    OLYMPIAD_RUSSIA_2,
    OLYMPIAD_MOSCOW_1,
    OLYMPIAD_MOSCOW_2
]
OLYMPIAD_ALIASES = {
    u'Изобразительное искусство': u'ИЗО',
    u'Мировая художественная культура': u'МХК',
    u'Основы безопасности жизнедеятельности': u'ОБЖ'
}
OLYMPIADS_ORDER = {
    u'Русский',
    u'Математика',
    u'Литература',
    u'Обществознание',
    u'Филология',
    u'Право',
    u'История',
    u'Физика',
    u'Информатика',
    u'Астрономия',
    u'Экономика',
    u'Английский',
    u'Биология',
    u'Химия',
    u'МХК',
    u'ИЗО',
    u'Физкультура',
    u'ОБЖ',
    u'Технология',
    u'География',
    u'Экология',
    u'Французский',
    u'Немецкий',
    u'Лингвистика',
    u'Испанский',
    u'Латинский язык',
}
VizSchoolOlympiads = namedtuple(
    'VizSchoolOlympiads',
    ['subject', 'type', 'count']
)

VizSchoolUniversity = namedtuple(
    'VizSchoolUniversity',
    ['name', 'ids', 'source']
)
VizSchoolContact = namedtuple(
    'VizSchoolContact',
    ['position', 'name', 'work', 'mobile', 'email']
)

CATEGORY_TOP = 'high'
CATEGORY_1 = '1'
CATEGORY_NO = 'no'
CATEGORY_UNKNOWN = 'unknown'
CATEGORIES_ORDER = [
    CATEGORY_TOP,
    CATEGORY_1,
    CATEGORY_NO,
    CATEGORY_UNKNOWN
]
VizTeacher = namedtuple(
    'VizTeacher',
    ['subject', 'url', 'category']
)

VizSchool = namedtuple(
    'VizSchool',
    ['url', 'sad_link', 'title', 'rating', 'programs', 'addresses',
     'ege', 'olympiads', 'universities', 'contacts',
     'teachers', 'images', 'reviews', 'polls', 'eduoffice_report']
)

VkUniversity = namedtuple(
    'VkUniversity',
    ['id', 'name', 'faculty', 'form', 'status', 'year']
)
VkPupil = namedtuple(
    'Pupil',
    ['id', 'name', 'universities']
)

VkPollAnswer = namedtuple(
    'VkPollAnswer',
    ['id', 'text']
)
VkPoll = namedtuple(
    'VkPoll',
    ['id', 'owner', 'question', 'answers']
)
VkPollVote = namedtuple(
    'VkPollVote',
    ['answer', 'user']
)
VkPollAnswerRecord = namedtuple(
    'VkPollAnswerRecord',
    ['text', 'votes']
)
VkPollStatsRecord = namedtuple(
    'VkPollStatsRecord',
    ['host', 'question', 'answers']
)

SchoolotzyvSchoolRecord = namedtuple(
    'SchoolotzyvSchoolRecord',
    ['url', 'site', 'reviews', 'merge']
)
SchoolotzyvVotes = namedtuple(
    'SchoolotzyvVotes',
    ['truth', 'lie']
)
SchoolotzyvReviewRecord = namedtuple(
    'SchoolotzyvReviewRecord',
    ['url', 'name', 'date', 'votes', 'text']
)
SchoolotzyvReviewsRecord = namedtuple(
    'SchoolotzyvReviewsRecord',
    ['urls', 'site', 'reviews']
)

MelVotesRecord = namedtuple(
    'MelVotesRecord',
    ['education', 'teachers', 'atmosphere', 'infrastructure']
)
MelReviewRecord = namedtuple(
    'MelReviewRecord',
    ['author', 'votes', 'text']
)
MelReviewsRecords = namedtuple(
    'MelReviewsRecord',
    ['url', 'site', 'reviews']
)
CheckReviewRecord = namedtuple(
    'CheckReviewRecord',
    ['host', 'url', 'date', 'name', 'text']
)

FeedbackPart = namedtuple(
    'FeedbackPart',
    ['name', 'text']
)
FeedbackRecord = namedtuple(
    'FeedbackRecord',
    ['id', 'date', 'type', 'sections']
)

EduofficeReportRecord = namedtuple(
    'EduofficeReportRecord',
    ['id', 'pupils', 'incoming', 'teachers', 'salaries']
)
EduofficeReportPupils = namedtuple(
    'EduofficeReportPupils',
    ['total', 'program_0', 'program_1_4', 'program_5_9', 'program_10_11']
)
EduofficeReportIncoming = namedtuple(
    'EduofficeReportIncoming',
    ['total', 'from_0']
)
EduofficeReportTeachers = namedtuple(
    'EduofficeReportTeachers',
    ['total', 'main', 'other_main', 'administration', 'other']
)
EduofficeReportSalaries = namedtuple(
    'EduofficeReportSalaries',
    ['total', 'teachers', 'administration', 'other']
)
EduofficeReportCheckRecord = namedtuple(
    'EduofficeReportCheckRecord',
    ['host', 'pupils', 'teachers', 'salaries', 'incoming']
)

BusSearchRecord = namedtuple(
    'BusSearchRecord',
    ['inn', 'id', 'url', 'name', 'other']
)
BusReportYears = namedtuple(
    'BusReportYears',
    ['id', 'year_2015']
)
BusReportRecord = namedtuple(
    'BusReportRecord',
    ['id', 'incomings', 'expenses']
)
BusReportIncomingsRecord = namedtuple(
    'BusReportIncomingsRecord',
    ['total', 'paid', 'subsidii']
)
BusReportExpensesRecord = namedtuple(
    'BusReportExpensesRecord',
    ['total', 'salaries', 'gkh']
)

MskobrReviewRecord = namedtuple(
    'MskobrReviewRecord',
    ['date', 'author', 'text']
)
MskobrReviewsRecord = namedtuple(
    'MskobrReviewsRecord',
    ['url', 'reviews']
)

Name = namedtuple(
    'Name',
    ['last', 'first', 'middle']
)

AlgfioTomitaFact = namedtuple(
    'AlgfioTomitaFact',
    ['start', 'size', 'substring', 'last', 'first', 'middle', 'known_surname']
)
TeacherMention = namedtuple(
    'TeacherMention',
    ['start', 'size', 'substring', 'name', 'teacher']
)
AlgfioFact = namedtuple(
    'AlgfioFact',
    ['start', 'size', 'name']
)
AlgfioReviewCheckRecord = namedtuple(
    'AlgfioReviewCheckRecord',
    ['url', 'author', 'date', 'text', 'facts']
)

SadCheckRecord = namedtuple(
    'SadCheckRecord',
    ['url', 'name']
)
SadTeacherMention = namedtuple(
    'SadTeacherMention',
    ['start', 'size', 'teacher', 'exact', 'sad']
)

VizSadAddress = namedtuple(
    'VizSadAddress',
    ['address', 'latitude', 'longitude']
)
VizSadTitle = namedtuple(
    'VizSadTitle',
    ['full', 'link']
)
VizTeacherMentions = namedtuple(
    'VizTeacherMentions',
    ['url', 'name', 'image', 'position', 'mentions', 'sample']
)
VizSadReview = namedtuple(
    'VizSadReview',
    ['id', 'author', 'date', 'html']
)
VizSad = namedtuple(
    'VizSad',
    ['url', 'title', 'school', 'addresses', 'contacts',
     'teacher_mentions', 'reviews', 'images', 'incoming']
)


def log_progress(sequence, every=None, size=None):
    from ipywidgets import IntProgress, HTML, VBox
    from IPython.display import display

    is_iterator = False
    if size is None:
        try:
            size = len(sequence)
        except TypeError:
            is_iterator = True
    if size is not None:
        if every is None:
            if size <= 200:
                every = 1
            else:
                every = size / 200     # every 0.5%
    else:
        assert every is not None, 'sequence is iterator, set every'

    if is_iterator:
        progress = IntProgress(min=0, max=1, value=1)
        progress.bar_style = 'info'
    else:
        progress = IntProgress(min=0, max=size, value=0)
    label = HTML()
    box = VBox(children=[label, progress])
    display(box)

    index = 0
    try:
        for index, record in enumerate(sequence, 1):
            if index == 1 or index % every == 0:
                if is_iterator:
                    label.value = '{index} / ?'.format(index=index)
                else:
                    progress.value = index
                    label.value = u'{index} / {size}'.format(
                        index=index,
                        size=size
                    )
            yield record
    except:
        progress.bar_style = 'danger'
        raise
    else:
        progress.bar_style = 'success'
        progress.value = index
        label.value = str(index or '?')


def jobs_manager():
    from IPython.lib.backgroundjobs import BackgroundJobManager
    from IPython.core.magic import register_line_magic
    from IPython import get_ipython
    
    jobs = BackgroundJobManager()

    @register_line_magic
    def job(line):
        ip = get_ipython()
        jobs.new(line, ip.user_global_ns)

    return jobs


def kill_thread(thread):
    import ctypes
    
    id = thread.ident
    code = ctypes.pythonapi.PyThreadState_SetAsyncExc(
        ctypes.c_long(id),
        ctypes.py_object(SystemError)
    )
    if code == 0:
        raise ValueError('invalid thread id')
    elif code != 1:
        ctypes.pythonapi.PyThreadState_SetAsyncExc(
            ctypes.c_long(id),
            ctypes.c_long(0)
        )
        raise SystemError('PyThreadState_SetAsyncExc failed')


def get_chunks(sequence, count):
    count = min(count, len(sequence))
    chunks = [[] for _ in range(count)]
    for index, item in enumerate(sequence):
        chunks[index % count].append(item) 
    return chunks


def hash_item(item):
    return sha1(item.encode('utf8')).hexdigest()


hash_url = hash_item
hash_address = hash_item


def get_html_filename(url):
    return '{hash}.html'.format(
        hash=hash_url(url)
    )


def get_html_path(url):
    return os.path.join(
        HTML_DIR,
        get_html_filename(url)
    )


def load_items_cache(path):
    with open(path) as file:
        for line in file:
            line = line.decode('utf8').rstrip('\n')
            hash, item = line.split('\t', 1)
            yield item


def list_html_cache():
    return load_items_cache(HTML_LIST)


def update_items_cache(item, path):
    with open(path, 'a') as file:
        hash = hash_item(item)
        file.write('{hash}\t{item}\n'.format(
            hash=hash,
            item=item.encode('utf8')
        ))
        

def update_html_cache(url):
    update_items_cache(url, HTML_LIST)


def dump_text(data, path):
    with open(path, 'w') as file:
        file.write(data.encode('utf8'))


def dump_html(url, html):
    path = get_html_path(url)
    if html is None:
        html = ''
    dump_text(html, path)
    update_html_cache(url)


def load_text(path):
    with open(path) as file:
        return file.read().decode('utf8')


def load_html(url):
    path = get_html_path(url)
    return load_text(path)


def download_url(url):
    try:
        response = requests.get(
            url,
            headers=HEADERS,
            timeout=5
        )
        return response.text
    except requests.RequestException:
        return None


def fetch_url(url):
    html = download_url(url)
    dump_html(url, html)


def fetch_urls(urls):
    for url in urls:
        fetch_url(url)


def parse_eduoffice_ids(data):
    for item in data.split('|'):
        id, _ = item.split(';', 1)
        yield int(id)


def download_json(url):
    try:
        response = requests.get(
            url,
            headers=HEADERS,
            timeout=5,
        )
        return response.json()
    except requests.RequestException:
        return None


def get_eduoffice_url(id):
    return 'http://map.development.mskobr.ru/api/eduoffices/{id}.json'.format(
        id=id
    )


def get_eduoffice_filename(id):
    return '{id}.json'.format(id=id)


def get_eduoffice_path(id):
    return os.path.join(
        EDUOFFICE_DIR,
        get_eduoffice_filename(id)
    )


def parse_eduoffice_filename(filename):
    id, _ = filename.split('.', 1)
    return int(id)


def list_eduoffice_cache():
    for filename in os.listdir(EDUOFFICE_DIR):
        yield parse_eduoffice_filename(filename)


def load_json(path):
    with open(path) as file:
        return cjson.decode(file.read())


def dump_json(data, path):
    with open(path, 'w') as file:
        json.dump(data, file)


def dump_eduoffice(data, id):
    path = get_eduoffice_path(id)
    dump_json(data, path)


def load_raw_eduoffice(id):
    path = get_eduoffice_path(id)
    return load_json(path)


def parse_eduoffice(data):
    data = data['list']
    id = data['eo_id']
    inn = data['inn']
    if not inn:  # 0 or ''
        inn = None

    state = data['reorganization']['name']

    founder = None
    if 'founder' in data:
        founder = data['founder']['founder']

    url = data['website']
    if url:
        url = 'http://' + url
    else:
        url = None

    title = data['title']
    full_title = data['title_full']
    programs = [_['name'] for _ in data['program'] if _['name'] != u'Н/Д']

    if 'address' in data:
        address = data['address']
        main_address = Address(
            address['title'],
            Coordinates(
                address['lng'],
                address['lat'],
            )
        )
    else:
        main_address = None

    other_addresses = []
    for address in data['others']:
        address = Address(
            address['title'],
            Coordinates(
                address['lng'],
                address['lat'],
            )
        )
        other_addresses.append(address)

    contacts = data['current_director']
    director = Contacts(
        contacts['fio'],
        contacts['phone'],
        contacts['eo_phone'],
        contacts['email']
    )

    return Eduoffice(
        id, inn, state, founder, url,
        title, full_title, programs,
        main_address, other_addresses,
        director
    )


def load_eduoffice(id):
    data = load_raw_eduoffice(id)
    return parse_eduoffice(data)


def get_eduoffice_report_url(id):
    return 'http://map.production.mskobr.ru/api/export/{id}'.format(id=id)


def get_eduoffice_report_filename(id):
    return '{id}.xlsx'.format(id=id)


def get_eduoffice_report_path(id):
    return os.path.join(
        EDUOFFICE_REPORT_DIR,
        get_eduoffice_report_filename(id)
    )


def download_eduoffice_report(id):
    url = get_eduoffice_report_url(id)
    path = get_eduoffice_report_path(id)
    check_call(['wget', url, '-O', path])


def get_host(url):
    return parse_url(url).netloc


def filter_eduoffices(eduoffices):
    for record in eduoffices:
        url = record.url
        programs = record.programs
        if (url and get_host(url).endswith('.mskobr.ru')
            and (u'среднее общее образование' in programs
                 or u'основное общее образование' in programs)
            and u'среднее профессиональное образование' not in programs
            and record.state != u'В стадии закрытия'):
            yield record


def format_eduoffice_full_title(title):
    match = re.match(
        (ur'Государственное бюджетное общеобразовательное учреждение '
         ur'города Москвы "(.+)"'),
        title
    )
    if match:
        return match.group(1)
    match = re.match(
        (ur'Государственное бюджетное (?:образовательное|общеобразовательное) '
         ur'учреждение города Москвы (.+)$'),
        title
    )
    if match:
        return match.group(1).capitalize()
    return


def dump_eduoffices_check(eduoffices, path=EDUOFFICES_CHECK):
    data = []
    for record in eduoffices:
        inn = record.inn
        url = record.url
        title = format_eduoffice_full_title(record.full_title)
        data.append((inn, url, title))
    table = pd.DataFrame(data, columns=['inn', 'url', 'title'])
    table.to_excel(path, index=False)


def load_eduoffices_check():
    data = read_excel(EDUOFFICES_CHECK)
    for index, row in data.iterrows():
        url, title, short, no = row
        yield CheckEduofficeRecord(url, title, short, no)


def get_soup(html):
    return BeautifulSoup(html, 'lxml')


def join_url(base, path):
    return base.rstrip('/') + '/' + path.lstrip('/')


def parse_mksobr_topmenubox(html, base):
    soup = get_soup(html)
    menu = soup.find('div', class_='topmenubox')
    if menu is None:
        return
    for dropdown in menu.find_all('li', class_='dropdown'):
        link = dropdown.find('a')
        program = link.get_text().strip()
        for item in dropdown.find('ul', class_='dropdown-menu').find_all('li'):
            link = item.find('a')
            url = join_url(base, link['href'])
            contents = link.contents
            if len(contents) == 2:
                title, address = link.contents
                title = title.text
                yield MskobrMenuItem(program, title, url, address)


def load_raw_url_menus(urls):
    url_menus = {}
    for url in urls:
        html = load_html(url)
        menu = list(parse_mksobr_topmenubox(html, url))
        url_menus[url] = menu
    return url_menus


def dump_url_menus(url_menus):
    dump_json(url_menus, MSKOBR_URL_MENUS)


def load_url_menus():
    url_menus = {}
    data = load_json(MSKOBR_URL_MENUS)
    for url, menus in data.iteritems():
        menus = [MskobrMenuItem(*_) for _ in menus]
        url_menus[url] = menus
    return url_menus


def load_raw_url_staff_urls(urls, url_menus, include_base=True,
                            programs=(u'Начальное', u'Основное и среднее')):
    def get_link(soup):
        return (
            soup.find(
                'a',
                text=u'Руководство. Педагогический (научно-педагогический) состав'
            )
            or soup.find('a', text=u'Руководство и педагогический состав')
        )

    url_staff = {}
    for base in urls:
        if include_base:
            html = load_html(base)
            soup = get_soup(html)
            link = get_link(soup)
            if link:
                staff = join_url(base, link['href'])
                url_staff[base] = staff
        menu = url_menus[base]
        for item in menu:
            if item.program in programs:
                url = item.url
                html = load_html(url)
                soup = get_soup(html)
                link = get_link(soup)
                if link:
                    staff = join_url(base, link['href'])
                    url_staff[url] = staff
    return url_staff


def dump_url_staff(url_staff, path=MSKOBR_URL_STAFF):
    dump_json(url_staff, path)


def load_url_staff(path=MSKOBR_URL_STAFF):
    return load_json(path)


def load_raw_url_teacher_urls(urls):
    url_teachers = {}
    for url in urls:
        html = load_html(url)
        soup = get_soup(html)
        base = 'http://' + get_host(url)
        teachers = []
        for item in soup.find_all('div', class_='teacherblock'):
            link = item.find('a')
            if link:
                teacher = join_url(base, link['href'])
                teachers.append(teacher)
        url_teachers[url] = teachers
    return url_teachers


def dump_url_teacher_urls(url_teacher_urls, path=MSKOBR_URL_TEACHER_URLS):
    dump_json(url_teacher_urls, path)
    

def load_url_teacher_urls(path=MSKOBR_URL_TEACHER_URLS):
    return load_json(path)


def load_raw_mskobr_teachers(urls):
    for url in urls:
        html = load_html(url)
        soup = get_soup(html)
        base = 'http://' + get_host(url)
        group = soup.find('div', class_='groupteachers')
        if group is not None:
            name = soup.find('div', class_='kris-component-head').find('h1').text
            image = soup.find('div', class_='photo').find('img')
            if image is not None:
                image = join_url(base, image['src'])
            data = {}
            for item in group.find_all('p'):
                components = item.find_all('span')
                if len(components) == 2:
                    title, value = components
                    if title.attrs.get('class') == ['title']:
                        data[title.text] = value.text
            position = data.get(u'Занимаемая должность (должности):')
            discipline = data.get(u'Преподаваемые дисциплины:')
            level = data.get(u'Уровень образования:')
            experience = data.get(u'Стаж работы по специальности: ')
            total_experience = data.get(u'Общий стаж работы:')
            university = data.get(u'Наименование оконченного учебного заведения:')
            category = data.get(u'Категория:')
            yield MskobrTeacherRecord(
                url, name, image, position, discipline, level,
                experience, total_experience, university, category
            )


def dump_mskobr_teachers(teachers, path=MSKOBR_TEACHERS):
    dump_json(teachers, path)


def load_mskobr_teachers(path=MSKOBR_TEACHERS):
    data = load_json(path)
    return [MskobrTeacherRecord(*_) for _ in data]


def split_podgon_cell(cell):
    return [part.strip() for part in cell.split(';')]


def podgon_float(value):
    return float(value.replace(',', '.'))


def parse_podgon_gia(row):
    subjects = row[37]
    if subjects is None:
        return None
    subjects = split_podgon_cell(subjects)
    subject_pupils = split_podgon_cell(row[38])
    grades = split_podgon_cell(row[39])
    gia = {}
    for subject, pupils, grade in zip(subjects, subject_pupils, grades):
        subject = subject.capitalize()
        pupils = int(pupils)
        grade = podgon_float(grade)
        if subject not in gia or gia[subject].pupils < pupils:
            gia[subject] = GiaResult(pupils, grade)
    return gia


def parse_podgon_ege(row):
    total = row[40]
    if total is None:
        return None
    more_220 = row[41]
    more_150 = row[42]
    if isinstance(total, basestring):
        total = [int(_) for _ in split_podgon_cell(total)]
        more_150 = [int(_) for _ in split_podgon_cell(more_150)]
        more_220 = [int(_) for _ in split_podgon_cell(more_220)]
        assert len(total) == len(more_150) == len(more_220)
        total = sum(total)
        more_150 = sum(more_150)
        more_220 = sum(more_220)
    return EgeResults(total, more_150, more_220)


def parse_podgon_olympiads(row):
    types = row[30]
    if types is None:
        return None
    types = split_podgon_cell(types)
    levels = row[31]
    if levels is None or isinstance(levels, int):
        levels = [levels]
    else:
        levels = [
            none_or_int(_)
            for _ in split_podgon_cell(levels)
        ]
        classes = row[32]
        if isinstance(classes, int):
                classes = [classes]
        else:
            classes = [
                none_or_int(_)
                for _ in split_podgon_cell(classes)
            ]
        subjects = [
            _.capitalize()
            for _ in split_podgon_cell(row[33])
        ]
        places = split_podgon_cell(row[34])
        olympiads = defaultdict(list)
        assert len(types) == len(levels) == len(classes)
        assert len(classes) == len(subjects) == len(places)
        for type, level, class_, subject, place in zip(
            types, levels,
            classes, subjects, places
        ):
            olympiads[subject].append(
                OlympiadRecord(type, level, class_, place)
            )
        return dict(olympiads)


def read_excel(path, sheet=0):
    table = pd.read_excel(path, na_values=('x', '-', '?'), sheetname=sheet)
    return table.where(pd.notnull(table), None)


def none_or_int(value):
    if value is not None and value != '':
        return int(value)


def parse_podgon_inn(row):
    inn = row[8]
    if inn is not None:
        return str(int(inn))


def parse_podgon_host(row):
    host = row[17]
    if host:
        host = host.rstrip('/')
        host = PODGON_HOST_ALIASES.get(host, host)
    return host


def load_podgon():
    table = read_excel(PODGON)
    for index, row in table.iterrows():
        id = row[0]
        inn = parse_podgon_inn(row)
        name = row[5]
        number = none_or_int(row[7])
        host = parse_podgon_host(row)
        gia = parse_podgon_gia(row)
        olympiads = parse_podgon_olympiads(row)
        ege = parse_podgon_ege(row)
        yield PodgonRecord(
            id, inn, name, number,
            host, gia, olympiads, ege
        )


def get_olimpiada_serp_url(query):
    return ('http://reg.olimpiada.ru/register/find-school'
            '?region=77&query={query}&captcha=6').format(
        query=quote(query.encode('utf8'))
    )


def download_olimpiada_serp(query):
    url = get_olimpiada_serp_url(query)
    response = requests.get(
        url,
        cookies={
            'captcha': ('0062364983EA3BEB160BA5093288F6E3EC52D8FCC9A9EFD9F731'
                        '73C14FD093C14898718C0406F34DCD068F27326AEAF6C178')
        }
    )
    return response.text


def fetch_olimpiada_serp(query):
    url = get_olimpiada_serp_url(query)
    html = download_olimpiada_serp(query)
    dump_html(url, html)


def parse_olimpiada_serp(html):
    soup = get_soup(html)
    table = soup.find_all('table')[-1]
    for row in table.find_all('tr')[1:]:
        cells = [_.text.strip() for _ in row.find_all('td')]
        if len(cells) == 11:
            login, type, name1, number, name2, name3, name4, region1, region2, region3, id = cells
            yield OlimpiadaSerpRecord(
                login, type, name1, number, name2,
                name3, name4, region1, region2, region3, id 
            )


def load_olimpiada_serps(queries):
    for query in queries:
        url = get_olimpiada_serp_url(query)
        html = load_html(url)
        results = list(parse_olimpiada_serp(html))
        yield OlimpiadaSerp(query, results)
        

def dump_olimpiada_serps_check(serps):
    data = []
    for query, results in serps:
        size = len(results)
        if size > 0:
            correct = size == 1
            for result in results:
                data.append((
                        query,
                        '+' if correct else None,
                        result.login,
                        result.name1,
                        result.name2,
                        result.name3,
                        result.name4
                ))
        else:
            data.append((
                query,
                None,
                None,
                None,
                None,
                None,
                None,
            ))
    table = pd.DataFrame(
        data,
        columns=['query', 'correct', 'login', 'name1', 'name2', 'name3', 'name4']
    )
    table.to_excel(OLIMPIADA_LOGINS, index=False)


def load_olympiada_logins(eduoffices_selection):
    mapping = {_.short: get_host(_.url) for _ in eduoffices_selection}
    table = read_excel(OLIMPIADA_LOGINS)
    logins = {}
    for _, row in table.iterrows():
        if row.correct == '+':
            short = row['query']
            host = mapping[short]
            login = row.login
            yield OlimpiadaLoginRecord(host, login)


def get_olimpiada_russia_results_url(login):
    return 'http://reg.olimpiada.ru/rusolymp-summary/search/?russia&login={login}'.format(
        login=login
    )


def get_olimpiada_moscow_results_url(login):
    return 'http://reg.olimpiada.ru/rusolymp-summary/search/?moscow&login={login}'.format(
        login=login
    )


def download_olimpiada_russia_results(login, year=2014):
    response = requests.post(
        'http://reg.olimpiada.ru/rusolymp-summary/search/',
        data={
            'login': login,
            'year': year,
            'stage': '_',
            'subject': '%'
        }
    )
    return response.text


def download_olimpiada_moscow_results(login, year=2014):
    response = requests.post(
        'http://reg.olimpiada.ru/rusolymp-summary/search/',
        data={
            'login': login,
            'year': year,
            'olygroup': 'mosh'
        }
    )
    return response.text


def fetch_olimpiada_russia_results(login):
    url = get_olimpiada_russia_results_url(login)
    html = download_olimpiada_russia_results(login)
    dump_html(url, html)
    
    
def fetch_olimpiada_moscow_results(login):
    url = get_olimpiada_moscow_results_url(login)
    html = download_olimpiada_moscow_results(login)
    dump_html(url, html)


def parse_olimpiada_results(html):
    soup = get_soup(html)
    table = soup.find_all('table')[-1]
    for row in table.find_all('tr')[1:]:
        cells = [_.text.strip() for _ in row.find_all('td')]
        yield cells
        
        
def parse_olimpiada_russia_results(host, html):
    mapping = {
        u'Русский язык': u'Русский',
        u'Английский язык': u'Английский',
        u'Информатика и ИКТ': u'Информатика',
        u'Основы безопасности жизнедеятельности': u'ОБЖ',
        u'Физическая культура': u'Физкультура',
        u'Искусство (МХК)': u'МХК',
        u'Французский язык': u'Французский',
        u'Немецкий язык': u'Немецкий',
        u'Технология (культура дома)': u'Технология',
        u'Технология (техника и техническое творчество)': u'Технология',
    }
    for surname, name, middle, year, year2, place, olympiad in parse_olimpiada_results(html):
        match = re.search(
            ur'Всероссийская олимпиада по предмету «([\(\)\s\w]+)», ([-\d]+) учебный год, (\d) этап',
            olympiad,
            re.U
        )
        subject, year3, stage = match.groups()
        subject = mapping.get(subject, subject)
        yield OlimpiadaResultRecord(
            host, surname, name, middle, year, year2, place,
            u'Всероссийская олимпиада', subject, year3, stage
        )
        
def parse_olimpiada_moscow_results(host, html):
    mapping = {
        u'математическая олимпиада': u'Математика',
        u'по изобразительному искусству': u'ИЗО',
        u'по физике': u'Физика',
        u'по филологии': u'Филология',
        u'по химии': u'Химия',
        u'по географии': u'География',
        u'по биологии': u'Биология',
        u'по технологии, номинация «Культура дома»': u'Технология',
        u'по экономике': u'Экономика',
        u'по истории': u'История',
        u'по информатике': u'Информатика',
        u'по технологии  «Техника и техническое творчество»': u'Технология',
        u'по обществознанию': u'Обществознание',
        u'по испанскому языку': u'Испанский',
        u'по технологии «Робототехника»': u'Технология',
        u'по праву': u'Право',
        u'по астрономии': u'Астрономия',
        u'по искусству (МХК)': u'МХК',
        u'по лингвистике': u'Лингвистика',
        u'по латинскому языку': u'Латинский язык',
    }
    for surname, name, middle, year, year2, place, olympiad in parse_olimpiada_results(html):
        match = re.search(
            ur'Московская олимпиада школьников (по [\(\)«»,\s\w]+), (\d+) год',
            olympiad,
            re.U
        ) or re.search(
            ur'Московская олимпиада (по [«»,\s\w]+), (\d+) год',
            olympiad,
            re.U
        ) or re.search(
            ur'Московская олимпиада школьников. Московская (математическая олимпиада). (\d+) год',
            olympiad,
            re.U
        ) or re.search(
            ur'Московская олимпиада школьников (по [\s\w]+). ([–\d]+) учебный год',
            olympiad,
            re.U
        ) or re.search(
            ur'Московская традиционная олимпиада (по \w+) \(([–\d]+) учебный год\)',
            olympiad,
            re.U
        )
        subject, year3 = match.groups()
        subject = mapping[subject]
        year3 = {u'2014–2015': u'2014-2015', u'2015': u'2014-2015'}[year3]
        yield OlimpiadaResultRecord(
            host, surname, name, middle, year, year2, place,
            u'Московская олимпиада', subject, year3, None
        )
        
        
def load_olimpiada_results(olimpiada_logins):
    for host, login in olimpiada_logins:
        url = get_olimpiada_moscow_results_url(login)
        html = load_html(url)
        for result in parse_olimpiada_moscow_results(host, html):
            yield result
        url = get_olimpiada_russia_results_url(login)
        html = load_html(url)
        for result in parse_olimpiada_russia_results(host, html):
            yield result

        
def call_geocoder(address):
    response = requests.get(
        'http://geocode-maps.yandex.ru/1.x/',
        params={
            'format': 'json',
            'geocode': address
        }
    )
    if response.status_code == 200:
        return response.json()


def get_address_filename(address):
    return '{hash}.json'.format(
        hash=hash_address(address)
    )


def get_address_path(address):
    return os.path.join(
        ADDRESSES_DIR,
        get_address_filename(address)
    )


def list_addresses_cache():
    return load_items_cache(ADDRESSES_LIST)


def update_addresses_cache(address):
    return update_items_cache(address, ADDRESSES_LIST)


def dump_address(address, data):
    path = get_address_path(address)
    dump_json(data, path)
    update_addresses_cache(address)


def load_raw_address(address):
    path = get_address_path(address)
    return load_json(path)


def geocode_address(address):
    data = call_geocoder(address)
    dump_address(address, data)


def geocode_addresses(addresses):
    for address in addresses:
        geocode_address(address)


def parse_geocode_data(data):
    if data and 'response' in data:
        response = data['response']['GeoObjectCollection']
        data = response['featureMember']
        if data:
            item = data[0]['GeoObject']
            regions = item['description']
            if u'Москва' in regions:
                meta = item['metaDataProperty']['GeocoderMetaData']
                if meta['kind'] == 'house' and meta['precision']:
                    longitude, latitude = item['Point']['pos'].split(' ')
                    longitude = float(longitude)
                    latitude = float(latitude)
                    description = item['name']
                    return Address(
                        description,
                        Coordinates(
                            latitude,
                            longitude
                        )
                    )


def load_address(address):
    data = load_raw_address(address)
    return parse_geocode_data(data)


def dump_addresses_check(eduoffices_selection, url_menus, eduoffices, addresses):
    mapping = {_.url: _ for _ in eduoffices}
    programs = Counter()
    data = []
    for record in eduoffices_selection:
        url = record.url
        menu = url_menus[url]
        items = [
            _ for _ in menu
            if _.program in (u'Начальное', u'Основное и среднее')
        ]
        if items:
            for item in items:
                program = item.program
                raw_address = u'Москва ' + item.address
                address = addresses[raw_address]
                latitude = None
                longitude = None
                if address:
                    coordinates = address.coordinates
                    latitude, longitude = coordinates
                    address = address.description
                else:
                    address = raw_address
                if program == u'Начальное':
                    program = 1
                elif program == u'Основное и среднее':
                    program = 6
                programs[address] |= program
                data.append((
                    address,
                    latitude,
                    longitude,
                    item.url
                ))
        else:
            eduoffice = mapping[url]
            eduoffice_programs = eduoffice.programs
            levels = [
                (_ in eduoffice_programs)
                for _ in (u'начальное общее образование',
                          u'основное общее образование',
                          u'среднее общее образование')
            ]
            assert (levels[1] and levels[2] or (not levels[1] and not levels[2])), (url, levels)
            if not levels[0]:
                program = 6
            else:
                program = 7
            raw_address = u'Москва ' + eduoffice.main_address.description
            address = addresses[raw_address]
            coordinates = address.coordinates
            address = address.description
            programs[address] |= program
            data.append((
                address,
                coordinates.latitude,
                coordinates.longitude,
                url
            ))
    program_data = set()
    for address, latitude, longitude, url in data:
        program = programs[address]
        if program == 1:
            program = u'Начальное'
        elif program == 6:
            program = u'Основное и среднее'
        elif program == 7:
            program = u'Начальное, основное и среднее'
        program_data.add((
            program, address,
            latitude, longitude, url
        ))
    table = pd.DataFrame(
        list(program_data),
        columns=['program', 'address', 'latitude', 'longitude', 'url']
    )
    table.to_excel(ADDRESSES_CHECK, index=False)
            

def load_address_check(path=ADDRESSES_CHECK):
    table = read_excel(path)
    for index, row in table.iterrows():
        program, address, latitude, longitude, url = row
        if latitude and longitude:
            address = Address(
                address,
                Coordinates(
                    latitude,
                    longitude
                )
            )
            yield CheckAddressRecord(program, address, url)


def load_raw_url_pennants(urls):
    url_pennants = {}
    for url in urls:
        html = load_html(url)
        soup = get_soup(html)
        block = soup.find('div', class_='pennants')
        pennants = []
        if block is not None:
            for item in block.find_all('a', class_='kris-regalii-box'):
                match = re.match(
                    (r"background-image: url\('/templates/img/"
                     r"pennants-([^\.]+).png'\);"),
                    item['style']
                )
                href = match.group(1)
                pennants.append(href)
        url_pennants[url] = pennants
    return url_pennants


def dump_url_pennants(url_pennants):
    dump_json(url_pennants, MSKOBR_URL_PENNANTS)


def load_url_pennants():
    return load_json(MSKOBR_URL_PENNANTS)


def parse_rating_2015(html):
    soup = get_soup(html)
    table = soup.find('table', class_='tablesorter')
    for row in table.find_all('tr')[1:]:
        place, link = row.find_all('td')
        place = int(place.text)
        link = link.find('a')
        if link:
            url = link['href']
            url = url.rstrip('/')
            if not url.startswith('http://'):
                url = 'http://' + url
            yield url, place


def load_rating_2015():
    rating = {}
    for url in RATING_2015_URLS:
        html = load_html(url)
        for url, place in parse_rating_2015(html):
            rating[url] = place
    return rating


def dump_ratings_check(eduoffices_selection, url_pennants, url_rating_2015):
    data = []
    for record in eduoffices_selection:
        url = record.url
        pennants = url_pennants[url]
        tops = [None, None, None]
        for index, label in enumerate(['top400-2015', 'top400-2014', 'top400-2013']):
            if label in pennants:
                tops[index] = '+'
        rating_2015 = url_rating_2015.get(url)
        data.append(
            [url] + tops + [rating_2015]
        )
    table = pd.DataFrame(
        data,
        columns=['url', 'top_2015', 'top_2014', 'top_2013', 'rating_2015']
    )
    table.to_excel(RATINGS_CHECK, index=False)


def parse_rating_2014(html):
    soup = get_soup(html)
    table = soup.find('table', class_='tablesorter')
    for row in table.find_all('tr')[1:]:
        cells = row.find_all('td')
        rating = int(cells[0].text)
        title = cells[1].text.strip()
        yield title, rating


def load_rating_2014():
    html = load_html(RATING_2014_URL)
    return dict(parse_rating_2014(html))


def parse_rating_2013(html):
    soup = get_soup(html)
    table = soup.find('table', class_='all_border')
    for row in table.find_all('tr')[1:]:
        cells = row.find_all('td')
        rating = int(cells[0].text)
        title = cells[1].text.strip()
        yield title, rating


def load_rating_2013():
    html = load_html(RATING_2013_URL)
    return dict(parse_rating_2013(html))


def load_ratings_check():
    table = read_excel(RATINGS_CHECK)
    for index, row in table.iterrows():
        yield CheckRatingRecord(
            row.url,
            none_or_int(row.rating_2013),
            none_or_int(row.rating_2014),
            none_or_int(row.rating_2015)
        )


def get_common_prefix(strings):
    first = min(strings)
    last = max(strings)
    for index, char in enumerate(first):
        if char != last[index]:
            return first[:index]
    return first


def get_viz_schools(eduoffices_selection, eduoffices, ratings, school_addresses,
                    ege, olympiads, url_universities, contacts, teachers,
                    images, reviews, polls, eduoffice_reports, sad_selection):
    host_eduoffices = {get_host(_.url): _ for _ in eduoffices if _.url}
    host_rating = {get_host(_.url): _ for _ in ratings}

    host_ege = {}
    for record in ege:
        host, less_150, between_150_220, more_220, a, b, source = record
        host_ege[host] = VizSchoolEge(
            less_150, between_150_220, more_220, a, b, source
        )

    host_olympiads = defaultdict(list)
    for record in olympiads:
        host, subject, olympiad, place, count = record
        subject = OLYMPIAD_ALIASES.get(subject, subject)
        if olympiad == u'Московская олимпиада' and place == u'победитель':
            type = OLYMPIAD_MOSCOW_1
        elif olympiad == u'Московская олимпиада' and place == u'призёр':
            type = OLYMPIAD_MOSCOW_2
        elif olympiad == u'Всероссийская олимпиада' and place == u'победитель':
            type = OLYMPIAD_RUSSIA_1
        elif olympiad == u'Всероссийская олимпиада' and place == u'призёр':
            type = OLYMPIAD_RUSSIA_2
        host_olympiads[host].append(VizSchoolOlympiads(
            subject, type, count
        ))

    host_universities = {}
    for url, universities in url_universities.iteritems():
        host = get_host(url)
        records = []
        if isinstance(universities, dict):
            for university, ids in universities.iteritems():
                records.append(VizSchoolUniversity(
                    university, ids, 'http://vk.com'
                ))
        elif isinstance(universities, list):
            for record in universities:
                records.append(VizSchoolUniversity(
                    record.university, record.count, record.source
                ))
        host_universities[host] = records

    host_addresses = defaultdict(set)
    for record in school_addresses:
        host = get_host(record.url)
        address = record.address
        coordinates = address.coordinates
        program = record.program
        if program == '1..11':
            program = PROGRAM_1_11
        elif program == '5..11':
            program = PROGRAM_5_11
        elif program == '6..11':
            program = PROGRAM_6_11
        elif program == '1..4':
            program = PROGRAM_1_4
        elif program == '1..9':
            program = PROGRAM_1_9
        elif program == '10..11':
            program = PROGRAM_10_11
        elif program == '8..11':
            program = PROGRAM_8_11
        elif program == '1..7':
            program = PROGRAM_1_7
        else:
            assert False, (host, program)
        host_addresses[host].add(VizSchoolAdress(
            address.description,
            coordinates.latitude,
            coordinates.longitude,
            program
        ))

    host_contacts = defaultdict(list)
    for record in contacts:
        url, position, name, work, mobile, email = record
        host = get_host(url)
        host_contacts[host].append(VizSchoolContact(
            position, name, work, mobile, email
        ))

    host_teachers = defaultdict(list)
    for record in teachers:
        url = record.url
        host = get_host(url)
        if record.name is None:
            url = EMAIL_SOURCE
        category = record.category
        if category == u'Высшая':
            category = CATEGORY_TOP
        elif category == u'Первая':
            category = CATEGORY_1
        elif category == u'Нет':
            category = CATEGORY_NO
        elif category is None:
            category = CATEGORY_UNKNOWN
        host_teachers[host].append(VizTeacher(
            record.subject,
            url,
            category
        ))

    host_images = defaultdict(list)
    for image in images:
        url = image.url
        match = re.match(
            r'^http://raw.githubusercontent.com/alexanderkuk'
            u'/obr.msk.ru-images/master/([^/]+)/',
            url
        )
        if match:
            host = match.group(1) + '.mskobr.ru'
        else:
            host = get_host(image.url)
        host_images[host].append(image)

    host_reviews = defaultdict(list)
    for review in reviews:
        host_reviews[review.host].append(review)

    host_polls = defaultdict(list)
    for poll in polls:
        host_polls[poll.host].append(poll)

    host_reports = {
        _.host: _
        for _ in eduoffice_reports
    }

    host_sads = set()
    for record in sad_selection:
        host = get_host(record.url)
        host_sads.add(host)

    for record in eduoffices_selection:
        host = get_host(record.url)
        eduoffice = host_eduoffices[host]
        
        full_title = eduoffice.full_title
        short_title = record.title
        label_title = record.short
        link_title = re.match(r'([\w\d-]+)\.mskobr\.ru', host).group(1)
        no_title = record.no

        sad_link = None
        if host in host_sads:
            sad_link = link_title + '-sad.html'
        
        rating = host_rating[host]
        addresses = list(host_addresses[host])

        programs = tuple({_.program for _ in addresses})

        ege = host_ege.get(host)        
        olympiads = host_olympiads.get(host)
        universities = host_universities.get(host)

        contacts = host_contacts[host]
        teachers = host_teachers.get(host)
        images = host_images.get(host)
        reviews = host_reviews.get(host)
        polls = host_polls.get(host)
        report = host_reports.get(host)

        yield VizSchool(
            record.url,
            sad_link,
            VizSchoolTitles(
                full_title,
                short_title,
                label_title,
                link_title,
                no_title
            ),
            VizSchoolRating(
                rating.year_2013,
                rating.year_2014,
                rating.year_2015
            ),
            programs,
            addresses,
            ege,
            olympiads,
            universities,
            contacts,
            teachers,
            images,
            reviews,
            polls,
            report
        )


def get_school_page_filename(school):
    return '{link}.html'.format(
        link=school.title.link
    )


def get_school_page_path(school):
    return os.path.join(
        SITE_DIR,
        get_school_page_filename(school)
    )


def format_review_date(date):
    return u'{month} {year}'.format(
        month=REVIEW_MONTHS[date.month],
        year=date.year
    )


def format_review_text(text, hyphen=HYPHEN):
    # KISS
    return text

    def hyphen_word(match, hyphen=hyphen):
        word = match.group(1)
        if len(word) < 4:
            return word
        return hyphen.inserted(word, '&shy;')
    
    return re.sub(ur'([^\s\.!"\'\(\),.:;\?]+)', hyphen_word, text)


def generate_school_page(school, template):
    path = get_school_page_path(school)
    rating = school.rating
    programs = school.programs
    ege = school.ege
    if ege:
        less_150, between_150_220, more_220, a, b, ege_source = ege
        ege_total = less_150 + between_150_220 + more_220
    else:
        less_150 = None
        between_150_220 = None
        more_220 = None
        a = None
        b = None
        ege_source = None
        ege_total = None

    olympiads = school.olympiads
    subject_olympiads = defaultdict(dict)
    olympiad_types = set()
    if olympiads:
        for olympiad in olympiads:
            type = olympiad.type
            subject_olympiads[olympiad.subject][type] = olympiad.count
            olympiad_types.add(type)
    olympiads = []
    for subject in OLYMPIADS_ORDER:
        if subject in subject_olympiads:
            types = subject_olympiads[subject]
            counts = []
            for type in OLYMPIAD_TYPES_ORDER:
                if type in types:
                    count = types[type]
                    counts.append((type, count))
            olympiads.append((subject, counts))

    universities = school.universities
    universities_source = None
    if universities:
        record = universities[0]
        universities_source = record.source
        if universities_source == 'http://vk.com':
            universities = [
                (_.name, ','.join(str(id) for id in _.ids))
                for _ in sorted(
                        universities,
                        key=lambda _: len(_.ids),
                        reverse=True
                )
            ][:20]
        else:
            universities = [
                (_.name, _.ids)
                for _ in sorted(
                        universities,
                        key=lambda _: _.ids,
                        reverse=True
                )
            ]

    mgu_estimate = None
    if school.title.label == u'Лицей № 1535':
        mgu_estimate = '50%'

    teachers = school.teachers
    subject_teachers = defaultdict(list)
    teacher_categories = set()
    teacher_sources = set()
    if teachers:
        for subject, url, category in teachers:
            if url == EMAIL_SOURCE:
                teacher_sources.add(EMAIL_SOURCE)
            else:
                teacher_sources.add(URL_SOURCE)
            subject_teachers[subject].append((url, category))
            teacher_categories.add(category)
    teachers = []
    for subject in TEACHER_SUBJECTS_ORDER:
        if subject in subject_teachers:
            categories = sorted(
                subject_teachers[subject],
                key=lambda (_, category): CATEGORIES_ORDER.index(category)
            )
            teachers.append((subject, categories))

    images = school.images
    images_source = None
    if images:
        if all('obr.msk.ru-images' in _.url for _ in images):
            images_source = EMAIL_SOURCE
        else:
            host = get_host(school.url)
            images_source = 'http://{host}/main_galleries'.format(
                host=host
            )
            
        images = [
            (_.url, _.filename, _.raw.width, _.raw.height)
            for _ in images
        ]

    reviews = school.reviews
    if reviews:
        reviews = sorted(reviews, key=lambda _: _.date, reverse=True)
        reviews = [
            (get_host(_.url), _.url, format_review_date(_.date),
             _.name, format_review_text(_.text))
            for _ in reviews
        ]

    polls = school.polls
    if polls:
        questions_mapping = {
            u'Если бы можно было начать всё сначала, вы бы снова пошли в вашу школу или выбрали бы другую?': u'Если бы можно было начать всё сначала, ...',
            u'Чтобы хорошо сдать ЕГЭ, занятий в школе достаточно или нужны репетиторы?': u'Чтобы хорошо сдать ЕГЭ, ...'
        }
        answers_mapping = {
            u'Снова в мою': u'снова пошли бы в эту школу',
            u'В другую': u'выбрали бы другую школу',
            u'Занятий в школе достаточно': u'занятий в школе достаточно',
            u'Нужны репетиторы': u'нужны репетиторы'
        }
        records = polls
        polls = []
        for poll in records:
            question = poll.question
            question = questions_mapping.get(question, question)
            answers = []
            for text, count in poll.answers:
                text = answers_mapping.get(text, text)
                if count > 0:
                    answers.append((text, count))
            if answers:
                polls.append((question, answers))

    report = school.eduoffice_report
    teachers_count = None
    pupils_count = None
    pupils_per_teacher = None
    teachers_salary = None
    if report:
        teachers_count = report.teachers
        pupils_count = report.pupils
        pupils_per_teacher = pupils_count / teachers_count
        teachers_salary = int(report.salaries)

    title = school.title
    url = school.url
    link = title.link
    page = link + '.html'

    html = template.render(
        url=url,
        host=get_host(url),
        full_title=title.short,
        label_title=title.label,
        link=link,
        sad_link=school.sad_link,
        path=page,

        rating_2013=rating.year_2013,
        rating_2014=rating.year_2014,
        rating_2015=rating.year_2015,

        program_1_4=(PROGRAM_1_4 in programs),
        program_5_11=(PROGRAM_5_11 in programs),
        program_6_11=(PROGRAM_6_11 in programs),
        program_1_11=(PROGRAM_1_11 in programs),
        program_1_9=(PROGRAM_1_9 in programs),
        program_10_11=(PROGRAM_10_11 in programs),
        program_1_7=(PROGRAM_1_7 in programs),
        program_8_11=(PROGRAM_8_11 in programs),
        addresses=school.addresses,

        ege_source=ege_source,
        ege_total=ege_total,
        less_150=less_150,
        between_150_220=between_150_220,
        more_220=more_220,
        a=a,
        b=b,

        moscow_1=(OLYMPIAD_MOSCOW_1 in olympiad_types),
        moscow_2=(OLYMPIAD_MOSCOW_2 in olympiad_types),
        russia_1=(OLYMPIAD_RUSSIA_1 in olympiad_types),
        russia_2=(OLYMPIAD_RUSSIA_2 in olympiad_types),
        olympiads=olympiads,

        universities_source=universities_source,
        universities=universities,
        mgu_estimate=mgu_estimate,

        contacts = school.contacts,
        
        top_category=(CATEGORY_TOP in teacher_categories),
        category_1=(CATEGORY_1 in teacher_categories),
        category_no=(CATEGORY_NO in teacher_categories),
        unknown_category=(CATEGORY_UNKNOWN in teacher_categories),
        email_teachers_source=(EMAIL_SOURCE in teacher_sources),
        url_teachers_source=(URL_SOURCE in teacher_sources),
        teachers=teachers,

        images_source=images_source,
        images=images,

        reviews=reviews,
        polls=polls,

        teachers_count=teachers_count,
        pupils_count=pupils_count,
        pupils_per_teacher=pupils_per_teacher,
        teachers_salary=teachers_salary
    )
    dump_text(html, path)
    
    
def format_salary(salary):
    return '{:,}'.format(salary).replace(',', '&nbsp;')


def generate_school_pages(schools):
    template = load_text(SCHOOL_TEMPLATE)
    env = Environment()
    env.filters['format_salary'] = format_salary
    template = env.from_string(template)
    for school in schools:
        generate_school_page(school, template)


def generate_index(schools, sads):
    records = []
    for sad in sads:
        url = sad.title.link + '.html'
        title = sad.title.full
        addresses = sad.addresses
        addresses = [
            (
                _.address, _.latitude, _.longitude, PROGRAM_0_1,
                0 if len(addresses) == 1 else index + 1
            )
            for index, _ in enumerate(addresses)
        ]
        records.append((
            url, title,
            addresses
        ))

    for school in schools:
        url = school.title.link + '.html'
        title = school.title.short
        addresses = school.addresses
        addresses = [
            (
                _.address, _.latitude, _.longitude, _.program,
                0 if len(addresses) == 1 else index + 1
            )
            for index, _ in enumerate(addresses)
        ]
        records.append((
            url, title,
            addresses
        ))
    
    template = load_text(INDEX_TEMPLATE)
    template = Template(template)
    html = template.render(records=records)
    dump_text(html, INDEX)


def generate_list(schools, sads):
    records = []
    schools = sorted(
        schools,
        key=lambda _: _.rating.year_2015 or float('inf')
    )
    host_sads = {
        get_host(_.url): _
        for _ in sads
    }
    for school in schools:
        path = school.title.link + '.html'
        title = school.title.short
        url = school.url
        host = get_host(url)
        sad_path = None
        if host in host_sads:
            sad = host_sads[host]
            sad_path = sad.title.link + '.html'
        rating = school.rating.year_2015
        records.append((
            path, sad_path, title, host, url,
            rating
        ))
    template = load_text(LIST_TEMPLATE)
    template = Template(template)
    html = template.render(schools=records)
    dump_text(html, LIST)
    

def load_ege():
    table = read_excel(EGE)
    table.columns = ['title', 'short', 'address', 'year',
                     'total', 'more_220', 'more_150', 'id']
    for index, row in table.iterrows():
        title = row.title
        short = row.short
        result = EgeResults(row.total, row.more_150 or 0, row.more_220 or 0)
        yield EgeRecord(title, short, result)


def dump_ege_title_host_check(ege, eduoffices_selection):
    data = []
    for record in ege:
        short = record.short
        match = re.search(ur'(№ \d+)', record.short, re.U)
        if match:
            pattern = match.group(1)
        else:
            pattern = short
        records = []
        for record in eduoffices_selection:
            short2 = record.short
            if pattern in short2:
                records.append(record)
        correct = False
        if len(records) == 1:
            correct = True
        for record in records:
            data.append((
                short,
                '+' if correct else None,
                record.short,
                get_host(record.url)
            ))
    table = pd.DataFrame(data, columns=['short', 'correct', 'short2', 'host'])
    table.to_excel(EGE_TITLE_HOST_CHECK, index=False)


def load_ege_title_host_check():
    table = read_excel(EGE_TITLE_HOST_CHECK)
    for index, row in table.iterrows():
        if row.correct == '+':
            yield EgeTitleHostRecord(row.short, row.host)


def get_ege_intervals(ege, ege_title_host, podgon):
    host_results = {}
    mapping = {_.title: _.host for _ in ege_title_host}
    for record in ege:
        host = mapping.get(record.short)
        if host:
            host_results[host] = record.result
    podgon_host_results = {}
    for record in podgon:
        host = record.host
        if host:
            ege = record.ege
            if ege:
                podgon_host_results[host] = ege
    for host in set(host_results) | set(podgon_host_results):
        results = host_results.get(host) or podgon_host_results[host]
        total, more_150, more_220 = results
        yield EgeIntervalsRecord(
            host,
            total - more_150,
            more_150 - more_220,
            more_220
        )


def get_distribution(a, b, scale=1):
    size = 300
    total = 0
    distribution = []
    for index in xrange(1, size + 1):
        p = float(index) / size
        item = (1 - p) ** a * p ** b
        total += item
        distribution.append(item)
    for item in distribution:
        yield item / total * scale
        

def get_disribution_intervals(distribution):
    return (
        sum(distribution[:150]),
        sum(distribution[150:220]),
        sum(distribution[220:])
    )
    
        
def get_intervals_error(etalon, guess):
    return sum((a - b) ** 2 for a, b in zip(etalon, guess))


def float_xrange(start, stop, step):
    while start < stop:
        yield start
        start += step

        
def fit_distribution(record, as_=(1, 30, 0.5), bs=(1, 30, 0.5)):
    _, less_150, between_150_220, more_220 = record
    etalon = (less_150, between_150_220, more_220)
    total = sum(etalon)
    errors = []
    for a in float_xrange(*as_):
        for b in float_xrange(*bs):
            distribution = list(get_distribution(a, b, scale=total))
            guess = get_disribution_intervals(distribution)
            error = get_intervals_error(etalon, guess)
            errors.append((error, a, b))
    _, a, b = min(errors)
    return a, b


def dump_ege_distributions_check(ege_distributions, eduoffices_selection):
    host_title = {get_host(_.url): _.title for _ in eduoffices_selection}
    data = []
    for record, (a, b) in ege_distributions.iteritems():
        host, less_150, between_150_220, more_220 = record
        title = host_title[host]
        data.append((host, title, less_150, between_150_220, more_220, a, b))
    table = pd.DataFrame(
        data,
        columns=['host', 'title', '0..150', '150..220', '220..300', 'a', 'b']
    )
    table.to_excel(EGE_DISTRIBUTIONS_CHECK, index=False)


def load_ege_distributions_check():
    table = read_excel(EGE_DISTRIBUTIONS_CHECK)
    for index, row in table.iterrows():
        host, title, less_150, between_150_220, more_220, a, b, source = row
        yield CheckEgeRecord(host, less_150, between_150_220, more_220, a, b, source)


def dump_olympiads_check(olimpiada_results):
    records = []
    for record in olimpiada_results:
        if record.year in ('9', '10', '11') and record.year3 == '2014-2015':
            if record.type == u'Всероссийская олимпиада' and record.stage == '4':
                records.append(record)
            elif (record.type == u'Московская олимпиада'
                  and record.place in (u'призёр', u'победитель')):
                records.append(record)
    counts = Counter()
    for record in records:
        counts[record.host, record.year, record.subject, record.type, record.place] += 1
    data = []
    for (host, year, subject, olympiad, place), count in counts.most_common():
        data.append((
            host, None, year, subject, olympiad, place, count
        ))
    table = pd.DataFrame(
        data,
        columns=['host', 'title', 'year', 'subject', 'olympiad', 'place', 'count']
    )
    table.to_excel(OLYMPIADS_CHECK, index=False)


def load_olympiads_check():
    table = read_excel(OLYMPIADS_CHECK)
    data = Counter()
    for index, row in table.iterrows():
        host, title, year, subject, olympiad, place, count = row
        data[host, subject, olympiad, place] += count
    for (host, subject, olympiad, place), count in data.iteritems():
        yield CheckOlympiadsRecord(host, subject, olympiad, place, count)


class ApiCallError(Exception):
    pass


def call_vk(method, v=5.37, token=VK_TOKEN, **params):
    params.update(v=v, access_token=token)
    response = requests.get(
        'https://api.vk.com/method/' + method,
        params=params,
        timeout=10
    )
    data = response.json()
    if 'error' in data:
        raise ApiCallError(data['error'])
    else:
        return data['response']


def get_vk_script_command(method, **params):
    return 'API.{method}({params})'.format(
        method=method,
        params=json.dumps(params)
    )


def search_vk_schools(pattern, call_vk=call_vk):
    return call_vk(
        'database.getSchools',
        q=pattern,
        city_id=1,
        count=10000
    )


def load_school_vk_ids():
    table = read_excel(VK_SCHOOL_IDS)
    url_ids = {}
    for index, row in table.iterrows():
        url, id = row
        if id:
            url_ids[url] = int(id)
    return url_ids


def download_vk_pupils(school_id, sex=0, age_from=None, age_to=None, call_vk=call_vk):
    return call_vk(
        'users.search',
        sort=1, # by registration, should be less skewed then by popularity
        offset=0,
        count=1000,
        fields='universities',
        city=1,
        country=1,
        school=school_id,
        sex=sex,
        age_from=age_from,
        age_to=age_to
    )


def get_vk_pupils_filename(school_id):
    return '{school_id}.json'.format(school_id=school_id)


def parse_vk_pupils_filename(filename):
    school_id, _ = filename.split('.', 1)
    school_id = int(school_id)
    return school_id


def list_vk_pupils_cache():
    for filename in os.listdir(VK_PUPILS_DIR):
        yield parse_vk_pupils_filename(filename)


def get_vk_pupils_path(school_id):
    filename = get_vk_pupils_filename(school_id)
    return os.path.join(
        VK_PUPILS_DIR,
        filename
    )


def dump_vk_pupils(pupils, school_id):
    path = get_vk_pupils_path(school_id)
    dump_json(pupils, path)


def parse_vk_pupils(data):
    if 'response' in data:
        data = data['response']
    if isinstance(data, list):
        items = [
            item for chunk in data
            for item in chunk['items']
        ]
    else:
        items = data['items']
    ids = set()
    for item in items:
        id = item['id']
        if id in ids:
            # Since same ids in different chunks can accure
            continue
        ids.add(id)
        name = u'{surname} {name}'.format(
            surname=item['last_name'],
            name=item['first_name']
        )
        universities = []
        if 'universities' in item:
            for university in item['universities']:
                university = VkUniversity(
                    university['id'],
                    university['name'],
                    university.get('faculty_name'),
                    university.get('education_form'),
                    university.get('education_status'),
                    university.get('graduation')
                )
                universities.append(university)
        yield VkPupil(id, name, universities)


def load_vk_pupils(id):
    path = get_vk_pupils_path(id)
    data = load_json(path)
    return list(parse_vk_pupils(data))


def get_vk_university_names(vk_pupils):
    university_options = defaultdict(set)
    for _, pupils in vk_pupils.iteritems():
        for pupil in pupils:
            for university in pupil.universities:
                university_options[university.id].add(university.name)
    names = {}
    for id, options in university_options.iteritems():
        name = min(options, key=len)
        name = name.strip()
        names[id] = name
    return names


def get_vk_school_universities(vk_pupils):
    # Ban MSU and some other universities that start with "A"
    banned_uninersity_ids = {2, 92, 87, 348, 86, 90, 89, 88, 93, 94, 95}
    university_names = get_vk_university_names(vk_pupils)
    url_universities = {}
    for url, pupils in vk_pupils.iteritems():
        universities = defaultdict(list)
        for pupil in pupils:
            for university in pupil.universities:
                university_id = university.id
                if university_id not in banned_uninersity_ids:
                    name = university_names[university_id]
                    universities[name].append(pupil.id)
        url_universities[url] = universities
    return url_universities


def dump_url_universities(url_universities):
    dump_json(url_universities, URL_UNIVERSITIES)


def load_url_universities():
    return load_json(URL_UNIVERSITIES)


def download_public_wall(name, offset=0, count=100):
    return call_vk(
        'wall.get',
        domain=name,
        offset=offset,
        count=count,
        filter='all',
        extended=0
    )


def parse_public_wall_polls(data):
    for item in data['items']:
        for attachment in item['attachments']:
            if attachment['type'] == 'poll':
                poll = attachment['poll']
                id = poll['id']
                owner = poll['owner_id']
                question = poll['question']
                answers = [
                    VkPollAnswer(_['id'], _['text'])
                    for _ in poll['answers']
                ]
                yield VkPoll(id, owner, question, answers)


def download_poll_votes(poll, offset=0, count=1000):
    return call_vk(
        'polls.getVoters',
        owner_id=poll.owner,
        poll_id=poll.id,
        answer_ids=','.join(str(_.id) for _ in poll.answers),
        id_board=0,
        friends_only=0,
        offset=offset,
        count=count
    )


def parse_poll_votes(data):
    for item in data:
        answer = item['answer_id']
        for user in item['users']['items']:
            yield VkPollVote(answer, user)


def get_poll_stats(vk_pupils, vk_polls, poll_answers):
    user_host = {}
    for url, users in vk_pupils.iteritems():
        host = get_host(url)
        for user in users:
            user_host[user.id] = host
    answer_question = {}
    answer_texts = {}
    for record in vk_polls:
        question = record.question
        for answer in record.answers:
            id = answer.id
            answer_question[id] = question
            answer_texts[id] = answer.text
    stats = defaultdict(Counter)
    for record in poll_answers:
        user = record.user
        host = user_host.get(user)
        answer = record.answer
        question = answer_question[answer]
        text = answer_texts[answer]
        stats[host][question, text] += 1
    for host in stats:
        counts = stats[host]
        for record in vk_polls:
            question = record.question
            answers = [
                VkPollAnswerRecord(
                    answer.text,
                    counts[question, answer.text]
                )
                for answer in record.answers
            ]
            yield VkPollStatsRecord(host, question, answers)


def load_contacts_check():
    table = read_excel(CHECK_CONTACTS)
    for index, row in table.iterrows():
        yield CheckContactsRecord(*row)


def parse_teacher_level(level):
    if level is not None:
        level = level.lower()
        if u'высш' in level:
            return u'высшее'
        elif u'средне' in level:
            if u'спец' in level:
                return u'среднее специальное'
            elif u'проф' in level:
                return u'среднее профессиональное'
            else:
                return u'среднее'
        elif u'спец' in level:
            return u'специалист'
        else:
            return None


def parse_teacher_experience(experience):
    if experience is not None:
        experience = experience.lower().strip()
        if experience.isdigit():
            return float(experience)
        match = (re.search(ur'^([,\d]+) (:?лет|года|год)', experience, re.U)
                 or re.search(ur'более (\d+) лет', experience, re.U))
        if match:
            return float(match.group(1).replace(',', '.'))
        match = re.search(ur'\bс (\d{4})\s*г?', experience, re.U)                 
        if match:
            return 2016 - float(match.group(1))
        match = re.search(ur'^([,\d]+) месяц', experience, re.U)
        if match:
            return float(match.group(1)) / 12
        return None


def parse_teacher_university(university):
    if university is not None:
        university = university.lower()
        if (u'мпгу' in university or u'мгпу' in university or u'мгпи' in university
            or u'педагог' in university):
            return u'педагогический'
        return university


def parse_teacher_category(category):
    if category is not None:
        category = category.lower()
        if u'высшая' in category:
            return u'Высшая'
        elif u'вторая' in category or u'2' in category or u'ii' in category:
            return u'Вторая'
        elif u'первая' in category or u'1' in category or u'i' in category:
            return u'Первая'
        elif (u'соотв' in category or u'нет' in category or u'без' in category
              or u'отсутст' in category):
            return u'Нет'
        return None


def parse_teacher_position(position):
    if position is not None:
        position = position.lower()
        if u'математ' in position or u'алгебра' in position:
            return u'Математика'
        elif u'физик' in position:
            return u'Физика'
        elif u'информ' in position:
            return u'Информатика'
        elif u'биолог' in position:
            return u'Юиология'
        elif u'хим' in position:
            return u'Химия'
        elif u'геогр' in position:
            return u'География'
        elif u'англ' in position:
            return u'Английский'
        elif u'иностр' in position or u'франц' in position or u'немец' in position:
            return u'Иностранный язык'
        elif u'русск' in position or u'литер' in position:
            return u'Русский и литература'
        elif u'истор' in position or u'обществ' in position:
            return u'История и обществознание'
        elif u'физич' in position or u'физкульт' in position:
            return u'Физкультура' 
        elif u'музык' in position:
            return u'Музыка'
        elif u'изо' in position or u'мхк' in position:
            return u'ИЗО'
        elif u'обж' in position or u'безопас' in position:
            return u'ОБЖ'
        elif u'технол' in position:
            return u'Технология'
        elif u'директор' in position:
            return u'Руководитель'
        elif u'начальн' in position:
            return u'Начальные классы'
        elif u'воспитат' in position:
            return u'Воспитатель'
        else:
            return None


def parse_sad_teacher_position(position):
    position = position.lower()
    if u'воспитат' in position:
        if u'старш' in position:
            return u'Старший воспитатель'
        elif u'помощ' in position:
            return u'Помощник воспитателя'
        else:
            return u'Воспитатель'
    elif u'муз' in position:
        return u'Муз. рук.'
    elif u'логопед' in position:
        return u'Логопед'
    elif u'псих' in position:
        return u'Психолог'
    elif (u'физиче' in position or u'физк' in position
          or u'физо' in position or u'плаван' in position):
        return u'Физ. рук.'
    elif u'метод' in position:
        return u'Методист'
    elif u'завхо' in position or u'хозяй' in position:
        return u'Завхоз'
    elif u'дефект' in position:
        return u'Дефектолог'
    elif u'руководит' in position:
        return u'Руководитель'
    else:
        return None


def unique_teachers(teachers):
    return {(get_host(_.url), _.name): _ for _ in teachers}.values()


def dump_teachers_check(teachers, path=CHECK_TEACHERS):
    data = []
    for record in teachers:
        position = (parse_teacher_position(record.position)
                    or parse_teacher_position(record.discipline))
        if position:
            if position == u'Юиология':
                position = u'Биология'
            if position in TEACHER_SUBJECTS_ORDER:
                category = parse_teacher_category(record.category)
                if category == u'Вторая':
                    category = u'Нет'
                data.append((
                    record.url,
                    position,
                    record.name.strip(),
                    category
                ))
    table = pd.DataFrame(data, columns=['url', 'subject', 'name', 'category'])
    table.to_excel(path, index=False)


def load_teachers_check(path=CHECK_TEACHERS):
    table = read_excel(path)
    for index, row in table.iterrows():
        yield CheckTeachersRecord(*row)


def get_raw_url_galleries(urls, url_menus, include_base=True,
                          programs=(u'Начальное', u'Основное и среднее')):
    url_galleries = {}
    for base in urls:
        menu = url_menus[base]
        menu_urls = [_.url for _ in menu if _.program in programs]
        if include_base:
            menu_urls += [base]
        for url in menu_urls:
            html = load_html(url)
            soup = get_soup(html)
            link = (soup.find('a', text=u'Фотогалерея')
                    or soup.find('a', text=u'Фото')
                    or soup.find('a', text=u'Фотоальбомы'))
            if link:
                gallery = join_url(base, link['href'])
                url_galleries[url] = gallery
    return url_galleries


def get_url_gallery_urls(urls, url_galleries):
    url_gallery_urls = {}
    for url in urls:
        gallery = url_galleries[url]
        html = load_html(gallery)
        soup = get_soup(html)
        base = 'http://' + get_host(url)
        gallery_urls = []
        for item in soup.find_all('div', class_='kris-album-item'):
            link = item.find('a')
            url = join_url(base, link['href'])
            if not url.endswith(('galleries/photo/', 'obwie_svedeniya/photo/')):
                # not link to subgallery
                gallery_urls.append(url)
        url_gallery_urls[gallery] = gallery_urls
    return url_gallery_urls


def get_url_image_urls(gallery_urls, url_gallery_urls):
    url_image_urls = {}
    for gallery_url in gallery_urls:
        for url in url_gallery_urls[gallery_url]:
            html = load_html(url)
            soup = get_soup(html)
            base = 'http://' + get_host(url)
            image_urls = []
            for item in soup.find_all('div', class_='kris-album-item-in'):
                link = item.find('a')
                image_url = join_url(base, link['href'])
                image_urls.append(image_url)
            url_image_urls[url] = image_urls
    return url_image_urls


def dump_gallery_images(gallery_images, path=GALLERY_IMAGES):
    dump_json(gallery_images, path)


def load_gallery_images(path=GALLERY_IMAGES):
    return load_json(path)


def images_check_filename(host):
    return host + '.html'


def images_check_path(host):
    return os.path.join(
        CHECK_IMAGES_DIR,
        images_check_filename(host)
    )


def get_img_html(url):
    return (u'<img src="{url}" style="width:30%;display:inline;margin:10px"'
            u'onclick="console.log(\'{stripped}\')"/>').format(
                url=url,
                stripped=url[7:]
            )

def dump_images_check_dir(gallery_images):
    urls = {url for urls in gallery_images.itervalues() for url in urls}
    htmls = defaultdict(list)
    for url in urls:
        host = get_host(url)
        html = get_img_html(url)
        htmls[host].append(html)
    for host in htmls:
        path = images_check_path(host)
        html = '\n'.join(sorted(htmls[host]))
        dump_text(html, path)


def load_images_check(path=CHECK_IMAGES):
    table = read_excel(path)
    for cell in table.url:
        parts = cell.split()
        url = parts[-1]
        if not url.startswith('http://'):
            url = 'http://' + url
        yield url


def show_html(lines):
    from IPython.display import HTML, display
    display(HTML('\n'.join(lines)))



def get_image_filename(url):
    _, extension = os.path.splitext(url)
    extension = extension.lower()
    assert extension in ('.jpg', '.jpeg', '.png', '.gif'), extension
    return hash_url(url) + extension


def get_image_path(url):
    return os.path.join(
        RAW_IMAGES_DIR,
        get_image_filename(url)
    )


def download_image(url):
    path = get_image_path(url)
    check_call(['wget', '--user-agent', 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.8; rv:21.0) Gecko/20100101 Firefox/21.0', url, '-O', path])


def download_images(urls):
    for url in urls:
        download_image(url)


def list_images_cache(path):
    for filename in os.listdir(path):
        if filename != '.DS_Store':
            yield filename
    

def list_raw_images_cache():
    return list_images_cache(RAW_IMAGES_DIR)


def list_thumb_images_cache():
    return list_images_cache(THUMB_IMAGE_DIR)


def get_thumb_path(filename):
    return os.path.join(
        THUMB_IMAGE_DIR,
        filename
    )


def get_raw_path(filename):
    return os.path.join(RAW_IMAGES_DIR, filename)


def resize_image(filename):
    source = get_raw_path(filename)
    target = get_thumb_path(filename)
    check_call(['convert', source, '-resize', '400x400', target])
    check_call(['convert', target, '-sharpen', '5', target])


def resize_images(filenames):
    for filename in filenames:
        resize_image(filename)


def get_image_size(path):
    from PIL import Image
    image = Image.open(path)
    return Size(*image.size)


def load_images(urls):
    for url in urls:
        filename = get_image_filename(url)
        raw = get_image_size(get_raw_path(filename))
        thumb = get_image_size(get_thumb_path(filename))
        yield Image(url, filename, raw, thumb)


def download_schoolotzyv_list(region):
    response = requests.post(
        'http://www.schoolotzyv.ru/index.php?option=com_schools&view=moscow&format=json',
        data={
            'coordinates': region
        }
    )
    return response.json()


def download_schoolotzyv_lists(regions):
    for region in regions:
        for item in download_schoolotzyv_list(region):
            yield item


def parse_schoolotzyv_school_page(url, html):
    soup = get_soup(html)
    link = soup.find('span', itemprop='url')
    site_url = None
    if link:
        site_url = 'http://' + link.text
    link = soup.find('a', text=u'Все отзывы')
    reviews_url = None
    if link:
        reviews_url = 'http://www.schoolotzyv.ru' + link['href']
    information = soup.find('span', text=u'Информация:')
    merge_url = None
    if information:
        text = information.next_sibling
        if u'присоед' in text:  # присоединили к, была присоединена 
            link = information.parent.find('a')
            merge_url = link['href']
    return SchoolotzyvSchoolRecord(url, site_url, reviews_url, merge_url)


def load_schoolotzyv_school_pages(urls):
    for url in urls:
        html = load_html(url)
        yield parse_schoolotzyv_school_page(url, html)


def parse_schoolotzyv_votes(soup):
    votes = soup.find('span', class_='scomments-vote')
    if votes:
        truth = votes.find('a', class_='scomments-vote-good')
        truth = truth.find('span')
        if truth:
            truth = int(truth.text)
        lie = votes.find('a', class_='scomments-vote-poor')
        lie = lie.find('span')
        if lie:
            lie = int(lie.text)
        return SchoolotzyvVotes(truth, lie)

    
def parse_datetime(value):
    return datetime.strptime(value[:10], '%Y-%m-%d')


def parse_schoolotzyv_reviews(url, html):
    soup = get_soup(html)
    for item in soup.find_all('div', class_='scomments-item'):
        votes = parse_schoolotzyv_votes(item)
        name = item.find('span', itemprop='author').text
        date = parse_datetime(item.find('span', class_='scomments-date').text)
        text = item.find('div', class_='scomments-text').text
        yield SchoolotzyvReviewRecord(url, name, date, votes, text)

        
def get_url_last_part(url):
    return url.split('/')[-1]
        

def load_raw_schoolotzyv_reviews(records):
    site_urls = defaultdict(list)
    reviews = defaultdict(list)
    for record in records:
        url = record.url
        site_url = record.site or url
        site_urls[site_url].append(url)
        reviews_url = record.reviews or url
        html = load_html(reviews_url)
        for review in parse_schoolotzyv_reviews(reviews_url, html):
            reviews[site_url].append(review)
    for site_url, records in reviews.iteritems():
        urls = site_urls[site_url]
        yield SchoolotzyvReviewsRecord(urls, site_url, records)


def dump_datetime(datetime):
    if datetime:
        return datetime.strftime('%Y-%m-%d')


def dump_schoolotzyv_reviews(records):
    data = []
    for urls, site, reviews in records:
        reviews = [
            (source, name, dump_datetime(date), votes, text)
            for source, name, date, votes, text in reviews
        ]
        data.append((urls, site, reviews))
    dump_json(data, SCHOOLOTZYV_REVIEWS)
    
    
def load_schoolotzyv_reviews():
    data = load_json(SCHOOLOTZYV_REVIEWS)
    for urls, site, reviews in data:
        reviews = [
            SchoolotzyvReviewRecord(
                source, name, parse_datetime(date),
                SchoolotzyvVotes(*votes) if votes else None,
                text
            )
            for source, name, date, votes, text in reviews
        ]
        yield SchoolotzyvReviewsRecord(urls, site, reviews)


def load_mel_school_urls():
    html = load_text(SEARCH_HTML)
    soup = get_soup(html)
    for item in soup.find_all('a', class_='b-school-list-item'):
        yield join_url('http://schools.mel.fm/', item['href'])


def parse_mel_votes(soup):
    header = soup.find('div', class_='b-comment__header')
    if header:
        sections = []
        for item in header.find_all('div', class_='b-comment__header-section'):
            stars = item.find_all('div', class_='b-stars__star_selected')
            sections.append(len(stars))
        if len(sections) == 4:
            # ignore case then just some sections were rated
            return MelVotesRecord(*sections)


def parse_mel_reviews(soup):
    for item in soup.find_all('div', class_='b-comment'):
        author = item.find('div', class_='b-comment__author').text.strip()
        votes = parse_mel_votes(item)
        text = (item.find('span', class_='b-comment__text_hidden')
                or item.find('span', class_='b-comment__text')).text
        yield MelReviewRecord(author, votes, text)
    
    
def parse_mel_school(url, html):
    soup = get_soup(html)
    site = (soup.find('a', text=u'Сайт школы')
            or soup.find('a', text=u'Лицей на сайте Департамента образования Москвы')
            or soup.find('a', text=u'Школа на сайте Департамента образования Москвы')
            or soup.find('a', text=u'Гимназия на сайте Департамента образования Москвы'))
    if site:
        site = site['href']
    reviews = list(parse_mel_reviews(soup))
    return MelReviewsRecords(url, site, reviews)



def load_raw_mel_reviews(urls):
    for url in urls:
        html = load_html(url)
        yield parse_mel_school(url, html)


def dump_mel_reviews(records):
    dump_json(records, MEL_REVIEWS)
    
    
def load_mel_reviews():
    data = load_json(MEL_REVIEWS)
    for url, site, reviews in data:
        reviews = [
            MelReviewRecord(
                author,
                MelVotesRecord(*votes) if votes else None,
                text
            )
            for author, votes, text in reviews
        ]
        yield MelReviewsRecords(url, site, reviews)


def dump_reviews_check(schoolotzyv_reviews, mel_reviews, ratings):
    host_review = {get_host(_.site): _ for _ in schoolotzyv_reviews}
    url_review = {_.site: _ for _ in schoolotzyv_reviews}
    host_mel_review = {get_host(_.site): _ for _ in mel_reviews if _.site}
    
    data = []
    for record in ratings:
        host = get_host(record.url)
        reviews = []
        if host in host_review:
            reviews = host_review[host].reviews
        elif host in SCHOOLOTZYV_SITE_ALIASES:
            url = SCHOOLOTZYV_SITE_ALIASES[host]
            reviews = url_review[url].reviews
        for record in reviews:
            date = record.date
            if date.year > 2013:
                url = record.url
                votes = record.votes
                if votes:
                    votes = ', '.join(str(_) for _ in votes if _)
                data.append((
                    host, url, date,
                    record.name, votes, None, record.text
                ))

        if host in host_mel_review:
            record = host_mel_review[host]
            url = record.url
            for review in record.reviews:
                date = parse_datetime('2016-04-01')
                votes = review.votes
                if votes:
                    votes = ', '.join(str(_) for _ in votes)
                pattern = ur'(Образование|Учителя|Инфраструктура|Атмосфера|Индивидуализация|Общее впечатление)'
                text = review.text
                for part in re.split(pattern, text):
                    if not re.match(pattern, part):
                        part = part.strip()
                        if part:
                            data.append((
                                host, url, date,
                                review.author, votes, None, part
                            ))

            
    table = pd.DataFrame(
        data,
        columns=['host', 'url', 'date', 'name', 'votes', 'check', 'text']
    )
    table.to_excel(REVIEWS_CHECK, index=False)


def load_reviews_check():
    table = read_excel(REVIEWS_CHECK)
    for _, row in table.iterrows():
        host, url, date, name, _, check, text = row
        if check == '+':
            yield CheckReviewRecord(host, url, date, name, text)
            

def get_imap(login=MAIL_LOGIN, password=MAIL_PASSWORD):
    assert password is not None, 'os.environ["AK_MAIL_PASSWORD"] = "XXX"'
    
    import imaplib

    imap = imaplib.IMAP4_SSL(
        'imap.yandex.ru',
        993
    )
    code, message = imap.login(login, password)
    assert code == 'OK', message
    return imap


def send_email(emails, subject, body, login=MAIL_LOGIN, password=MAIL_PASSWORD):
    assert password is not None, 'os.environ["AK_MAIL_PASSWORD"] = "XXX"'
    from smtplib import SMTP_SSL
    from email.mime.text import MIMEText 

    ak = login
    su = 'su@obr.msk.ru'

    email = MIMEText(body.encode('utf8'))
    email['Subject'] = subject.encode('utf8')
    email['From'] = ak
    email['To'] = ', '.join(emails)
    email['Cc'] = su

    smtp = SMTP_SSL()
    smtp.connect('smtp.yandex.ru')
    smtp.login(ak, password)
    smtp.sendmail(ak, emails + [su], email.as_string())
    smtp.quit()


def get_feedback_mail_indexes(imap):
    imap.select('Inbox', readonly=True)
    code, (data,) = imap.search(None, 'SUBJECT', 'New submission from obr.msk.ru/form.html')
    indexes = data.split()
    return indexes


def parse_feedback_payload(payload):
    parts = re.split(ur'This form was submitted at|(review|bug|feedback)-(\w+):', payload)
    parts = parts[1:-3]  # trip header and footer
    name = None
    sections = []
    types = set()
    for index, part in enumerate(parts):
        position = index % 3
        if position == 1:
            name = part
        elif position == 0:
            types.add(part)
        elif position == 2:
            text = part
            text = text.strip('\r\n')
            sections.append(
                FeedbackPart(name, text)
            )
    assert len(types) == 1
    type = types.pop()
    return type, sections
    

def parse_feedback_date(date):
    return datetime.strptime(date, '%a, %d %b %Y %H:%M:%S +0000')


def parse_feedback_message(index, body):
    import email

    message = email.message_from_string(body)
    date = parse_feedback_date(message['Date'])
    for part in message.walk():
        if part.get_content_type() == 'text/plain':
            payload = part.get_payload(decode=True)
            payload = payload.decode('utf8')
            type, sections = parse_feedback_payload(payload)
            return FeedbackRecord(index, date, type, sections)
            
            
def load_feedback_messages(imap, indexes):
    for index in indexes:
        code, [(header, body), some_stuff] = imap.fetch(index, '(RFC822)')
        yield parse_feedback_message(index, body)


def close_imap(imap):
    imap.close()
    imap.logout()


def load_universities_check():
    table = read_excel(UNIVERSITIES_CHECK)
    universities = defaultdict(list)
    for _, (url, university, count, source) in table.iterrows():
        record = CheckUniversitiesRecord(url, university, int(count), source)
        universities[url].append(record)
    return dict(universities)


def parse_eduoffice_report(id, table):
    sections = {}
    for index, row in table.iterrows():
        section = unicode(row[0])
        # 10 -- 2015 Q4
        if re.match('[\.\d]+', section):
            value = row[10]
            sections[section] = value
    pupils = EduofficeReportPupils(
        total=sections['1'],
        program_0=sections['1.1.'],
        program_1_4=sections['1.2.'],
        program_5_9=sections['1.3.'],
        program_10_11=sections['1.4.']
    )
    incoming = EduofficeReportIncoming(
        total=sections['2'],
        from_0=sections['2.1.']
    )
    teachers = EduofficeReportTeachers(
        total=sections['5'],
        main=sections['5.1.'],
        other_main=sections['5.2.'],
        administration=sections['5.3.'],
        other=sections['5.4.'],
    )
    salaries= EduofficeReportSalaries(
        total=sections['6'],
        teachers=sections['6.1.1.1.'],
        administration=sections['6.2.'],
        other=sections['6.3.'],
    )
    return EduofficeReportRecord(
        id, pupils, incoming, teachers, salaries
    )
      
    
def load_eduoffice_report(id):
    path = get_eduoffice_report_path(id)
    table = read_excel(path)
    return parse_eduoffice_report(id, table)


def load_eduoffice_reports(ids):
    for id in ids:
        yield load_eduoffice_report(id)


def get_host_inn(eduoffices, eduoffices_selection):
    hosts = {get_host(_.url) for _ in eduoffices_selection}
    host_inn = {}
    for record in eduoffices:
        url = record.url
        if url:
            host = get_host(url)
            if host in hosts:
                inn = record.inn
                host_inn[host] = inn
    return host_inn


def get_host_eduoffice_id(eduoffices, eduoffices_selection):
    hosts = {get_host(_.url) for _ in eduoffices_selection}
    host_id = {}
    for record in eduoffices:
        url = record.url
        if url:
            host = get_host(url)
            if host in hosts:
                host_id[host] = record.id
    return host_id


def dump_eduoffice_report_check(host_id, eduoffice_reports):
    mapping = {_.id: _ for _ in eduoffice_reports}
    data = []
    for host, id in host_id.iteritems():
        record = mapping[id]
        pupils = record.pupils.total
        teachers = record.teachers.main
        salaries = record.salaries.teachers
        if pupils and teachers and salaries:
            data.append((host, pupils, teachers, salaries))
    table = pd.DataFrame(
        data,
        columns=['host', 'pupils', 'teachers', 'salaries']
    )
    table.to_excel(EDUOFFICE_REPORT_CHECK, index=False)


def load_eduoffice_report_check():
    table = read_excel(EDUOFFICE_REPORT_CHECK)
    for index, row in table.iterrows():
        host, pupils, teachers, salaries, incoming_total, incoming_from_0 = row
        yield EduofficeReportCheckRecord(
            host, pupils, teachers, salaries,
            EduofficeReportIncoming(incoming_total, incoming_from_0)
        )


def get_bus_search_url(inn):
    return ('http://bus.gov.ru/public/agency/extendedSearchAgencyNew.json?action=&agency={inn}'
            '&documentTypes=A&okatoSubElements=false&orderAttributeName=rank&orderDirectionASC=false'
            '&page=1&pageSize=10&ppoSubElements=false&primaryActivitySubElements=false&searchTermCondition=and'
            '&secondaryActivitySubElements=false&vguSubElements=true&withBranches=true').format(
        inn=inn
    )


def get_bus_search_filename(inn):
    return '{inn}.json'.format(inn=inn)


def get_bus_search_path(inn):
    return os.path.join(
        BUS_SEARCH_DIR,
        get_bus_search_filename(inn)
    )


def fetch_bus_search(inn):
    url = get_bus_search_url(inn)
    data = download_json(url)
    path = get_bus_search_path(inn)
    dump_json(data, path)


def parse_bus_search_record(inn, data):
    agencies = data['agencies']
    # assume first result is best
    other = len(agencies) - 1
    agency = agencies[0]
    id = agency['agencyId']
    name = agency['shortName']
    url = agency['website']
    return BusSearchRecord(inn, id, url, name, other)


def load_bus_search_record(inn):
    path = get_bus_search_path(inn)
    data = load_json(path)
    return parse_bus_search_record(inn, data)


def get_bus_latest_report_url(id):
    return ('http://bus.gov.ru/public/agency/last-annual-balance-F0503721-info.json'
            '?agencyId={id}'.format(id=id))


def get_bus_latest_report_filename(id):
    return '{id}.json'.format(id=id)


def get_bus_latest_report_path(id):
    return os.path.join(
        BUS_REPORT_DIR,
        get_bus_latest_report_filename(id)
    )


def fetch_latest_bus_report(id):
    url = get_bus_latest_report_url(id)
    data = download_json(url)
    path = get_bus_latest_report_path(id)
    dump_json(data, path)


def parse_bus_report_years(id, data):
    years = {
        _['financialYear']: _['id']
        for _ in data[u'formationPeriods']
    }
    return BusReportYears(id, year_2015=years.get(2015))


def load_bus_report_years(id):
    path = get_bus_latest_report_path(id)
    data = load_json(path)
    return parse_bus_report_years(id, data)


def get_bus_report_url(id, year):
    return ('http://bus.gov.ru/public/agency/annual-balance-F0503721-info.json'
            '?agencyId={id}&annualBalanceId={year}'.format(
            id=id,
            year=year
        ))


def get_bus_report_filename(id, year):
    return '{id}_{year}.json'.format(id=id, year=year)


def get_bus_report_path(id, year):
    return os.path.join(
        BUS_REPORT_DIR,
        get_bus_report_filename(id, year)
    )


def fetch_bus_report(id, year):
    url = get_bus_report_url(id, year)
    data = download_json(url)
    path = get_bus_report_path(id, year)
    dump_json(data, path)


def parse_bus_float(value):
    if value:
        # 554,404,818.06
        return float(value.replace(',', ''))

    
def parse_bus_report(id, data):
    assert data['feedbackFinancialYear'] == 2015
    balance = data['annualBalance']
    sections = {}
    for record in balance['incomings']:
        section = record['lineCode']
        value = parse_bus_float(record['totalEndYear'])
        sections[section] = value
    incomings = BusReportIncomingsRecord(
        total=sections['010'],
        paid=sections['040'],
        subsidii=sections['101']
    )
    sections = {}
    for record in balance['expenses']:
        section = record['lineCode']
        value = parse_bus_float(record['totalEndYear'])
        sections[section] = value
    expenses = BusReportExpensesRecord(
        total=sections['150'],
        salaries=sections['160'],
        gkh=sections['173']
    )
    return BusReportRecord(id, incomings, expenses)
    
    
def load_bus_report(id, year):
    path = get_bus_report_path(id, year)
    data = load_json(path)
    return parse_bus_report(id, data)


def get_url_review_urls(urls):
    url_review_urls = {}
    for url in urls:
        html = load_html(url)
        soup = get_soup(html)
        link = (soup.find('a', text=u'Отзывы об учреждении')
                or soup.find('a', text=u'Отзывы о комплексе'))
        if link:
            review_url = join_url(url, link['href'])
            url_review_urls[url] = review_url
    return url_review_urls


def parse_mskobr_review_meta(meta):
    match = re.search(ur'^(\d+) (\w+) (\d+) в [:\d]+\s+(.+)$', meta, re.U | re.M)
    if match:
        day, month, year, author = match.groups()
        day = int(day)
        month = MSKOBR_REVIEW_MONTHS[month]
        year = int(year)
        date = datetime(day=day, month=month, year=year)
        return date, author
    return None, None

    
def load_mskobr_reviews_page(url):
    html = load_html(url)
    soup = get_soup(html)
    for item in soup.find_all('div', class_='kris-ques-otziv'):
        text = item.find('div', class_='kris-quesman-bot')
        if text:
            text = text.text
        meta = item.find('div', class_='kris-quesman-dt-name')
        date = None
        author = None
        if meta:
            date, author = parse_mskobr_review_meta(meta.text)
        yield MskobrReviewRecord(date, author, text)

        
def change_mskobr_review_page(url, page):
    return '{url}?p={page}'.format(url=url, page=page)


def load_raw_mskobr_reviews(urls):
    for review_url in urls:
        reviews = []
        for page in xrange(0, 1000):
            url = change_mskobr_review_page(review_url, page=page)
            if url not in cache:
                fetch_url(url)
            records = list(load_mskobr_reviews_page(url))
            if not records:
                break
            reviews.extend(records)
        yield MskobrReviewsRecord(review_url, reviews)


def dump_mskobr_reviews(mskobr_reviews):
    data = []
    for url, reviews in mskobr_reviews:
        reviews = [
            (dump_datetime(_.date), _.author, _.text)
            for _ in reviews
        ]
        data.append((url, reviews))
    dump_json(data, MSKOBR_REVIEWS)


def load_mskobr_reviews():
    data = load_json(MSKOBR_REVIEWS)
    for url, reviews in data:
        reviews = [
            MskobrReviewRecord(
                parse_datetime(date) if date else None,
                author, text
            )
            for date, author, text in reviews
        ]
        yield MskobrReviewsRecord(url, reviews)


def parse_teacher_name(name):
    match = re.search(ur'^\s*(\w+)[^\w]+(\w+)[^\w]+(\w+)', name, re.U)
    if match:
        last, first, middle = match.groups()
        return Name(last, first, middle)


def parse_algfio_tomita_facts(text, xml):
    tree = ET.fromstring(xml.encode('utf8'))
    document = tree.find('document')
    if document is None:
        return
    facts = document.find('facts')
    for item in facts.findall('Person'):
        start = int(item.get('pos'))
        size = int(item.get('len'))
        substing = text[start:start + size]
        last = item.find('Name_Surname')
        if last is not None:
            last = last.get('val') or None
        first = item.find('Name_FirstName')
        if first is not None:
            first = first.get('val')
        middle = item.find('Name_Patronymic')
        if middle is not None:
            middle = middle.get('val')
        known_surname = item.find('Name_SurnameIsDictionary')
        if known_surname is not None:
            known_surname = int(known_surname.get('val'))
        known_surname = bool(known_surname)
        yield AlgfioTomitaFact(
            start, size, substing, last, first, middle, known_surname
        )
        
        
def run_algfio_tomita(text):
    dump_text(text, ALGFIO_TEXT)
    bin = os.path.relpath(TOMITA_BIN, ALGFIO_DIR)
    config = os.path.relpath(ALGFIO_CONFIG, ALGFIO_DIR)
    check_call([bin, config], cwd=ALGFIO_DIR)
    xml = load_text(ALGFIO_FACTS)
    for record in parse_algfio_tomita_facts(text, xml):
        yield record


def br(text):
    return text.replace('\n', '<br/>')


def display_tomite_facts(text, facts):
    from IPython.display import HTML, display

    CLOSING = 'closing'
    OPENING = 'opening'
    COLOR = '#ffffc2'

    tags = defaultdict(list)
    for fact in facts:
        start = fact.start
        stop = start + fact.size
        tags[start].append((OPENING, COLOR))
        tags[stop].append((CLOSING, None))
    chunks = []
    previous = 0
    for index in sorted(tags):
        chunk = br(text[previous:index])
        chunks.append(chunk)
        previous = index
        for tag, color in tags[index]:
            if tag == OPENING:
                chunk = ('<span style="background-color:{color}">'.format(
                    color=color
                ))
            elif tag == CLOSING:
                chunk = '</span>'
            chunks.append(chunk)
    if tags:
        chunks.append(br(text[index:]))
    else:
        chunks.append(br(text))
    html = ''.join(chunks)
    display(HTML(html))


class MorphNormalizer(object):
    morph = None

    def __init__(self):
        import pymorphy2
        self.morph = pymorphy2.MorphAnalyzer()
        
    def __call__(self, string):
        forms = self.morph.normal_forms(string)
        form = forms[0]
        return form

    
class StemNormalizer(object):
    stemmer = None

    def __init__(self):
        from nltk.stem.snowball import SnowballStemmer
        self.stemmer = SnowballStemmer('russian')
        
    def __call__(self, string):
        return self.stemmer.stem(string)


def handle_yo(string):
    return string.replace(u'ё', u'е')


def normalize_first_name(first):
    if first == u'наталия':
        return u'наталья'
    return first


def build_normal_name(last, first, middle, normalize_word):
    if last:
        last = handle_yo(normalize_word(last)).capitalize()
    if first:
        first = first.lower()
        first = normalize_first_name(
            handle_yo(first)
        ).capitalize()
    if middle:
        middle = middle.lower()
        middle = handle_yo(middle).capitalize()
    return Name(last, first, middle)


def get_host_teacher_normal_names(teachers, normalize_word):
    host_teachers = defaultdict(dict)
    for record in teachers:
        url = record.url
        host = get_host(url)
        name = record.name
        if name:
            # some manually added record lack name
            name = parse_teacher_name(name)
            if name:
                name = build_normal_name(
                    name.last, name.first, name.middle,
                    normalize_word
                )
                host_teachers[host][url] = name
    return host_teachers


hash_text = hash_item


def get_algfio_filename(text):
    return '{hash}.json'.format(
        hash=hash_text(text)
    )


def get_algfio_path(text):
    return os.path.join(
        ALGFIO_DATA_DIR,
        get_algfio_filename(text)
    )


def list_algfio_cache():
    with open(ALGFIO_LIST) as file:
        for line in file:
            yield line.strip()
            
            
def update_algfio_cache(text):
    hash = get_algfio_filename(text)
    with open(ALGFIO_LIST, 'a') as file:
        file.write(hash + '\n')
        
        
def dump_algfio_facts(text, facts):
    path = get_algfio_path(text)
    dump_json(facts, path)
        
        

def load_algfio_facts(text):
    path = get_algfio_path(text)
    data = load_json(path)
    for item in data:
        yield AlgfioTomitaFact(*item)
    

def convert_algfio_facts(text):
    facts = list(run_algfio_tomita(text))
    dump_algfio_facts(text, facts)
    update_algfio_cache(text)
    

def display_teacher_mentions(text, mentions):
    from IPython.display import HTML, display

    CLOSING = 'closing'
    OPENING = 'opening'
    COLOR = '#ffffc2'

    tags = defaultdict(list)
    for mention in mentions:
        start = mention.start
        stop = start + mention.size
        name = mention.name
        teacher = mention.teacher
        url = None
        if teacher:
            url = teacher.url
        tags[start].append((OPENING, name, url))
        tags[stop].append((CLOSING, name, url))
    chunks = []
    previous = 0
    for index in sorted(tags):
        chunk = br(text[previous:index])
        chunks.append(chunk)
        previous = index
        for tag, name, url in tags[index]:
            if tag == OPENING:
                if url:
                    chunk = (
                        '<span style="background-color:{color}">'
                        '<a href="{url}">'.format(
                            color=COLOR,
                            url=url
                    ))
                else:
                    chunk = (
                        '<span style="background-color:{color}">'.format(
                            color=COLOR
                    ))
            elif tag == CLOSING:
                name = u'({last} {first} {middle})'.format(
                    last=name.last or '_',
                    first=name.first or '_',
                    middle=name.middle or '_'
                )
                if url:
                    chunk = u'</a>{name}</span>'.format(
                        name=name
                    )
                else:
                    chunk = u'{name}</span>'.format(
                        name=name
                    )
            chunks.append(chunk)
    if tags:
        chunks.append(br(text[index:]))
    else:
        chunks.append(br(text))
    html = ''.join(chunks)
    display(HTML(html))


def lookup_teacher_name(name, host, host_teachers, url_teachers, normalize_last=None):
    teachers = host_teachers.get(host)
    if not teachers:
        return
    if normalize_last:
        name = build_normal_name(
            name.last, name.first, name.middle,
            normalize_last
        )
    first = name.first
    middle = name.middle
    if first and middle:
        last = name.last
        if last is not None:
            if len(first) == 1 and len(middle) == 1:
                for url, name in teachers.iteritems():
                    if name.last == last and name.first[0] == first and name.middle[0] == middle:
                        yield url_teachers[url]
            else:
                for url, name in teachers.iteritems():
                    if name.last == last and name.first == first and name.middle == middle:
                        yield url_teachers[url]
        else:
            for url, name in teachers.iteritems():
                if name.first == first and name.middle == middle:
                    yield url_teachers[url]

           
def get_robust_sad_teacher_mentions(facts, host, host_sad_teachers, host_teachers,
                                    url_teachers, normalize_word):
    for fact in facts:
        sad_matches = list(lookup_teacher_name(
            fact, host, host_sad_teachers, url_teachers, normalize_word
        ))
        matches = list(lookup_teacher_name(
            fact, host, host_teachers, url_teachers, normalize_word
        ))
        if sad_matches and len(sad_matches) + len(matches) == 1:
            teacher = sad_matches[0]
            name = build_normal_name(
                fact.last, fact.first, fact.middle, normalize_word
            )
            yield TeacherMention(
                fact.start, fact.size, fact.substring,
                name, teacher
            )
 
    
def format_algfio_text(text, facts, normalize_word):
    CLOSING = 'closing'
    OPENING = 'opening'

    tags = defaultdict(list)
    for fact in facts:
        start = fact.start
        stop = start + fact.size
        name = build_normal_name(
            fact.last,
            fact.first,
            fact.middle,
            normalize_word
        )
        if name.first and fact.middle:
            tags[start].append((OPENING, None))
            tags[stop].append((CLOSING, name))
    chunks = []
    previous = 0
    for index in sorted(tags):
        chunk = text[previous:index]
        chunks.append(chunk)
        previous = index
        for tag, name in tags[index]:
            if tag == OPENING:
                chunk = '['
            elif tag == CLOSING:
                chunk = u']({last} {first} {middle})'.format(
                    last=name.last or '_',
                    first=name.first or '_',
                    middle=name.middle or '_'
                )
            chunks.append(chunk)
    if tags:
        chunks.append(text[index:])
    else:
        chunks.append(text)
    text = ''.join(chunks)
    text = text.replace('\r', '').strip()
    return text


def format_sad_algfio_reviews_check(data, normalize_word):
    yield '<reviews>'
    for url, review, facts in data:
        yield '  <review>'
        yield '    <url>' + url + '</url>'
        author = review.author
        if author:
            yield '    <author>' + author + '</author>'
        date = review.date
        if date:
            yield '    <date>' + dump_datetime(date) + '</date>'
        content = format_algfio_text(review.text, facts, normalize_word)
        content = content.replace('<', '&lt;').replace('>', '&gt;').replace('&', '&amp;')
        yield '    <text>' + content + '</text>'
        yield '  </review>'
    yield '</reviews>'
    
    
def dump_sad_algfio_reviews_check(mskobr_reviews, host_sad_teachers, host_teachers, url_teachers, normalize_word):
    data = []
    for url, reviews in mskobr_reviews:
        host = get_host(url)
        for review in sorted(
                reviews, key=lambda _: _.date or datetime(1990, 1, 1),
                reverse=True
        ):
            text = review.text
            if text is None:
                continue
            facts = list(load_algfio_facts(text))
            mentions = list(get_robust_sad_teacher_mentions(
                facts, host, host_sad_teachers, host_teachers, url_teachers,
                normalize_word
            ))
            if mentions:
                data.append((url, review, facts))

    with open(CHECK_SAD_ALGFIO_REVIEWS, 'w') as file:
        for line in format_sad_algfio_reviews_check(data, normalize_word):
            file.write(line.encode('utf8') + '\n')


def parse_algfio_text(content):
    shift = 0
    text = ''
    previous = 0
    facts = []
    for match in re.finditer(r'\[(.+?)\]\(([_\w]+) ([_\w]+) ([_\w]+)\)', content, re.U | re.S):
        substring, last, first, middle = match.groups()
        if last == '_':
            last = None
        if first == '_':
            first = None
        if middle == '_':
            middle = None
            
        start = match.start()
        end = match.end()
        text += content[previous:start]
        start = len(text)
        text += substring
        previous = end
        
        size = len(substring)
        name = Name(last, first, middle)
        facts.append(AlgfioFact(
            start, size, name
        ))
    text += content[previous:]
    return text, facts


def load_algfio_reviews_check(path):
    tree = ET.parse(path)
    reviews = tree.getroot()
    for review in reviews.findall('review'):
        url = review.find('url').text
        author = review.find('author')
        if author is not None:
            author = author.text
        date = review.find('date')
        if date is not None:
            date = parse_datetime(date.text)
        text, facts = parse_algfio_text(review.find('text').text)
        yield AlgfioReviewCheckRecord(url, author, date, text, facts)


def dump_teachers_staff_check(staff, teachers):
    data = []
    urls = {_.url for _ in teachers}
    for record in staff:
        url = record.url
        if url not in urls:
            data.append((
                url, record.name, record.position
            ))
    table = pd.DataFrame(data, columns=['url', 'name', 'position'])
    table.to_excel(TEACHERS_STAFF_CHECK, index=False)


def load_teachers_staff_check(path):
    table = read_excel(path)
    for _, row in table.iterrows():
        yield CheckTeachersStaffRecord(*row)


def load_sad_teachers_check(path):
    table = read_excel(path)
    for _, row in table.iterrows():
        yield CheckSadTeachersRecord(*row)


def dump_sad_teachers_check(teachers):
    data = []
    for record in teachers:
        data.append((
            record.url, record.name, record.position
        ))
    table = pd.DataFrame(data, columns=['url', 'name', 'position'])
    table.to_excel(SAD_TEACHERS_CHECK, index=False)


def dump_sads_check(eduoffices_selection, eduoffices):
    hosts = set()
    for record in eduoffices:
        url = record.url
        if url and u'дошкольное образование' in record.programs:
            host = get_host(url)
            hosts.add(host)
    data = []
    for record in eduoffices_selection:
        url = record.url
        host = get_host(url)
        if host in hosts:
            name = record.short
            if name.startswith(u'Школа'):
                name = name.replace(u'Школа', u'Дошкольное отделение школы')
            elif name.startswith(u'Лицей'):
                name = name.replace(u'Лицей', u'Дошкольное отделение лицея')
            elif name.startswith(u'Гимназия'):
                name = name.replace(u'Гимназия', u'Дошкольное отделение гимназии')
            elif name.startswith(u'Центр образования'):
                name = name.replace(u'Центр образования', u'Дошкольное отделение центра образования')
            elif name.startswith(u'Комплекс'):
                name = name.replace(u'Комплекс', u'Дошкольное отделение комплекса')
            data.append((url, name))
    table = pd.DataFrame(data, columns=['url', 'name'])
    table.to_excel(SAD_CHECKS, index=False)
        

def load_sads_check():
    table = read_excel(SAD_CHECKS)
    for _, row in table.iterrows():
        yield SadCheckRecord(*row)


def dump_sad_addresses_check(sad_selection, url_menus, eduoffices):
    mapping = {_.url: _ for _ in eduoffices}
    data = []
    for record in sad_selection:
        url = record.url
        for item in url_menus[url]:
            if item.program == u'Дошкольное':
                address = item.address
                if address in addresses_cache:
                    address = addresses_cache[address]
                else:
                    address = u'Москва ' + address
                    if address in addresses_cache:
                        address = addresses_cache[address]
                    else:
                        # these are ok to skip actually
                        address = None
                if address:
                    coordinates = address.coordinates
                    data.append((
                        '0..1', address.description,
                        coordinates.latitude, coordinates.longitude,
                        url
                        ))
        if not addresses:
            eduoffice = mapping[url]
            address = u'Москва ' + eduoffice.main_address.description
            address = addresses_cache[address]
            coordinates = address.coordinates
            data.append((
                '0..1', address.description,
                coordinates.latitude, coordinates.longitude,
                url
            ))
    table = pd.DataFrame(
        data,
        columns=['program', 'address', 'latitude', 'longitude', 'url']
    )
    table.to_excel(SAD_ADDRESSES_CHECK, index=False)


def get_sad_teacher_mentions(facts, host, host_sad_teachers, host_teachers,
                             url_teachers):
    for fact in facts:
        name = fact.name
        start = fact.start
        size = fact.size
        
        matches = list(lookup_teacher_name(
            name, host, host_sad_teachers, url_teachers
        ))
        exact = len(matches) == 1
        for teacher in matches:
            yield SadTeacherMention(start, size, teacher, exact, True)
        if not matches:
            matches = list(lookup_teacher_name(
                name, host, host_teachers, url_teachers
            ))
            exact = len(matches) == 1
            for teacher in matches:
                yield SadTeacherMention(start, size, teacher, exact, False)
            

def get_sad_teachers_reviews(sad_algfio_reviews, host_sad_teachers, host_teachers,
                             url_teachers):
    for url, author, date, text, facts in sad_algfio_reviews:
        host = get_host(url)
        mentions = list(get_sad_teacher_mentions(
            facts, host, host_sad_teachers, host_teachers, url_teachers
        ))
        yield AlgfioReviewCheckRecord(
            url, author, date, text,
            mentions
        )

def format_sad_review_html(text, facts):
    CLOSING = 'closing'
    OPENING = 'opening'

    tags = defaultdict(list)
    for fact in facts:
        start = fact.start
        stop = start + fact.size
        url = fact.teacher.url
        # TODO same start may refer to different teachers
        tags[start] = (OPENING, url)
        tags[stop] = (CLOSING, None)
    chunks = []
    previous = 0
    for index in sorted(tags):
        chunk = br(text[previous:index])
        chunks.append(chunk)
        previous = index
        tag, url = tags[index]
        if tag == OPENING:
            chunk = ('<a href="{url}">'.format(
                url=url
            ))
        elif tag == CLOSING:
            chunk = '</a>'
        chunks.append(chunk)
    if tags:
        chunks.append(br(text[index:]))
    else:
        chunks.append(br(text))
    html = ''.join(chunks)
    return html


def get_teacher_mentions_sad_reviews(reviews, url_photos):
    teacher_url_mentions = Counter()
    teacher_url_samples = {}
    url_teachers = {}
    sad_reviews = []
    reviews = sorted(
        reviews, key=lambda _: _.date or datetime(1990, 1, 1),
        reverse=True
    )
    for review in reviews:
        sample = False
        text = review.text
        id = hash_text(text)
        facts = review.facts
        for fact in facts:
            if fact.sad and fact.exact:
                teacher = fact.teacher
                url = teacher.url
                if url not in url_teachers:
                    sample = True
                    teacher_url_samples[url] = id
                    url_teachers[url] = teacher
                teacher_url_mentions[url] += 1
        if sample:
            sad_reviews.append(VizSadReview(
                id, review.author, review.date,
                format_sad_review_html(text, facts)
            ))
    teacher_mentions = []
    for url in url_teachers:
        teacher = url_teachers[url]
        mentions = teacher_url_mentions[url]
        sample = teacher_url_samples[url]
        name = parse_teacher_name(teacher.name)
        name = build_normal_name(
            name.last, name.first, name.middle,
            lambda _: _
        )
        image = url_photos.get(url)
        position = parse_sad_teacher_position(teacher.position)
        teacher_mentions.append(VizTeacherMentions(
            url, name, image, position, mentions, sample
        ))
    return teacher_mentions, sad_reviews


def get_viz_sads(sad_selection, schools, sad_addresses, sad_review_mentions,
                 url_photos, sad_images, eduoffice_reports):
    host_schools = {get_host(_.url): _ for _ in schools}

    host_addresses = defaultdict(list)
    for record in sad_addresses:
        host = get_host(record.url)
        address = record.address
        coordinates = address.coordinates
        address = VizSadAddress(
            address.description,
            coordinates.latitude, coordinates.longitude 
        )
        host_addresses[host].append(address)
    
    host_reviews = defaultdict(list)
    for record in sad_review_mentions:
        host = get_host(record.url)
        host_reviews[host].append(record)

    host_images = defaultdict(list)
    for record in sad_images:
        host = get_host(record.url)
        host_images[host].append(record)

    host_incoming = {}
    for record in eduoffice_reports:
        host_incoming[record.host] = record.incoming

    for record in sad_selection:
        url = record.url
        full = record.name
        host = get_host(url)
        school = host_schools[host]
        link = school.title.link + '-sad'
        title = VizSadTitle(full, link)
        contacts = school.contacts
        addresses = host_addresses[host]
        reviews = host_reviews[host]
        teacher_mentions, sad_reviews = get_teacher_mentions_sad_reviews(
            reviews,
            url_photos
        )
        images = host_images[host]
        incoming = host_incoming.get(host)
        yield VizSad(
            url, title, school, addresses, contacts,
            teacher_mentions, sad_reviews,
            images, incoming
        )


def get_sad_page_filename(sad):
    return '{link}.html'.format(
        link=sad.title.link
    )
    

def get_sad_page_path(sad):
    return os.path.join(
        SITE_DIR,
        get_sad_page_filename(sad)
    )


def make_table(sequence, columns):
    size = len(sequence)
    rows = int(ceil(float(size) / columns))
    table = []
    for row in xrange(rows):
        table.append([None for _ in xrange(columns)])
    for index, item in enumerate(sequence):
        row = index / columns
        column = index % columns
        table[row][column] = item
    return table


def generate_sad_page(sad, template):
    path = get_sad_page_path(sad)
    sad_url = sad.url
    host = get_host(sad_url)
    title = sad.title
    page = get_sad_page_filename(sad)
    addresses = sad.addresses
    contacts = sad.contacts
    
    teacher_mentions = sad.teacher_mentions
    if teacher_mentions:
        teacher_mentions = sorted(
            teacher_mentions,
            key=lambda _: _.mentions, reverse=True
        )
        teacher_mentions = [
            VizTeacherMentions(
                url,
                u'{last} {first} {middle}'.format(
                    last=name.last,
                    first=name.first,
                    middle=name.middle
                ),
                image,
                position, mentions, sample
            )
            for url, name, image, position, mentions, sample in teacher_mentions
        ]
        teacher_mentions = make_table(teacher_mentions, columns=4)
        
    reviews = [
        VizSadReview(
            id, author,
            format_review_date(date) if date else None,
            html
        )
        for id, author, date, html in sad.reviews
    ]

    images = [
        (_.url, _.filename, _.thumb.width, _.thumb.height)
        for _ in sad.images
    ]

    incoming_share = None
    incoming = sad.incoming
    if incoming:
        total, from_0 = incoming
        if total > 0:
            incoming_share = int(float(from_0) / total * 100)


    school = sad.school
    school_url = school.title.link + '.html'
    school_no_title = school.title.no

    html = template.render(
        full_title=title.full,
        path=page,
        link=title.link,
        addresses=addresses,
        contacts=contacts,
        sad_url=sad_url,
        host=host,
        teacher_mentions=teacher_mentions,
        reviews=reviews,
        images=images,

        incoming_share=incoming_share,
        school_url=school_url,
        school_no_title=school_no_title
    )
    dump_text(html, path)

    
def format_mentions(mentions):
    if mentions % 10 == 1 and mentions != 11:
        return u'{mentions} благодарность'.format(
            mentions=mentions
        )
    elif mentions % 10 in (2, 3, 4) and mentions not in (12, 13, 14):
        return u'{mentions} благодарности'.format(
            mentions=mentions
        )
    else:
        return u'{mentions} благодарностей'.format(
            mentions=mentions
        )

    
def generate_sad_pages(sads):
    template = load_text(SAD_TEMPLATE)
    env = Environment()
    env.filters['format_mentions'] = format_mentions
    template = env.from_string(template)
    for sad in sads:
        generate_sad_page(sad, template)


def load_haar_cascade():
    return cv2.CascadeClassifier(HAAR_CASCADE)


def get_face(image, cascade):
    gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
    boxes = cascade.detectMultiScale(gray)
    if boxes != ():
        count, _ = boxes.shape
        if count == 1:
            x, y, width, height = boxes[0]
            assert width == height
            size = width
            height, width, _ = image.shape
            shift = min([
                size,
                x, y,
                height - size - y,
                width - size - x
            ])

            image = image[
                y - shift:y + size + shift,
                x - shift:x + size + shift
            ]
            image = cv2.resize(image, (150, 150))
            return image
    

def get_image_base64(image):
    from io import BytesIO
    from PIL import Image as PILImage

    image = PILImage.fromarray(image)
    bytes = BytesIO()
    image.save(bytes, format='png')
    return bytes.getvalue().encode('base64')


def format_img(image):
    return ('<img src="data:image/png;base64,{data}" '
            'style="display:inline"/>').format(
        data=get_image_base64(image)
    )


def display_image(image):
    from IPython.display import HTML, display

    html = format_img(image)
    display(HTML(html))

    
def load_image_data(path):
    image = cv2.imread(path)
    try:
        image = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
    except:
        return None
    return image


def get_photo_path(filename):
    return os.path.join(PHOTO_DIR, filename)


def dump_image(image, path):
    image = cv2.cvtColor(image, cv2.COLOR_RGB2BGR)
    cv2.imwrite(path, image)


def list_photos_cache():
    return list_images_cache(PHOTO_DIR)
    
   
def convert_photos(filenames):
    cascade = load_haar_cascade()
    for filename in log_progress(every=10):
        path = get_raw_path(filename)
        image = load_image_data(path)
        if image is not None:
            face = get_face(image, cascade)
            if face is not None:
                path = get_photo_path(filename)
                dump_image(face, path)


def load_url_photos(sad_teachers):
    url_photos = {}
    cache = set(list_photos_cache())
    for record in sad_teachers:
        url = record.image
        if url:
            filename = get_image_filename(url)
            if filename in cache:
                url_photos[record.url] = filename
    return url_photos


def scale_sharpen_dir(dir, filenames):
    for filename in filenames:
        path = os.path.join(dir, filename)
        check_call(['convert', path, '-resize', '100x100', path])
        check_call(['convert', path, '-sharpen', '5', path])
