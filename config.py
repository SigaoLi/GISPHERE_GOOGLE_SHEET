"""
配置文件 - 存储所有常量和配置信息
支持跨平台（Windows, macOS, Linux）
"""
import os
import platform
import pytz

# 系统信息
SYSTEM_PLATFORM = platform.system()  # 'Windows', 'Darwin' (macOS), 'Linux'

# Google Sheets API 配置
SCOPES_SHEETS = ['https://www.googleapis.com/auth/spreadsheets']
SCOPES_DOCS = ['https://www.googleapis.com/auth/documents']
SPREADSHEET_ID = '1LcfxcTCuj9ZJXXMxyFQwt-xnbAviNP8j9oDr6OG5-Go'
DOCUMENT_ID = '1PhNqalVi-5BWEiqANN4NAw26V3c-JcJpjrtQJVgenvY'

# 工作表ID
UNFILLED_SHEET_ID = 0

# 时区设置
CHINA_TZ = pytz.timezone('Asia/Shanghai')

# 文件路径配置（跨平台兼容）
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS_FILE = os.path.join(BASE_DIR, 'credentials.json')
TOKEN_PICKLE_FILE = os.path.join(BASE_DIR, 'token.pickle')
TOKEN_JSON_FILE = os.path.join(BASE_DIR, 'token.json')
EMAIL_CREDENTIALS_FILE = os.path.join(BASE_DIR, 'email_credentials.txt')
GROUP_MEMBERS_FILE = os.path.join(BASE_DIR, 'group_members.txt')
SQL_CREDENTIALS_FILE = os.path.join(BASE_DIR, 'sql_credentials.txt')

# SMTP配置
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# 国家字典
COUNTRY_DICTIONARY = {
    '几内亚': 'Guinea',
    '几内亚比绍': 'Guinea-Bissau',
    '土耳其': 'Turkey',
    '土库曼斯坦': 'Turkmenistan',
    '也门': 'Yemen',
    '马尔代夫': 'Maldives',
    '马耳他': 'Malta',
    '马达加斯加': 'Madagascar',
    '马来西亚': 'Malaysia',
    '马里': 'Mali',
    '马拉维': 'Malawi',
    '马绍尔群岛': 'Marshall Islands',
    '不丹': 'Bhutan',
    '厄瓜多尔': 'Ecuador',
    '厄立特里亚': 'Eritrea',
    '牙买加': 'Jamaica',
    '比利时': 'Belgium',
    '瓦努阿图': 'Vanuatu',
    '日本': 'Japan',
    '中国': 'China',
    '中非': 'Central African Republic',
    '冈比亚': 'Gambia',
    '贝宁': 'Benin',
    '毛里求斯': 'Mauritius',
    '毛里塔尼亚': 'Mauritania',
    '丹麦': 'Denmark',
    '乌干达': 'Uganda',
    '乌克兰': 'Ukraine',
    '乌拉圭': 'Uruguay',
    '乌兹别克斯坦': 'Uzbekistan',
    '文莱': 'Brunei',
    '巴巴多斯': 'Barbados',
    '巴布亚新几内亚': 'Papua New Guinea',
    '巴西': 'Brazil',
    '巴拉圭': 'Paraguay',
    '巴林': 'Bahrain',
    '巴哈马': 'Bahamas',
    '巴拿马': 'Panama',
    '巴勒斯坦': 'State of Palestine',
    '巴基斯坦': 'Pakistan',
    '以色列': 'Israel',
    '古巴': 'Cuba',
    '布基纳法索': 'Burkina Faso',
    '布隆迪': 'Burundi',
    '东帝汶': 'Timor-Leste',
    '卡塔尔': 'Qatar',
    '卢旺达': 'Rwanda',
    '卢森堡': 'Luxembourg',
    '乍得': 'Chad',
    '白俄罗斯': 'Belarus',
    '印度': 'India',
    '印度尼西亚': 'Indonesia',
    '立陶宛': 'Lithuania',
    '尼日尔': 'Niger',
    '尼日利亚': 'Nigeria',
    '尼加拉瓜': 'Nicaragua',
    '尼泊尔': 'Nepal',
    '加纳': 'Ghana',
    '加拿大': 'Canada',
    '加蓬': 'Gabon',
    '圣马力诺': 'San Marino',
    '圣文森特和格林纳丁斯': 'Saint Vincent and the Grenadines',
    '圣多美和普林西比': 'Sao Tome and Principe',
    '圣基茨和尼维斯': 'Saint Kitts and Nevis',
    '圣卢西亚': 'Saint Lucia',
    '圭亚那': 'Guyana',
    '吉布提': 'Djibouti',
    '吉尔吉斯斯坦': 'Kyrgyzstan',
    '老挝': 'Laos',
    '亚美尼亚': 'Armenia',
    '西班牙': 'Spain',
    '列支敦士登': 'Liechtenstein',
    '北马其顿共和国': 'Macedonia',
    '刚果民主共和国': 'Democratic Republic of the Congo',
    '刚果共和国': 'Republic of the Congo',
    '伊拉克': 'Iraq',
    '伊朗': 'Iran',
    '危地马拉': 'Guatemala',
    '匈牙利': 'Hungary',
    '多米尼加共和国': 'Dominican Republic',
    '多米尼克': 'Dominica',
    '多哥': 'Togo',
    '冰岛': 'Iceland',
    '汤加': 'Tonga',
    '安哥拉': 'Angola',
    '安提瓜和巴布达': 'Antigua and Barbuda',
    '安道尔': 'Andorra',
    '约旦': 'Jordan',
    '赤道几内亚': 'Equatorial Guinea',
    '芬兰': 'Finland',
    '克罗地亚': 'Croatia',
    '苏丹': 'Sudan',
    '苏里南': 'Suriname',
    '利比亚': 'Libya',
    '利比里亚': 'Liberia',
    '伯利兹': 'Belize',
    '佛得角': 'Cape Verde',
    '希腊': 'Greece',
    '沙特阿拉伯': 'Saudi Arabia',
    '阿尔及利亚': 'Algeria',
    '阿尔巴尼亚': 'Albania',
    '阿拉伯联合酋长国': 'United Arab Emirates',
    '阿根廷': 'Argentina',
    '阿曼': 'Oman',
    '阿富汗': 'Afghanistan',
    '阿塞拜疆': 'Azerbaijan',
    '纳米比亚': 'Namibia',
    '坦桑尼亚': 'Tanzania',
    '拉脱维亚': 'Latvia',
    '英国': 'United Kingdom',
    '肯尼亚': 'Kenya',
    '罗马尼亚': 'Romania',
    '帕劳': 'Palau',
    '图瓦卢': 'Tuvalu',
    '委内瑞拉': 'Venezuela',
    '所罗门群岛': 'Solomon Islands',
    '法国': 'France',
    '波兰': 'Poland',
    '波斯尼亚和黑塞哥维那': 'Bosnia and Herzegovina',
    '孟加拉国': 'Bangladesh',
    '玻利维亚': 'Bolivia',
    '挪威': 'Norway',
    '南苏丹共和国': 'The Republic of South Sudan',
    '南非': 'South Africa',
    '柬埔寨': 'Cambodia',
    '哈萨克斯坦': 'Kazakhstan',
    '科威特': 'Kuwait',
    '科特迪瓦': "Cote d'Ivoire",
    '科摩罗': 'Comoros',
    '保加利亚': 'Bulgaria',
    '俄罗斯': 'Russia',
    '叙利亚': 'Syria',
    '美国': 'United States of America',
    '洪都拉斯': 'Honduras',
    '津巴布韦': 'Zimbabwe',
    '突尼斯': 'Tunisia',
    '泰国': 'Thailand',
    '埃及': 'Egypt',
    '埃塞俄比亚': 'Ethiopia',
    '莱索托': 'Lesotho',
    '莫桑比克': 'Mozambique',
    '荷兰': 'Netherlands',
    '格林纳达': 'Grenada',
    '格鲁吉亚': 'Georgia',
    '索马里': 'Somalia',
    '哥伦比亚': 'Colombia',
    '哥斯达黎加': 'Costa Rica',
    '特立尼达和多巴哥': 'Trinidad and Tobago',
    '秘鲁': 'Peru',
    '爱尔兰': 'Ireland',
    '爱沙尼亚': 'Estonia',
    '海地': 'Haiti',
    '捷克': 'The Czech Republic',
    '基里巴斯': 'Kiribati',
    '菲律宾': 'Philippines',
    '萨尔瓦多': 'El Salvador',
    '萨摩亚': 'Samoa',
    '密克罗尼西亚联邦': 'The Federated States of Micronesia',
    '梵蒂冈': 'Vatican City State',
    '塔吉克斯坦': 'Tajikistan',
    '越南': 'Vietnam',
    '博茨瓦纳': 'Botswana',
    '斯里兰卡': 'Sri Lanka',
    '斯威士兰': 'Swaziland',
    '斯洛文尼亚': 'Slovenia',
    '斯洛伐克': 'Slovakia',
    '葡萄牙': 'Portugal',
    '韩国': 'South Korea',
    '朝鲜': 'North Korea',
    '斐济': 'Fiji',
    '喀麦隆': 'Cameroon',
    '黑山': 'Montenegro',
    '智利': 'Chile',
    '奥地利': 'Austria',
    '缅甸': 'Burma',
    '瑞士': 'Switzerland',
    '瑞典': 'Sweden',
    '瑙鲁': 'Nauru',
    '蒙古': 'Mongolia',
    '新加坡': 'Singapore',
    '新西兰': 'New Zealand',
    '意大利': 'Italy',
    '塞内加尔': 'Senegal',
    '塞尔维亚': 'Republic of Serbia',
    '塞舌尔': 'Republic of Seychelles',
    '塞拉利昂': 'Sierra Leone',
    '塞浦路斯': 'Cyprus',
    '墨西哥': 'Mexico',
    '黎巴嫩': 'Lebanon',
    '德国': 'Germany',
    '摩尔多瓦': 'Moldova',
    '摩纳哥': 'Monaco',
    '摩洛哥': 'Morocco',
    '澳大利亚': 'Australia',
    '赞比亚': 'Zambia',
    '纽埃': 'Niue',
    '库克群岛': 'Cook Islands'
}

# 职位字典
JOB_DICTIONARY = {
    "MSc": "硕士",
    "MA": "硕士",
    "PhD": "博士",
    "RA": "研究助理",
    "Master Student": "硕士研究生",
    "Doctoral Student": "博士研究生",
    "PostDoc": "博士后",
    "Research Assistant": "研究助理",
    "Competition": "竞赛",
    "Summer School": "暑期学校",
    "Conference": "学术会议",
    "Workshop": "研讨会"
}

# 学科字典
SUBJECT_DICTIONARY = {
    "Physical_Geo": "自然地理学",
    "Human_Geo": "人文地理学",
    "GIS": "地理信息科学",
    "Urban": "城市规划",
    "RS": "遥感",
    "GNSS": "测绘学"
}

# 标签列
LABEL_COLUMNS = ["Physical_Geo", "Human_Geo", "Urban", "GIS", "RS", "GNSS"]

# 必填列
REQUIRED_COLUMNS = ["Source", "Deadline", "Country_CN", "University_CN", "University_EN", "Direction"]

