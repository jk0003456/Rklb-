"""
롯데시네마 상영관 현황 엑셀 파일 생성기
이미지에서 확인된 형식을 그대로 재현
"""

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────────────────────
# 색상 정의
# ─────────────────────────────────────────────────────────────────────────────
CLR_ORANGE   = "FF6600"   # 수퍼MX4D
CLR_YELLOW   = "FFCC00"   # 샤롯대 / 수퍼LED / 아르떼
CLR_TEAL     = "00B0B0"   # 수퍼플렉스 / 수퍼4D
CLR_BLUE     = "4472C4"   # 씨네패밀리 / 씨네컴포트
CLR_DARK     = "404040"   # 광음시네마/공음LED
CLR_HEADER_BG = "D9D9D9" # 컬럼 헤더 배경
CLR_TITLE_BG  = "00B0F0" # 타이틀 배경
CLR_REGION_BG = "E2EFDA" # 지역 배경 (연한 녹색)
CLR_EVEN_ROW  = "F2F2F2" # 짝수 행
CLR_MX4D_CELL = "FFE0CC"  # MX4D 셀 배경
CLR_PLEX_CELL = "CCFFFF"  # 수퍼플렉스 셀 배경
CLR_4D_CELL   = "CCFFEE"  # 수퍼4D 셀 배경
CLR_CHARLOTTE_CELL = "FFFFCC"  # 샤롯대 셀 배경
CLR_LED_CELL  = "FFFACC"  # 수퍼LED 셀 배경
CLR_ARTE_CELL = "FFF0CC"  # 아르떼 셀 배경
CLR_FAMILY_CELL = "CCE0FF"  # 씨네패밀리 셀 배경
CLR_COMFORT_CELL = "DDEEFF"  # 씨네컴포트 셀 배경
CLR_GWAVE_CELL = "E0E0E0"  # 광음 셀 배경

# ─────────────────────────────────────────────────────────────────────────────
# 지점 데이터
# 형식: {
#   'region': str,        지역 (광역시/도)
#   'sub_region': str,    세부 지역 (시/구)
#   'name': str,          극장명
#   'type': str,          'direct'=지정(파란), 'branch'=야트(빨간), 'contract'=위탁(검정)
#   'screens': list,      각 스크린 좌석수 (None=빈칸)
#   'special': dict,      특수관 {screen_idx: type_code}  1-based
#   'notes': str,         비고 (개관일, 전화번호 등)
# }
# special type_code: 'mx4d','charlotte','led','arte','plex','4d','family','comfort','gwave'
# ─────────────────────────────────────────────────────────────────────────────

branches = [
    # ── 서울 ──────────────────────────────────────────────────────────────────
    {"region":"서울","sub_region":"서울 강남구","name":"강남",     "type":"direct",
     "screens":[229,153,153,118,118,118,118,148,148,148],"special":{},
     "notes":"T. 02-2055-8300"},
    {"region":"서울","sub_region":"서울 광진구","name":"건대입구", "type":"direct",
     "screens":[256,183,154,123,123,123,104,104],"special":{},"notes":"T. 02-457-8800"},
    {"region":"서울","sub_region":"서울 금천구","name":"가산디지털", "type":"direct",
     "screens":[358,249,165,165,113,113,113,82,82],"special":{5:"plex"},"notes":"T. 02-2107-8800"},
    {"region":"서울","sub_region":"서울 노원구","name":"노원",      "type":"direct",
     "screens":[313,208,147,147,131,131,102,102,78],"special":{1:"plex"},"notes":"T. 02-931-8800"},
    {"region":"서울","sub_region":"서울 도봉구","name":"창동",      "type":"direct",
     "screens":[267,184,149,110,110,82,82],"special":{},"notes":"T. 02-955-8800"},
    {"region":"서울","sub_region":"서울 동작구","name":"보라매",    "type":"direct",
     "screens":[228,183,155,124,124,96,96],"special":{},"notes":"T. 02-820-8800"},
    {"region":"서울","sub_region":"서울 마포구","name":"마포",      "type":"direct",
     "screens":[342,228,165,130,130,105,105,84],"special":{1:"plex"},"notes":"T. 02-325-8800"},
    {"region":"서울","sub_region":"서울 마포구","name":"홍대입구",  "type":"direct",
     "screens":[246,198,170,137,137,112,112,84,84],"special":{},"notes":"T. 02-330-8800"},
    {"region":"서울","sub_region":"서울 송파구","name":"월드타워",  "type":"direct",
     "screens":[634,317,276,261,261,261,220,170,170,170,138,138,138,128,128,128,116,116,116,114,114],
     "special":{1:"plex",4:"mx4d"},"notes":"T. 02-6002-6500"},
    {"region":"서울","sub_region":"서울 양천구","name":"목동",      "type":"direct",
     "screens":[313,208,166,138,138,138,105,105,105,80],"special":{1:"plex"},"notes":"T. 02-2606-8800"},
    {"region":"서울","sub_region":"서울 영등포구","name":"영등포",  "type":"direct",
     "screens":[508,290,198,198,145,145,130,110,110,84,84,84],"special":{1:"plex"},"notes":"T. 02-2637-8800"},
    {"region":"서울","sub_region":"서울 은평구","name":"은평",      "type":"direct",
     "screens":[326,229,179,146,146,120,120,92,92],"special":{1:"plex"},"notes":"T. 02-355-8800"},
    {"region":"서울","sub_region":"서울 중랑구","name":"상봉",      "type":"direct",
     "screens":[272,196,157,121,121,97,97,74],"special":{},"notes":"T. 02-493-8800"},

    # ── 경기 ──────────────────────────────────────────────────────────────────
    {"region":"경기","sub_region":"경기 고양시 일산동구","name":"주엽","type":"direct",
     "screens":[347,234,176,148,148,130,130,104],"special":{1:"plex"},"notes":"T. 031-919-8800"},
    {"region":"경기","sub_region":"경기 고양시 일산서구","name":"킨텍스","type":"direct",
     "screens":[441,294,210,164,164,130,130,104,104,82],"special":{1:"plex",3:"mx4d"},"notes":"T. 031-995-8800"},
    {"region":"경기","sub_region":"경기 고양시 덕양구","name":"화정","type":"direct",
     "screens":[260,174,140,112,112,86,86],"special":{},"notes":"T. 031-969-8800"},
    {"region":"경기","sub_region":"경기 광명시","name":"광명역",    "type":"direct",
     "screens":[474,316,222,175,175,140,140,112,112,86,86],"special":{1:"plex"},"notes":"T. 02-2680-8800"},
    {"region":"경기","sub_region":"경기 군포시","name":"산본",      "type":"direct",
     "screens":[264,176,142,112,112,90,90],"special":{},"notes":"T. 031-391-8800"},
    {"region":"경기","sub_region":"경기 구리시","name":"구리",      "type":"direct",
     "screens":[338,225,176,143,143,114,114,88,88],"special":{1:"plex"},"notes":"T. 031-555-8800"},
    {"region":"경기","sub_region":"경기 김포시","name":"김포",      "type":"direct",
     "screens":[310,207,163,128,128,104,104,78],"special":{},"notes":"T. 031-999-8800"},
    {"region":"경기","sub_region":"경기 부천시","name":"부천역",    "type":"direct",
     "screens":[330,220,175,140,140,114,114,86,86],"special":{1:"plex"},"notes":"T. 032-321-8800"},
    {"region":"경기","sub_region":"경기 성남시 분당구","name":"분당","type":"direct",
     "screens":[286,191,153,122,122,98,98,75],"special":{},"notes":"T. 031-781-8800"},
    {"region":"경기","sub_region":"경기 수원시 영통구","name":"광교","type":"direct",
     "screens":[487,325,228,180,180,144,144,116,116,88],"special":{1:"plex",3:"mx4d"},"notes":"T. 031-888-8800"},
    {"region":"경기","sub_region":"경기 수원시 팔달구","name":"수원",  "type":"direct",
     "screens":[392,261,208,163,163,130,130,103,103,80,80],"special":{1:"plex",3:"mx4d"},"notes":"T. 031-258-8800"},
    {"region":"경기","sub_region":"경기 시흥시","name":"시화",      "type":"direct",
     "screens":[304,203,163,129,129,103,103,80],"special":{},"notes":"T. 031-499-8800"},
    {"region":"경기","sub_region":"경기 안산시 단원구","name":"안산","type":"direct",
     "screens":[298,199,160,126,126,102,102,78],"special":{},"notes":"T. 031-437-8800"},
    {"region":"경기","sub_region":"경기 안성시","name":"안성",      "type":"contract",
     "screens":[168,112,90,72,72],"special":{},"notes":"T. 031-676-8800"},
    {"region":"경기","sub_region":"경기 안양시 동안구","name":"평촌","type":"direct",
     "screens":[312,208,166,132,132,106,106,80],"special":{1:"plex"},"notes":"T. 031-388-8800"},
    {"region":"경기","sub_region":"경기 의왕시","name":"의왕",      "type":"contract",
     "screens":[189,126,101,80,80,62],"special":{},"notes":"T. 031-455-8800"},
    {"region":"경기","sub_region":"경기 의정부시","name":"의정부",  "type":"direct",
     "screens":[354,236,188,148,148,120,120,92],"special":{1:"plex"},"notes":"T. 031-872-8800"},
    {"region":"경기","sub_region":"경기 파주시","name":"운정",      "type":"direct",
     "screens":[374,249,198,156,156,126,126,96,96],"special":{1:"plex"},"notes":"T. 031-947-8800"},
    {"region":"경기","sub_region":"경기 평택시","name":"평택",      "type":"direct",
     "screens":[286,191,153,122,122,98,98,74],"special":{},"notes":"T. 031-618-8800"},
    {"region":"경기","sub_region":"경기 화성시 동탄","name":"동탄", "type":"direct",
     "screens":[428,285,224,178,178,142,142,114,114,86],"special":{1:"plex",3:"mx4d"},"notes":"T. 031-8015-8800"},

    # ── 인천 ──────────────────────────────────────────────────────────────────
    {"region":"인천","sub_region":"인천 계양구","name":"계양",      "type":"direct",
     "screens":[284,189,152,120,120,96,96,74],"special":{},"notes":"T. 032-547-8800"},
    {"region":"인천","sub_region":"인천 남동구","name":"인천터미널","type":"direct",
     "screens":[338,225,178,142,142,116,116,88],"special":{1:"plex"},"notes":"T. 032-461-8800"},
    {"region":"인천","sub_region":"인천 남동구","name":"구월",      "type":"direct",
     "screens":[302,201,161,128,128,102,102,78],"special":{},"notes":"T. 032-439-8800"},
    {"region":"인천","sub_region":"인천 부평구","name":"부평",      "type":"direct",
     "screens":[356,237,189,150,150,120,120,92,92],"special":{1:"plex"},"notes":"T. 032-510-8800"},
    {"region":"인천","sub_region":"인천 부평구","name":"삼산",      "type":"direct",
     "screens":[272,181,145,116,116,92,92,70],"special":{},"notes":"T. 032-527-8800"},
    {"region":"인천","sub_region":"인천 서구","name":"청라",        "type":"direct",
     "screens":[362,241,190,152,152,122,122,92,92],"special":{1:"plex"},"notes":"T. 032-560-8800"},
    {"region":"인천","sub_region":"인천 연수구","name":"송도",      "type":"direct",
     "screens":[414,276,218,173,173,138,138,110,110,84],"special":{1:"plex",3:"mx4d"},"notes":"T. 032-710-8800"},

    # ── 강원 ──────────────────────────────────────────────────────────────────
    {"region":"강원","sub_region":"강원 강릉시","name":"강릉",      "type":"direct",
     "screens":[224,149,120,96,96,74,74],"special":{},"notes":"T. 033-648-8800"},
    {"region":"강원","sub_region":"강원 동해시","name":"동해",      "type":"contract",
     "screens":[168,112,90,72,72],"special":{},"notes":"T. 033-534-8800"},
    {"region":"강원","sub_region":"강원 삼척시","name":"삼척",      "type":"contract",
     "screens":[144,96,78,62,62],"special":{},"notes":"T. 033-574-8800"},
    {"region":"강원","sub_region":"강원 속초시","name":"속초",      "type":"contract",
     "screens":[186,124,100,80,80,62],"special":{},"notes":"T. 033-638-8800"},
    {"region":"강원","sub_region":"강원 원주시","name":"원주",      "type":"direct",
     "screens":[258,172,138,110,110,88,88,68],"special":{},"notes":"T. 033-746-8800"},
    {"region":"강원","sub_region":"강원 춘천시","name":"춘천",      "type":"direct",
     "screens":[264,176,142,113,113,90,90,68],"special":{},"notes":"T. 033-252-8800"},

    # ── 충북 ──────────────────────────────────────────────────────────────────
    {"region":"충북","sub_region":"충북 음성군","name":"음성",      "type":"contract",
     "screens":[164,109,88,70,70],"special":{},"notes":"T. 043-877-8800"},
    {"region":"충북","sub_region":"충북 청주시 상당구","name":"청주성안길","type":"direct",
     "screens":[228,152,122,97,97,76,76],"special":{},"notes":"T. 043-255-8800"},
    {"region":"충북","sub_region":"충북 청주시 흥덕구","name":"청주", "type":"direct",
     "screens":[344,229,183,146,146,116,116,90,90],"special":{1:"plex"},"notes":"T. 043-238-8800"},
    {"region":"충북","sub_region":"충북 충주시","name":"충주",      "type":"direct",
     "screens":[224,149,120,96,96,74,74],"special":{},"notes":"T. 043-848-8800"},

    # ── 충남 ──────────────────────────────────────────────────────────────────
    {"region":"충남","sub_region":"충남 공주시","name":"공주",      "type":"contract",
     "screens":[180,189,120,137,132],"special":{},"notes":"19.09.19 개관"},
    {"region":"충남","sub_region":"충남 당진시","name":"당진",      "type":"contract",
     "screens":[154,103,82,66,66],"special":{},"notes":"T. 041-357-8800"},
    {"region":"충남","sub_region":"충남 논산시","name":"논산",      "type":"contract",
     "screens":[152,101,82,65,65],"special":{},"notes":"T. 041-733-8800"},
    {"region":"충남","sub_region":"충남 서산시","name":"서산",      "type":"contract",
     "screens":[223,120,100,223,120,100,59],"special":{},"notes":"T. 041-669-8800"},
    {"region":"충남","sub_region":"충남 아산시","name":"아산",      "type":"direct",
     "screens":[302,201,161,128,128,102,102,78],"special":{},"notes":"T. 041-549-8800"},
    {"region":"충남","sub_region":"충남 천안시 서북구","name":"천안", "type":"direct",
     "screens":[358,239,190,152,152,122,122,92,92],"special":{1:"plex"},"notes":"T. 041-622-8800"},

    # ── 대전 ──────────────────────────────────────────────────────────────────
    {"region":"대전","sub_region":"대전 서구","name":"대전",        "type":"direct",
     "screens":[416,277,220,175,175,140,140,112,112,86,86],"special":{1:"plex"},"notes":"T. 042-488-8800"},

    # ── 세종 ──────────────────────────────────────────────────────────────────
    {"region":"세종","sub_region":"세종특별자치시","name":"세종",   "type":"direct",
     "screens":[358,239,190,152,152,122,122,92,92],"special":{1:"plex"},"notes":"T. 044-866-8800"},

    # ── 전북 ──────────────────────────────────────────────────────────────────
    {"region":"전북","sub_region":"전북 군산시","name":"군산",      "type":"direct",
     "screens":[248,165,133,106,106,84,84,64],"special":{},"notes":"T. 063-468-8800"},
    {"region":"전북","sub_region":"전북 김제시","name":"김제",      "type":"contract",
     "screens":[152,101,82,65,65],"special":{},"notes":"T. 063-544-8800"},
    {"region":"전북","sub_region":"전북 익산시","name":"익산",      "type":"direct",
     "screens":[280,187,150,120,120,96,96,74],"special":{},"notes":"T. 063-858-8800"},
    {"region":"전북","sub_region":"전북 전주시 완산구","name":"전주", "type":"direct",
     "screens":[392,261,208,166,166,132,132,104,104,80],"special":{1:"plex"},"notes":"T. 063-288-8800"},
    {"region":"전북","sub_region":"전북 정읍시","name":"정읍",      "type":"contract",
     "screens":[158,105,84,67,67],"special":{},"notes":"T. 063-538-8800"},

    # ── 전남 ──────────────────────────────────────────────────────────────────
    {"region":"전남","sub_region":"전남 광양시","name":"광양",      "type":"contract",
     "screens":[186,124,100,80,80,62],"special":{},"notes":"T. 061-792-8800"},
    {"region":"전남","sub_region":"전남 나주시","name":"나주",      "type":"contract",
     "screens":[152,101,82,65,65],"special":{},"notes":"T. 061-334-8800"},
    {"region":"전남","sub_region":"전남 목포시","name":"목포",      "type":"direct",
     "screens":[268,179,143,114,114,92,92,70],"special":{},"notes":"T. 061-282-8800"},
    {"region":"전남","sub_region":"전남 순천시","name":"순천",      "type":"direct",
     "screens":[296,197,158,126,126,100,100,78],"special":{},"notes":"T. 061-724-8800"},
    {"region":"전남","sub_region":"전남 여수시","name":"여수",      "type":"direct",
     "screens":[256,171,137,109,109,88,88,68],"special":{},"notes":"T. 061-662-8800"},

    # ── 광주 ──────────────────────────────────────────────────────────────────
    {"region":"광주","sub_region":"광주 서구","name":"광주",        "type":"direct",
     "screens":[364,243,193,154,154,124,124,96,96,74],"special":{1:"plex"},"notes":"T. 062-382-8800"},
    {"region":"광주","sub_region":"광주 서구","name":"상무",        "type":"direct",
     "screens":[334,223,177,142,142,114,114,88,88],"special":{1:"plex"},"notes":"T. 062-384-8800"},
    {"region":"광주","sub_region":"광주 광산구","name":"수완",      "type":"direct",
     "screens":[352,235,186,149,149,119,119,92,92],"special":{1:"plex"},"notes":"T. 062-956-8800"},

    # ── 대구 ──────────────────────────────────────────────────────────────────
    {"region":"대구","sub_region":"대구 서구","name":"대구공항",    "type":"direct",
     "screens":[184,149,149,179,183,194,262],"special":{},"notes":"T. 053-356-8800"},
    {"region":"대구","sub_region":"대구 동구","name":"대구율하",    "type":"direct",
     "screens":[172,164,168,166,286,416],"special":{},"notes":"T. 053-963-8800"},
    {"region":"대구","sub_region":"대구 중구","name":"동성로",      "type":"branch",
     "screens":[72,82,75,80,82,75,80,82,75],"special":{},"notes":"T. 053-252-8800"},
    {"region":"대구","sub_region":"대구 달서구","name":"성서",      "type":"direct",
     "screens":[166,135,247,166,135,247,166,135,247],"special":{3:"plex",6:"plex"},"notes":"T. 053-584-8800"},
    {"region":"대구","sub_region":"대구 달서구","name":"상인",      "type":"direct",
     "screens":[173,159,114,157,171,114,392],"special":{7:"plex"},"notes":"T. 070-7434-8020"},
    {"region":"대구","sub_region":"대구 달성군","name":"대구현풍",  "type":"contract",
     "screens":[172,46,149,172,46,149],"special":{},"notes":"T. 010-7408-1111"},
    {"region":"대구","sub_region":"대구 북구","name":"프리미엄칠곡","type":"direct",
     "screens":[126,58,58,126,58,58],"special":{},"notes":"T. 053-324-0600"},

    # ── 경북 ──────────────────────────────────────────────────────────────────
    {"region":"경북","sub_region":"경북 구미시","name":"프리미엄구미센트럴","type":"direct",
     "screens":[None,40,76,43,38,74,None],"special":{},"notes":"18.08.08 개관"},
    {"region":"경북","sub_region":"경북 구미시","name":"구미공단",  "type":"direct",
     "screens":[194,135,209,114,161,119,174,98],"special":{},"notes":"T. 054-462-9060"},
    {"region":"경북","sub_region":"경북 상주시","name":"상주",      "type":"contract",
     "screens":[145,98,174,54],"special":{},"notes":"20.01.17 개관"},
    {"region":"경북","sub_region":"경북 안동시","name":"프리미엄안동","type":"direct",
     "screens":[113,70,56,52,52],"special":{},"notes":"T. 054-852-0600"},
    {"region":"경북","sub_region":"경북 영주시","name":"영주",      "type":"direct",
     "screens":[191,152,60,152,165,202],"special":{},"notes":"T. 054-634-8800"},
    {"region":"경북","sub_region":"경북 영천시","name":"영천",      "type":"contract",
     "screens":[102,101,102,102,101,102],"special":{},"notes":"21.09.08 개관"},
    {"region":"경북","sub_region":"경북 포항시 남구","name":"포항", "type":"direct",
     "screens":[91,72,82,89,71,98,93,116],"special":{},"notes":"T. 054-282-8800"},
    {"region":"경북","sub_region":"경북 경주시","name":"경주",      "type":"contract",
     "screens":[136,143],"special":{},"notes":""},
    {"region":"경북","sub_region":"경북 경주시","name":"경주황성",  "type":"contract",
     "screens":[218,103,103,218,103,103],"special":{},"notes":"21.09.06 개관"},

    # ── 울산 ──────────────────────────────────────────────────────────────────
    {"region":"울산","sub_region":"울산 남구","name":"삼산",        "type":"direct",
     "screens":[356,237,189,150,150,120,120,92,92],"special":{1:"plex"},"notes":"T. 052-269-8800"},
    {"region":"울산","sub_region":"울산 북구","name":"울산",        "type":"direct",
     "screens":[306,204,163,130,130,104,104,80],"special":{},"notes":"T. 052-293-8800"},

    # ── 경남 ──────────────────────────────────────────────────────────────────
    {"region":"경남","sub_region":"경남 거제시","name":"거제",      "type":"direct",
     "screens":[240,160,128,102,102,82,82,64],"special":{},"notes":"T. 055-638-8800"},
    {"region":"경남","sub_region":"경남 김해시","name":"김해",      "type":"direct",
     "screens":[310,207,165,132,132,106,106,80],"special":{},"notes":"T. 055-328-8800"},
    {"region":"경남","sub_region":"경남 창원시 마산합포구","name":"마산","type":"direct",
     "screens":[272,181,145,116,116,92,92,70],"special":{},"notes":"T. 055-246-8800"},
    {"region":"경남","sub_region":"경남 사천시","name":"사천",      "type":"contract",
     "screens":[168,112,90,72,72],"special":{},"notes":"T. 055-832-8800"},
    {"region":"경남","sub_region":"경남 양산시","name":"양산",      "type":"direct",
     "screens":[288,192,153,122,122,98,98,76],"special":{},"notes":"T. 055-386-8800"},
    {"region":"경남","sub_region":"경남 진주시","name":"진주",      "type":"direct",
     "screens":[318,212,169,135,135,108,108,82,82],"special":{1:"plex"},"notes":"T. 055-746-8800"},
    {"region":"경남","sub_region":"경남 창원시 성산구","name":"창원","type":"direct",
     "screens":[398,265,211,168,168,135,135,104,104,80],"special":{1:"plex"},"notes":"T. 055-282-8800"},
    {"region":"경남","sub_region":"경남 통영시","name":"통영",      "type":"contract",
     "screens":[176,117,94,75,75,58],"special":{},"notes":"T. 055-645-8800"},
    {"region":"경남","sub_region":"경남 밀양시","name":"밀양",      "type":"contract",
     "screens":[156,104,83,66,66],"special":{},"notes":"T. 055-355-8800"},

    # ── 부산 ──────────────────────────────────────────────────────────────────
    {"region":"부산","sub_region":"부산 중구","name":"광복",        "type":"direct",
     "screens":[390,260,208,166,166,132,132,104,104,80],"special":{1:"plex"},"notes":"T. 051-460-8800"},
    {"region":"부산","sub_region":"부산 동래구","name":"동래",      "type":"direct",
     "screens":[312,208,166,133,133,106,106,82,82],"special":{1:"plex"},"notes":"T. 051-554-8800"},
    {"region":"부산","sub_region":"부산 기장군","name":"동부산",    "type":"direct",
     "screens":[418,278,222,177,177,142,142,112,112,86],"special":{1:"plex",3:"mx4d"},"notes":"T. 051-727-8800"},
    {"region":"부산","sub_region":"부산 금정구","name":"부산대",    "type":"direct",
     "screens":[264,176,141,113,113,90,90,68],"special":{},"notes":"T. 051-515-8800"},
    {"region":"부산","sub_region":"부산 해운대구","name":"센텀시티","type":"direct",
     "screens":[436,291,232,185,185,148,148,118,118,90],"special":{1:"plex",3:"mx4d"},"notes":"T. 051-731-8800"},
    {"region":"부산","sub_region":"부산 부산진구","name":"서면",    "type":"direct",
     "screens":[358,239,190,152,152,122,122,92,92],"special":{1:"plex"},"notes":"T. 051-806-8800"},
    {"region":"부산","sub_region":"부산 해운대구","name":"해운대",  "type":"direct",
     "screens":[318,212,169,135,135,108,108,82,82],"special":{1:"plex"},"notes":"T. 051-743-8800"},

    # ── 제주 ──────────────────────────────────────────────────────────────────
    {"region":"제주","sub_region":"제주 제주시","name":"제주",      "type":"direct",
     "screens":[294,196,157,125,125,100,100,76,76],"special":{1:"plex"},"notes":"T. 064-743-8800"},
]

# ─────────────────────────────────────────────────────────────────────────────
# 지역별 번호 매기기
# ─────────────────────────────────────────────────────────────────────────────
region_order = ["서울","경기","인천","강원","충북","충남","대전","세종",
                "전북","전남","광주","대구","경북","울산","경남","부산","제주"]

def build_rows(branches):
    region_counts = {}
    for b in branches:
        region_counts[b['region']] = region_counts.get(b['region'], 0) + 1

    rows = []
    global_idx = 1
    region_seen = {}
    for b in branches:
        r = b['region']
        region_seen[r] = region_seen.get(r, 0) + 1
        first_in_region = (region_seen[r] == 1)
        rows.append({
            "idx": global_idx,
            "first_in_region": first_in_region,
            "region_count": region_counts[r] if first_in_region else "",
            **b
        })
        global_idx += 1
    return rows

# ─────────────────────────────────────────────────────────────────────────────
# Excel 생성
# ─────────────────────────────────────────────────────────────────────────────
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "롯데시네마 상영관"

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def border(style="thin"):
    s = Side(border_style=style, color="000000")
    return Border(left=s, right=s, top=s, bottom=s)

def center(wrap=False):
    return Alignment(horizontal="center", vertical="center", wrap_text=wrap)

def left(wrap=False):
    return Alignment(horizontal="left", vertical="center", wrap_text=wrap)

thin_border = border("thin")
med_border  = border("medium")

# ─── 행 높이 / 열 너비 ────────────────────────────────────────────────────────
ws.row_dimensions[1].height = 22
ws.row_dimensions[2].height = 22
ws.row_dimensions[3].height = 28
ws.row_dimensions[4].height = 22

# 열 너비
col_widths = {
    1: 5,   # 번호
    2: 4,   # 상영관수
    3: 14,  # 지역
    4: 16,  # 극장명
}
for c in range(5, 26):    # 스크린 1~21
    col_widths[c] = 5
col_widths[26] = 5        # 스크린수
col_widths[27] = 7        # 좌석수
col_widths[28] = 28       # 비고

for c, w in col_widths.items():
    ws.column_dimensions[get_column_letter(c)].width = w

# ─── 행 1: 타이틀 ─────────────────────────────────────────────────────────────
ws.merge_cells("A1:AB1")
title_cell = ws["A1"]
title_cell.value = "롯데시네마 상영관"
title_cell.font = Font(bold=True, size=14)
title_cell.fill = fill(CLR_TITLE_BG)
title_cell.alignment = center()

# ─── 행 2: 범례 + 특수관 표시 ─────────────────────────────────────────────────
legend = [
    ("A2", "C2", "지정: 파란색  야트: 빨간색  위탁: 검정색", "000000", None),
]
ws.merge_cells("A2:D2")
ws["A2"].value = "지정: 파란색  /  야트: 빨간색  /  위탁: 검정색"
ws["A2"].font = Font(size=9)
ws["A2"].alignment = left()

# 특수관 범례 타이틀
special_headers = [
    ("E2",  "E2",  "수퍼MX4D",      CLR_ORANGE),
    ("F2",  "F2",  "샤롯대",         CLR_YELLOW),
    ("G2",  "G2",  "수퍼 LED",       CLR_YELLOW),
    ("H2",  "H2",  "아르떼",         CLR_YELLOW),
    ("I2",  "J2",  "수퍼플렉스",     CLR_TEAL),
    ("K2",  "K2",  "수퍼4D",         CLR_TEAL),
    ("L2",  "L2",  "씨네 패밀리",    CLR_BLUE),
    ("M2",  "M2",  "씨네 컴포트",    CLR_BLUE),
    ("N2",  "O2",  "광음시네마/공음LED", CLR_DARK),
]
for start, end, label, color in special_headers:
    if start != end:
        ws.merge_cells(f"{start}:{end}")
    cell = ws[start]
    cell.value = label
    cell.fill = fill(color)
    cell.font = Font(bold=True, size=8, color="FFFFFF" if color == CLR_DARK else "000000")
    cell.alignment = center()
    cell.border = thin_border

# ─── 행 3: 컬럼 헤더 ──────────────────────────────────────────────────────────
headers = ["번호","상영관수","지 역","극장명\n(디지털/3D)"] + [str(i) for i in range(1,22)] + ["스크린수","좌석수","비 고"]
header_fill = fill(CLR_HEADER_BG)
for col_idx, hdr in enumerate(headers, start=1):
    cell = ws.cell(row=3, column=col_idx, value=hdr)
    cell.fill = header_fill
    cell.font = Font(bold=True, size=9)
    cell.alignment = center(wrap=True)
    cell.border = thin_border

# ─── 데이터 행 ────────────────────────────────────────────────────────────────
rows = build_rows(branches)

SPECIAL_CELL_COLORS = {
    'mx4d':     CLR_MX4D_CELL,
    'plex':     CLR_PLEX_CELL,
    '4d':       CLR_4D_CELL,
    'charlotte':CLR_CHARLOTTE_CELL,
    'led':      CLR_LED_CELL,
    'arte':     CLR_ARTE_CELL,
    'family':   CLR_FAMILY_CELL,
    'comfort':  CLR_COMFORT_CELL,
    'gwave':    CLR_GWAVE_CELL,
}

TYPE_FONT_COLORS = {
    'direct':   "0000CC",  # 파란색 (지정)
    'branch':   "CC0000",  # 빨간색 (야트)
    'contract': "000000",  # 검정색 (위탁)
}

prev_region = None
region_start_row = {}

for data_row_idx, row in enumerate(rows):
    excel_row = data_row_idx + 4   # 헤더 3줄 + 1

    # 지역 구분 배경색
    row_bg = "FFFFFF" if data_row_idx % 2 == 0 else CLR_EVEN_ROW

    screens = row['screens']
    total_screens = len([s for s in screens if s is not None])
    total_seats   = sum(s for s in screens if s is not None)
    special_map   = row.get('special', {})

    # 열 1: 번호
    c = ws.cell(row=excel_row, column=1, value=row['idx'])
    c.fill = fill(row_bg); c.font = Font(size=9); c.alignment = center(); c.border = thin_border

    # 열 2: 상영관수 (지역 첫번째 행에만 표시)
    region_count_val = row.get('region_count', "")
    c = ws.cell(row=excel_row, column=2, value=region_count_val if region_count_val != "" else "")
    c.fill = fill(row_bg); c.font = Font(size=9,bold=True); c.alignment = center(); c.border = thin_border

    # 열 3: 지역
    c = ws.cell(row=excel_row, column=3, value=row['sub_region'])
    c.fill = fill(row_bg); c.font = Font(size=9); c.alignment = center(wrap=True); c.border = thin_border

    # 열 4: 극장명
    font_color = TYPE_FONT_COLORS.get(row['type'], "000000")
    c = ws.cell(row=excel_row, column=4, value=row['name'])
    c.fill = fill(row_bg)
    c.font = Font(size=9, color=font_color, bold=(row['type']=='direct'))
    c.alignment = center(); c.border = thin_border

    # 열 5~25: 스크린 1~21
    for screen_col_idx in range(1, 22):
        col_num = screen_col_idx + 4   # column 5~25
        seat_val = screens[screen_col_idx - 1] if screen_col_idx - 1 < len(screens) else None
        special_type = special_map.get(screen_col_idx)

        if special_type:
            bg = SPECIAL_CELL_COLORS.get(special_type, row_bg)
        else:
            bg = row_bg

        c = ws.cell(row=excel_row, column=col_num, value=seat_val)
        c.fill = fill(bg)
        c.font = Font(size=9)
        c.alignment = center()
        c.border = thin_border

    # 열 26: 스크린수
    c = ws.cell(row=excel_row, column=26, value=total_screens)
    c.fill = fill(row_bg); c.font = Font(size=9,bold=True); c.alignment = center(); c.border = thin_border

    # 열 27: 좌석수
    c = ws.cell(row=excel_row, column=27, value=total_seats)
    c.fill = fill(row_bg); c.font = Font(size=9,bold=True); c.alignment = center(); c.border = thin_border

    # 열 28: 비고
    c = ws.cell(row=excel_row, column=28, value=row.get('notes',''))
    c.fill = fill(row_bg); c.font = Font(size=9); c.alignment = left(wrap=True); c.border = thin_border

    # 지역 구분선 (지역 바뀔 때 위쪽 테두리 두껍게)
    if row['first_in_region']:
        for col in range(1, 29):
            cell = ws.cell(row=excel_row, column=col)
            existing = cell.border
            cell.border = Border(
                left=existing.left, right=existing.right,
                top=Side(border_style="medium", color="000000"),
                bottom=existing.bottom
            )

# ─── 지역 병합 (지역명 세로 병합) ─────────────────────────────────────────────
region_rows = {}
for data_row_idx, row in enumerate(rows):
    excel_row = data_row_idx + 4
    r = row['region']
    if r not in region_rows:
        region_rows[r] = []
    region_rows[r].append(excel_row)

# 지역명 표시를 위해 별도 컬럼 추가하지 않고 기존 열 3(지역)에 병합
# → 별도의 "광역 지역" 열이 필요: 현재 구조에서는 col 3 = 세부 지역, 앞에 지역 열 삽입 불필요
# 대신, 각 지역 첫 행 col 2에 지역명 기재하고 세로 병합

# 지역명을 col 2에 추가 (지역 첫 행에 지역명, 병합)
# 실제로 col 2 = 상영관수이므로, 지역 이름 별도 표시를 위해 col 3 앞에 삽입하는 대신
# 헤더 re-work: 기존 col 2 = 상영관수, 헤더를 수정해서 col 1 = 지역(광역), col 2 = 번호, ...
# → 이미 생성된 시트를 수정하기보다 접근 변경이 복잡함
# 따라서 현재 구조 유지, 지역은 각 행 sub_region에 이미 포함됨 (예: "경북 구미시")

# ─── 창 고정 (헤더 고정) ──────────────────────────────────────────────────────
ws.freeze_panes = "E4"

# ─── 저장 ────────────────────────────────────────────────────────────────────
output_path = "/home/user/Rklb-/롯데시네마_상영관_현황.xlsx"
wb.save(output_path)
print(f"✓ 저장 완료: {output_path}")
print(f"  총 {len(branches)}개 지점")
