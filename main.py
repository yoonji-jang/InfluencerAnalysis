from googleapiclient.discovery import build
import json
import requests
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image
from enum import Enum
import io

# mode
class Index(Enum):
    URL = 0
    TITLE = 1
    VIEW = 2
    LIKE = 3
    COMMENTS = 4
    THUMBNAIL = 5


# input
#VIDEO_ID = "JpTqSzm4JOk"
INFLUENCER_SHEET = 0
VIDEO_SHEET = 1
START_ROW = 5
START_COL = 2

# youtube setting
YOUTUBE_API_SERVICE_NAME = "youtube"
YOUTUBE_API_VERSION = "v3"
# FREEBASE_SEARCH_URL = "https://www.googleapis.com/freebase/v1/search?%s"


# read excel
xlsx = openpyxl.load_workbook('./InputSample_2.xlsx')
sheet = xlsx.worksheets[VIDEO_SHEET]
max_row = sheet.max_row + 1
print("open excel")


def RequestInfo(vID):
    VIDEO_SEARCH_URL = "https://www.googleapis.com/youtube/v3/videos?id=" + vID + "&key=" + DEVELOPER_KEY + "&part=snippet,statistics&fields=items(id,snippet(channelId,title, thumbnails.high.url),statistics)"
    response = requests.get(VIDEO_SEARCH_URL).json()
    return response

class Index(Enum):
    URL = 0
    TITLE = 1
    VIEW = 2
    LIKE = 3
    COMMENTS = 4
    THUMBNAIL = 5


def UpdateToExcel(r, start_c, data):
    sheet.cell(row=r, column=start_c + Index.URL.value).value = data[Index.URL]
    sheet.cell(row=r, column=start_c + Index.TITLE.value).value = data[Index.TITLE]
    sheet.cell(row=r, column=start_c + Index.VIEW.value).value = data[Index.VIEW]
    sheet.cell(row=r, column=start_c + Index.LIKE.value).value = data[Index.LIKE]
    sheet.cell(row=r, column=start_c + Index.COMMENTS.value).value = data[Index.COMMENTS]
    response = requests.get(data[Index.THUMBNAIL])
    img_file = io.BytesIO(response.content)
    thumbnailImage = Image(img_file)
    thumbnailImage.width = 96
    thumbnailImage.height = 72
    colChar = get_column_letter(start_c + Index.THUMBNAIL.value)
    thumbnailImage.anchor = "%s"%colChar + "%s"%r
    sheet.add_image(thumbnailImage)
    sheet.column_dimensions[colChar].width = thumbnailImage.width
    sheet.row_dimensions[r].height = thumbnailImage.height


def GetVideoData(input_json):
    arr = json.dumps(input_json)
    jsonObject = json.loads(arr)
    item = jsonObject['items'][0]
    ret = {}
    ret[Index.URL] = "https://www.youtube.com/watch?v=" + item['id']
    ret[Index.TITLE] = item['snippet']['title']
    ret[Index.VIEW] = item['statistics']['viewCount']
    ret[Index.LIKE] = item['statistics']['likeCount']
    ret[Index.COMMENTS] = item['statistics']['commentCount']
    ret[Index.THUMBNAIL] = item['snippet']['thumbnails']['high']['url']
    return ret


def run_VideoAnalysis():
    print("progressing...")
    for row in range(START_ROW, max_row):
        vID = sheet.cell(row, START_COL).value
        if vID == None:
            continue
        res_json = RequestInfo(vID)
        df_just_video = GetVideoData(res_json)
        UpdateToExcel(row, START_COL + 1, df_just_video)




run_VideoAnalysis()

xlsx.save('output.xlsx')
print("done saving excel: output.xlsx")





