from googleapiclient.discovery import build
import json
import requests
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image
from enum import Enum
import io

# Todo: add error handling, excel cell size

# mode
class Index(Enum):
    URL = 0
    TITLE = 1
    VIEW = 2
    LIKE = 3
    COMMENTS = 4
    THUMBNAIL = 5


#input
input_file = open(".\input.txt", "r", encoding="UTF8")
input_data=input_file.readlines()
input_file.close()
dict = {}
for line in input_data:
    key_value = line.strip().split('=')
    if len(key_value)==2:
        dict[key_value[0]] = key_value[1]

DEVELOPER_KEY = dict["DEVELOPER_KEY"]
INPUT_EXCEL = dict["INPUT_EXCEL"]
OUTPUT_EXCEL = dict["OUTPUT_EXCEL"]
INFLUENCER_SHEET = int(dict["INFLUENCER_SHEET"])
VIDEO_SHEET = int(dict["VIDEO_SHEET"])
START_ROW = int(dict["START_ROW"])
START_COL = int(dict["START_COL"])
MAX_RESULT = int(dict["MAX_RESULT"])


# youtube api
YOUTUBE_API_SERVICE_NAME="youtube"
YOUTUBE_API_VERSION="v3"

class vIndex(Enum):
    URL = 0
    TITLE = 1
    VIEW = 2
    LIKE = 3
    COMMENTS = 4
    THUMBNAIL = 5

class cIndex(Enum):
    PROFILE_URL = 0
    PROFILE_IMG = 1
    SUBSCRIBER = 2
    POST_VIEW = 3
    POST_LIKE = 4
    POST_COMMENT = 5
    POST_ENGAGE = 6
    AGE = 7
    GENDER = 8
    LOCATION = 9
    LANGUAGE = 10


def InsertImage(sheet, img_url, row, col):
    response = requests.get(img_url)
    img_file = io.BytesIO(response.content)
    thumbnailImage = Image(img_file)
    thumbnailImage.width = 96
    thumbnailImage.height = 72
    colChar = get_column_letter(col)
    thumbnailImage.anchor = "%s"%colChar + "%s"%row
    sheet.add_image(thumbnailImage)
    sheet.column_dimensions[colChar].width = 30
    sheet.row_dimensions[row].height = 72


def RequestVideoInfo(vID):
    VIDEO_SEARCH_URL = "https://www.googleapis.com/youtube/v3/videos?id=" + vID + "&key=" + DEVELOPER_KEY + "&part=snippet,statistics&fields=items(id,snippet(channelId, title, thumbnails.high),statistics)"
    response = requests.get(VIDEO_SEARCH_URL).json()
    return response


def RequestChannelInfo(cID):
    CHANNEL_SEARCH_URL = "https://www.googleapis.com/youtube/v3/channels?id=" + cID + "&key=" + DEVELOPER_KEY + "&part=snippet,statistics&fields=items(id,snippet(title, thumbnails.high),statistics)"
    response = requests.get(CHANNEL_SEARCH_URL).json()
    return response


def RequestChannelContentsInfo(youtube, cID):
    response = youtube.search().list(
        channelId = cID,
        type = "video",
        order = "date",
        part = "id",
        fields = "items(id)",
        maxResults = MAX_RESULT
    ).execute()
    return response


def UpdateVideoInfoToExcel(sheet, r, start_c, data):
    sheet.cell(row=r, column=start_c + vIndex.URL.value).value = data[vIndex.URL]
    sheet.cell(row=r, column=start_c + vIndex.TITLE.value).value = data[vIndex.TITLE]
    sheet.cell(row=r, column=start_c + vIndex.VIEW.value).value = data[vIndex.VIEW]
    sheet.cell(row=r, column=start_c + vIndex.LIKE.value).value = data[vIndex.LIKE]
    sheet.cell(row=r, column=start_c + vIndex.COMMENTS.value).value = data[vIndex.COMMENTS]
    InsertImage(sheet, data[vIndex.THUMBNAIL], r, start_c + vIndex.THUMBNAIL.value)


def UpdateChannelInfoToExcel(sheet, r, start_c, data):
    sheet.cell(row=r, column=start_c + cIndex.PROFILE_URL.value).value = data[cIndex.PROFILE_URL]
    InsertImage(sheet, data[cIndex.PROFILE_IMG], r, start_c + cIndex.PROFILE_IMG.value)
    sheet.cell(row=r, column=start_c + cIndex.SUBSCRIBER.value).value = data[cIndex.SUBSCRIBER]

    sheet.cell(row=r, column=start_c + cIndex.POST_VIEW.value).value = data[cIndex.POST_VIEW]
    sheet.cell(row=r, column=start_c + cIndex.POST_LIKE.value).value = data[cIndex.POST_LIKE]
    sheet.cell(row=r, column=start_c + cIndex.POST_COMMENT.value).value = data[cIndex.POST_COMMENT]
    sheet.cell(row=r, column=start_c + cIndex.POST_ENGAGE.value).value = data[cIndex.POST_ENGAGE]


def GetChannelData(channel_info, channel_contents_info):
    arr = json.dumps(channel_info)
    jsonObject = json.loads(arr)
    item = jsonObject['items'][0]
    ret = {}
    ret[cIndex.PROFILE_URL] = "https://www.youtube.com/channel/" + item['id']
    ret[cIndex.PROFILE_IMG] = item['snippet']['thumbnails']['high']['url']
    ret[cIndex.SUBSCRIBER] = item['statistics']['subscriberCount']

    nViewCnt = 0
    nLikeCnt = 0
    nCommentCnt = 0
    for content in channel_contents_info.get("items", []):
        if content["id"]["kind"] != "youtube#video":
            print("type is not video!! check the input")
            # return -1

        vID = content["id"]["videoId"]
        res_json = RequestVideoInfo(vID)

        video_info = GetVideoData(res_json)
        nViewCnt += int(video_info[vIndex.VIEW])
        nLikeCnt += int(video_info[vIndex.LIKE])
        nCommentCnt += int(video_info[vIndex.COMMENTS])

    ret[cIndex.POST_VIEW] = nViewCnt
    ret[cIndex.POST_LIKE] = nLikeCnt
    ret[cIndex.POST_COMMENT] = nCommentCnt
    if nViewCnt != 0:
        ret[cIndex.POST_ENGAGE] = ((nLikeCnt + nCommentCnt) / nViewCnt) * 100
    return ret


def GetVideoData(input_json):
    arr = json.dumps(input_json)
    jsonObject = json.loads(arr)
    item = jsonObject['items'][0]
    ret = {}
    ret[vIndex.URL] = "https://www.youtube.com/watch?v=" + item['id']
    ret[vIndex.TITLE] = item['snippet']['title']
    ret[vIndex.VIEW] = item['statistics']['viewCount']
    ret[vIndex.LIKE] = item['statistics']['likeCount']
    ret[vIndex.COMMENTS] = item['statistics']['commentCount']
    ret[vIndex.THUMBNAIL] = item['snippet']['thumbnails']['high']['url']
    return ret


def run_VideoAnalysis(sheet):
    max_row = sheet.max_row + 1
    for row in range(START_ROW, max_row):
        vID = sheet.cell(row, START_COL).value
        if vID == None:
            continue
        res_json = RequestVideoInfo(vID)
        df_just_video = GetVideoData(res_json)
        UpdateVideoInfoToExcel(sheet, row, START_COL + 1, df_just_video)


def run_InfluencerAnalysis(sheet):
    youtube = build(YOUTUBE_API_SERVICE_NAME, YOUTUBE_API_VERSION, developerKey=DEVELOPER_KEY)
    max_row = sheet.max_row + 1
    for row in range(START_ROW, max_row):
        cID = sheet.cell(row, START_COL).value
        if cID == None:
            continue
        channel_info = RequestChannelInfo(cID)
        channel_contents_info = RequestChannelContentsInfo(youtube, cID)
        df_just_channel = GetChannelData(channel_info, channel_contents_info)
        UpdateChannelInfoToExcel(sheet, row, START_COL + 1, df_just_channel)




# read excel
xlsx = openpyxl.load_workbook(INPUT_EXCEL)
cSheet = xlsx.worksheets[INFLUENCER_SHEET]
vSheet = xlsx.worksheets[VIDEO_SHEET]
print("open excel")

# run Analysis
run_VideoAnalysis(vSheet)
run_InfluencerAnalysis(cSheet)

# save excel
xlsx.save(OUTPUT_EXCEL)
print("done saving excel: " + OUTPUT_EXCEL)
