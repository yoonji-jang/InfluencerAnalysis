from googleapiclient.discovery import build
import json
import requests
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image
from enum import Enum
import io
from tqdm import trange


# version info
VERSION = 2

# Todo: add error handling, excel cell size
RETURN_ERR = -1


# mode
class Index(Enum):
    URL = 0
    TITLE = 1
    VIEW = 2
    LIKE = 3
    COMMENTS = 4
    THUMBNAIL = 5


#input
print("[Info] InfluencerAnalysis V" + str(VERSION))
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
    image_scale = 10
    response = requests.get(img_url)
    img_file = io.BytesIO(response.content)
    thumbnailImage = Image(img_file)
    thumbnailImage.width /= image_scale
    thumbnailImage.height /= image_scale
    colChar = get_column_letter(col)
    thumbnailImage.anchor = "%s"%colChar + "%s"%row
    sheet.add_image(thumbnailImage)
    #sheet.column_dimensions[colChar].width = thumbnailImage.width
    sheet.row_dimensions[row].height = thumbnailImage.height


def RequestVideoInfo(vID):
    VIDEO_SEARCH_URL = "https://www.googleapis.com/youtube/v3/videos?id=" + vID + "&key=" + DEVELOPER_KEY + "&part=snippet,statistics&fields=items(id,snippet(channelId, title, thumbnails.high),statistics)"
    response = requests.get(VIDEO_SEARCH_URL).json()
    return response


def RequestChannelInfo(cID):
    CHANNEL_SEARCH_URL = "https://www.googleapis.com/youtube/v3/channels?id=" + cID + "&key=" + DEVELOPER_KEY + "&part=snippet,statistics&fields=items(id,snippet(title, thumbnails.high),statistics)"
    response = requests.get(CHANNEL_SEARCH_URL).json()
    return response


def RequestChannelContentsInfo(youtube, cID):
    try:
        response = youtube.search().list(
            channelId = cID,
            type = "video",
            order = "date",
            part = "id",
            fields = "items(id)",
            maxResults = MAX_RESULT
        ).execute()
    except Exception as exception:
        print("[Warning] " + str(exception))
        return RETURN_ERR
    return response


def UpdateVideoInfoToExcel(sheet, r, start_c, data):
    sheet.cell(row=r, column=start_c + vIndex.URL.value).value = data[vIndex.URL]
    sheet.cell(row=r, column=start_c + vIndex.TITLE.value).value = data[vIndex.TITLE]
    sheet.cell(row=r, column=start_c + vIndex.VIEW.value).value = round(float(data[vIndex.VIEW]), 2)
    sheet.cell(row=r, column=start_c + vIndex.LIKE.value).value = round(float(data[vIndex.LIKE]), 2)
    sheet.cell(row=r, column=start_c + vIndex.COMMENTS.value).value = round(float(data[vIndex.COMMENTS]), 2)
    InsertImage(sheet, data[vIndex.THUMBNAIL], r, start_c + vIndex.THUMBNAIL.value)


def UpdateChannelInfoToExcel(sheet, r, start_c, data):
    sheet.cell(row=r, column=start_c + cIndex.PROFILE_URL.value).value = data[cIndex.PROFILE_URL]
    InsertImage(sheet, data[cIndex.PROFILE_IMG], r, start_c + cIndex.PROFILE_IMG.value)
    sheet.cell(row=r, column=start_c + cIndex.SUBSCRIBER.value).value = data[cIndex.SUBSCRIBER]

    sheet.cell(row=r, column=start_c + cIndex.POST_VIEW.value).value = round(float(data[cIndex.POST_VIEW]), 2)
    sheet.cell(row=r, column=start_c + cIndex.POST_LIKE.value).value = round(float(data[cIndex.POST_LIKE]), 2)
    sheet.cell(row=r, column=start_c + cIndex.POST_COMMENT.value).value = round(float(data[cIndex.POST_COMMENT]),2 )
    sheet.cell(row=r, column=start_c + cIndex.POST_ENGAGE.value).value = round(float(data[cIndex.POST_ENGAGE]), 2)


def GetChannelData(cID, channel_info, channel_contents_info):
    arr = json.dumps(channel_info)
    jsonObject = json.loads(arr)
    if ((jsonObject.get('error')) or ('items' not in jsonObject)):
        print("[Warning] response error! : " + cID)
        print(jsonObject['error'])
        return RETURN_ERR
    items = jsonObject['items']
    if len(items) <= 0:
        print("[Error] no items for Channel Data: " + cID)
        return RETURN_ERR
    item = jsonObject['items'][0]
    ret = {}
    ret[cIndex.PROFILE_URL] = ""
    ret[cIndex.PROFILE_IMG] = ""
    ret[cIndex.SUBSCRIBER] = 0
    ret[cIndex.POST_VIEW] = 0
    ret[cIndex.POST_LIKE] = 0
    ret[cIndex.POST_COMMENT] = 0
    ret[cIndex.POST_ENGAGE] = 0
    
    try:
        ret[cIndex.PROFILE_URL] = "https://www.youtube.com/channel/" + item['id']
        ret[cIndex.PROFILE_IMG] = item['snippet']['thumbnails']['high']['url']
        ret[cIndex.SUBSCRIBER] = item['statistics']['subscriberCount']

        nViewCnt = 0
        nLikeCnt = 0
        nCommentCnt = 0
        nView = 0;
        nLike = 0;
        nComment = 0;
        for content in channel_contents_info.get("items", []):
            if content["id"]["kind"] != "youtube#video":
                print("[Warning] Type is not video!! check the input: " + cID)
                return RETURN_ERR

            vID = content["id"]["videoId"]
            res_json = RequestVideoInfo(vID)

            video_info = GetVideoData(vID, res_json)
            view = int(video_info[vIndex.VIEW])
            like = int(video_info[vIndex.LIKE])
            comments = int(video_info[vIndex.COMMENTS])

            if view > 0:
                nViewCnt += view
                nView += 1
            if like > 0:
                nLikeCnt += like
                nLike += 1
            if comments > 0:
                nCommentCnt += comments
                nComment += 1
    except Exception as exception:
        print("[Warning]: " + str(exception) + ", Channel ID: " + cID)
        pass
        
    if nView > 0:
        ret[cIndex.POST_VIEW] = nViewCnt / nView
    if nLike > 0:
        ret[cIndex.POST_LIKE] = nLikeCnt / nLike
    if nComment > 0:
        ret[cIndex.POST_COMMENT] = nCommentCnt / nComment
    if nViewCnt > 0:
        ret[cIndex.POST_ENGAGE] = ((nLikeCnt + nCommentCnt) / nViewCnt) * 100
    return ret


def GetVideoData(vID, input_json):
    arr = json.dumps(input_json)
    jsonObject = json.loads(arr)
    if ((jsonObject.get('error')) or ('items' not in jsonObject)):
        print("[Warning] response error! : " + vID)
        print(jsonObject['error'])
        return RETURN_ERR
    items = jsonObject['items']
    if len(items) <= 0:
        print("[Warning] no items for Video Data: " + vID)
        return RETURN_ERR
    item = jsonObject['items'][0]
    ret = {}
    ret[vIndex.URL] = ""
    ret[vIndex.TITLE] = ""
    ret[vIndex.VIEW] = 0
    ret[vIndex.LIKE] = 0
    ret[vIndex.COMMENTS] = 0 
    ret[vIndex.THUMBNAIL] = ""    
    
    if ('id' in item):
        ret[vIndex.URL] = "https://www.youtube.com/watch?v=" + item['id']
    if ('snippet' in item) and ('title' in item['snippet']):
        ret[vIndex.TITLE] = item['snippet']['title']
    if ('statistics' in item):
        statistics = item['statistics']
        if ('viewCount' in statistics):
            ret[vIndex.VIEW] = statistics['viewCount']
        if ('likeCount' in statistics):
            ret[vIndex.LIKE] = statistics['likeCount']
        if ('commentCount' in statistics):
            ret[vIndex.COMMENTS] = statistics['commentCount']
    ret[vIndex.THUMBNAIL] = "https://img.youtube.com/vi/" + vID + "/maxresdefault.jpg"

    return ret


def run_VideoAnalysis(sheet):
    print("[Info] Running VideoAnalysis")
    max_row = sheet.max_row + 1
    for row in trange(START_ROW, max_row):
        vID = sheet.cell(row, START_COL).value
        if vID == None:
            continue
        res_json = RequestVideoInfo(vID)
        df_just_video = GetVideoData(vID, res_json)
        if df_just_video == RETURN_ERR:
            continue
        UpdateVideoInfoToExcel(sheet, row, START_COL + 1, df_just_video)


def run_InfluencerAnalysis(sheet):
    print("[Info] Running InfluencerAnalysis")
    youtube = build(YOUTUBE_API_SERVICE_NAME, YOUTUBE_API_VERSION, developerKey=DEVELOPER_KEY)
    max_row = sheet.max_row + 1
    for row in trange(START_ROW, max_row):
        cID = sheet.cell(row, START_COL).value
        if cID == None:
            continue
        channel_info = RequestChannelInfo(cID)
        channel_contents_info = RequestChannelContentsInfo(youtube, cID)
        if channel_contents_info == RETURN_ERR:
            continue
        df_just_channel = GetChannelData(cID, channel_info, channel_contents_info)
        if df_just_channel == RETURN_ERR:
            continue
        UpdateChannelInfoToExcel(sheet, row, START_COL + 1, df_just_channel)




# read excel
xlsx = openpyxl.load_workbook(INPUT_EXCEL)
cSheet = xlsx.worksheets[INFLUENCER_SHEET]
vSheet = xlsx.worksheets[VIDEO_SHEET]
print("[Info] Open input excel: " + INPUT_EXCEL)

# run Analysis
run_VideoAnalysis(vSheet)
run_InfluencerAnalysis(cSheet)

# save excel
xlsx.save(OUTPUT_EXCEL)
print("[Info] Done saving excel: " + OUTPUT_EXCEL)
