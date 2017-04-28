# -*- coding: utf-8 -*-
from apiclient.discovery import build
from datetime import datetime, timezone,timedelta
years = -50
days_per_year = 365.24

#from apiclient.errors import HttpError
#from oauth2client.tools import argparser # removed by Dongho
import argparse
import xlsxwriter
import unidecode
import re
regex = re.compile(r"^.*interfaceOpDataFile.*$", re.IGNORECASE)
# Set DEVELOPER_KEY to the API key value from the APIs & auth > Registered apps
# tab of
#   https://cloud.google.com/console
# Please ensure that you have enabled the YouTube Data API for your project.
DEVELOPER_KEY = "AIzaSyA1L8VxJdkL_Br6mIznY5jx6CZR3EBJtNQ"
YOUTUBE_API_SERVICE_NAME = "youtube"
YOUTUBE_API_VERSION = "v3"

def youtube_search(options):
    youtube = build(YOUTUBE_API_SERVICE_NAME, YOUTUBE_API_VERSION, developerKey=DEVELOPER_KEY)
    
    dtMaxEnd = datetime.utcnow() # <-- get time in UTC
    dtMaxStart = dtMaxEnd + timedelta(days=(years*days_per_year))
    
    #print("Start Date: " + dtStart)
    #print("End Date: " + dtEnd)
    
    row = 1
    workbook = xlsxwriter.Workbook('video_list_' + options.q + '.xlsx')
    worksheet = workbook.add_worksheet()

    worksheet.write(0, 0, "Sr No")
    worksheet.write(0, 1, "title")
    worksheet.write(0, 2, "videoId")
    worksheet.write(0, 3, "viewCount")
    worksheet.write(0, 4, "likeCount")
    worksheet.write(0, 5, "dislikeCount")
    worksheet.write(0, 6, "commentCount")

    url_format = workbook.add_format({
        'font_color': 'blue',
        'underline':  1
    })
    arrVideoID = []
    # Call the search.list method to retrieve results matching the specified
    # query term.
    
    #while (dtMaxEnd >= dtMaxStart):
    while (row <= 5000):
        dtOneLess = dtMaxEnd + timedelta(hours=0)
        dtMaxEnd = dtMaxEnd + timedelta(hours=-12)
        dtEndZ = dtOneLess.isoformat("T") + "Z"
        dtStartZ = dtMaxEnd.isoformat("T") + "Z"
        search_response = youtube.search().list(q=options.q, part="id,snippet", maxResults=options.max_results, publishedAfter=dtStartZ, publishedBefore=dtEndZ).execute()    
        
        videos = []
        channels = []
        playlists = []
        
        for search_result in search_response.get("items", []):
            if (row >5000):
                return
            if search_result["id"]["kind"] == "youtube#video":
                title = search_result["snippet"]["title"]
                title = unidecode.unidecode(title)  # Dongho 08/10/16
                videoId = search_result["id"]["videoId"]
                if not videoId in arrVideoID:
                    arrVideoID.append(videoId)
                    #publishedAt = search_result["snippet"]["publishedAt"]
                    
                    video_response = youtube.videos().list(id=videoId,part="statistics").execute()
                    for video_result in video_response.get("items",[]):
                        viewCount = video_result["statistics"]["viewCount"]
                        if 'likeCount' not in video_result["statistics"]:
                            likeCount = 0
                        else:
                            likeCount = video_result["statistics"]["likeCount"]
                        if 'dislikeCount' not in video_result["statistics"]:
                            dislikeCount = 0
                        else:
                            dislikeCount = video_result["statistics"]["dislikeCount"]
                        if 'commentCount' not in video_result["statistics"]:
                            commentCount = 0
                        else:
                            commentCount = video_result["statistics"]["commentCount"]                    
                    
                    if(int(commentCount) != 0 and str(commentCount) != "0"):
                        worksheet.write(row, 0, row)
                        worksheet.write(row, 1, title)
                        worksheet.write(row, 2, videoId)
                        worksheet.write(row, 3, viewCount)
                        worksheet.write(row, 4, likeCount)
                        worksheet.write(row, 5, dislikeCount)
                        worksheet.write(row, 6, commentCount)
                        #worksheet.write(row, 3, publishedAt)
                        if(row > 5000):
                            return
                        
                        row = row + 1
        
        nextPageToken = search_response.get('nextPageToken')
        
        while ('nextPageToken' in search_response):
            if(row <= 5000):
                nextPage = youtube.search().list(q=options.q, part="id,snippet", maxResults=options.max_results, pageToken=nextPageToken, publishedAfter=dtStartZ, publishedBefore=dtEndZ).execute()
                for search_result in nextPage.get("items", []):
                    if search_result["id"]["kind"] == "youtube#video":
                        title = search_result["snippet"]["title"]
                        title = unidecode.unidecode(title)  # Dongho 08/10/16
                        videoId = search_result["id"]["videoId"]
                        if not videoId in arrVideoID:
                            arrVideoID.append(videoId)
                            video_response = youtube.videos().list(id=videoId,part="statistics").execute()
                            for video_result in video_response.get("items",[]):
                                viewCount = video_result["statistics"]["viewCount"]
                                if 'likeCount' not in video_result["statistics"]:
                                    likeCount = 0
                                else:
                                    likeCount = video_result["statistics"]["likeCount"]
                                if 'dislikeCount' not in video_result["statistics"]:
                                    dislikeCount = 0
                                else:
                                    dislikeCount = video_result["statistics"]["dislikeCount"]
                                if 'commentCount' not in video_result["statistics"]:
                                    commentCount = 0
                                else:
                                    commentCount = video_result["statistics"]["commentCount"]
                                    
                            if(int(commentCount) != 0 and str(commentCount) != "0"):
                                worksheet.write(row, 0, row)
                                worksheet.write(row, 1, title)
                                worksheet.write(row, 2, videoId)
                                worksheet.write(row, 3, viewCount)
                                worksheet.write(row, 4, likeCount)
                                worksheet.write(row, 5, dislikeCount)
                                worksheet.write(row, 6, commentCount)
                                
                                #worksheet.write(row, 3, publishedAt)
                                if(row > 5000):
                                    return
                                
                                row = row + 1
                            
                if 'nextPageToken' not in nextPage:
                    search_response.pop('nextPageToken', None)
                else:
                    nextPageToken = nextPage.get('nextPageToken')
            

    workbook.close()
    print("executed")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Search on YouTube')
    parser.add_argument("--q", help="Search term", default="mcdonalds")
    parser.add_argument("--max-results", help="Max results", default=50)
    args = parser.parse_args()
    #try:
    youtube_search(args)
    #except HttpError, e:
    #    print ("An HTTP error %d occurred:\n%s" % (e.resp.status, e.content))