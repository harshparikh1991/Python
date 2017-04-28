# -*- coding: utf-8 -*-
from apiclient.discovery import build
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

workbook = xlsxwriter.Workbook('video_result.xlsx')
worksheet = workbook.add_worksheet()

url_format = workbook.add_format({
    'font_color': 'blue',
    'underline':  1
})

def get_comment_threads(youtube, video_id, comments):
        threads = []
        results = youtube.commentThreads().list(
            part="snippet",
            videoId=video_id,
            textFormat="plainText",
        ).execute()
        for item in results["items"]:
          threads.append(item)
          comment = item["snippet"]["topLevelComment"]
          text = comment["snippet"]["textDisplay"]
          comments.append(text)
          
        while ("nextPageToken" in results):
            results = youtube.commentThreads().list(
                part="snippet",
                videoId=video_id,
                pageToken=results["nextPageToken"],
                 textFormat="plainText",
             ).execute()
            for item in results["items"]:
                threads.append(item)
                comment = item["snippet"]["topLevelComment"]
                text = comment["snippet"]["textDisplay"]
                comments.append(text)
        print ("Total threads: " + str(len(threads)))

        return threads


def get_comments(youtube, parent_id, comments):
    results = youtube.comments().list(
        part="snippet",
        parentId=parent_id,
        textFormat="plainText"
    ).execute()
    for item in results["items"]:
        text = item["snippet"]["textDisplay"]
        comments.append(text)
    
    return results["items"]

def youtube_search(options):
    youtube = build(YOUTUBE_API_SERVICE_NAME, YOUTUBE_API_VERSION, developerKey=DEVELOPER_KEY)
    # Call the search.list method to retrieve results matching the specified
    # query term.
    search_response = youtube.search().list(q=options.q, part="id,snippet", maxResults=options.max_results).execute()
    
    videos = []
    channels = []
    playlists = []
    comments = []
    row = 1
 
    worksheet.write(0, 0, "Sr No")
    worksheet.write(0, 1, "title")
    worksheet.write(0, 2, "videoId")
    worksheet.write(0, 3, "viewCount")
    worksheet.write(0, 4, "likeCount")
    worksheet.write(0, 5, "dislikeCount")
    worksheet.write(0, 6, "commentCount")
    worksheet.write(0, 7, "favoriteCount")
    
    # create a CSV output for video list    
    # Add each result to the appropriate list, and then display the lists of
    # matching videos, channels, and playlists.
    for search_result in search_response.get("items", []):
        if search_result["id"]["kind"] == "youtube#video":
            #videos.append("%s (%s)" % (search_result["snippet"]["title"],search_result["id"]["videoId"]))
            title = search_result["snippet"]["title"]
            title = unidecode.unidecode(title)  # Dongho 08/10/16
            videoId = search_result["id"]["videoId"]
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
                    worksheet1 = workbook.add_worksheet()
                    worksheet1.name = str(row)
                    worksheet1.write(0,0,"Posted Date")
                    worksheet1.write(0,1,"User")
                    worksheet1.write(0,2,"Comment")
                    if row == 1:
                        #url = urlopen('https://www.googleapis.com/youtube/v3/commentThreads?key=AIzaSyA1L8VxJdkL_Br6mIznY5jx6CZR3EBJtNQ&textFormat=plainText&part=snippet&videoId=' + videoId + '&maxResults=100').read()
                        #result = json.loads(url)  # result is now a dict
                        #print ("comment: " + result['items']['snippet']['topLevelComment']['snippet']['authorDisplayName'])
                        video_comment_threads = get_comment_threads(youtube, videoId, comments)
                        count=1
                        for thread in video_comment_threads:
                            if count == 20:
                                break
                            get_comments(youtube, thread["id"], comments)
                            count = count + 1
                        print("Completed Phase 2")
                        row1=1
                        for comment in comments:
                            worksheet1.write(row1, 2, str(comment.encode("utf-8")))
                            row1 = row1 + 1
                        print("Completed Phase 3")
                        print ("Total comments:" + str(len(comments)))
                if 'favoriteCount' not in video_result["statistics"]:
                    favoriteCount = 0
                else:
                    favoriteCount = video_result["statistics"]["favoriteCount"]
                    
                    
            worksheet.write(row, 0, row)
            worksheet.write_url(row, 1, 'internal:'+ worksheet1.name +'!A1')
            worksheet.write(row, 1, title)
            worksheet.write(row, 2, videoId)
            worksheet.write(row, 3, viewCount)
            worksheet.write(row, 4, likeCount)
            worksheet.write(row, 5, dislikeCount)
            worksheet.write(row, 6, commentCount)
            worksheet.write(row, 7, favoriteCount)
            row = row + 1
    workbook.close()
  
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Search on YouTube')
    parser.add_argument("--q", help="Search term", default="Google")
    parser.add_argument("--max-results", help="Max results", default=50)
    args = parser.parse_args()
    #try:
    youtube_search(args)
    #except HttpError, e:
    #    print ("An HTTP error %d occurred:\n%s" % (e.resp.status, e.content))