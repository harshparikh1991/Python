import tweepy
import xlsxwriter
import unidecode
auth = tweepy.auth.OAuthHandler('nUic4g2T6bfMVEGDivadoxpml', 'EfZyFbizw0Unzd2ILzUHmb6C8B7F6tWwKdeAuuGNDxD0PBvML6')
auth.set_access_token('836243372960423937-pJaxhGs6Dxii8vV1Cu6h3M9J4HIRaj5', 'uOZprMqD4hMAwBdzQjn7hER6C3qMPBOyhobJFgNRpQh9Q')

row = 1
workbook = xlsxwriter.Workbook('tweet_list_mcdonalds.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write(0, 0, "Sr No")
worksheet.write(0, 1, "Date")
worksheet.write(0, 2, "User Name")
worksheet.write(0, 3, "Tweet")
worksheet.write(0, 4, "In Reply to")
worksheet.write(0, 5, "Retweet")
worksheet.write(0, 6, "Favorite")

try:
    api = tweepy.API(auth)

    for tweet in tweepy.Cursor(api.search, 
                        q="vassar brother", 
                        lang="en").items():
        if (row >5000):
            break
        format7 = workbook.add_format({'num_format': 'mmm d yyyy hh:mm AM/PM'})
        in_reply_to_screen_name = ""
        created_at = tweet.created_at
        text = unidecode.unidecode(tweet.text)
        uID = unidecode.unidecode(tweet.user.name)
        if tweet.in_reply_to_screen_name is not None:
            in_reply_to_screen_name = unidecode.unidecode(tweet.in_reply_to_screen_name)
        
        retweet_count = tweet.retweet_count
        favorite_count = tweet.favorite_count
        
        worksheet.write(row, 0, row)
        worksheet.write(row, 1, created_at, format7)
        worksheet.write(row, 2, uID)
        worksheet.write(row, 3, text)
        worksheet.write(row, 4, in_reply_to_screen_name)
        worksheet.write(row, 5, retweet_count)
        worksheet.write(row, 6, favorite_count)

        row = row + 1
    worksheet.freeze_panes(1, 0)
    workbook.close()
except Exception as e:
    print(e)
finally:
    workbook.close()

    
import win32com.client as win32
import os

fPath = os.path.dirname(os.path.abspath(__file__))
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(fPath + '\\tweet_list_mcdonalds.xlsx')
ws = wb.Worksheets("Sheet1")
ws.Columns.AutoFit()
wb.Save()
excel.Application.Quit()

print("collected: " + str(row) + " tweets!!!")