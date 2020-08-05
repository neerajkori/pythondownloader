#!/usr/bin/python3.6
import xlrd
import os


file_path = "/home/neeraj/Desktop/Neon_videos_links.xlsx"
download_dir = "/home/neeraj/Desktop/Neon_videos_download"

if not os.path.exists(download_dir):
    print("Download directory not exists: Creating.....", download_dir)
    os.mkdir(download_dir)
else:
    print("Download directory present.")


wb = xlrd.open_workbook(file_path)
sheet = wb.sheet_by_index(0)
for rows in range(0, sheet.nrows):
    row = sheet.row_values(rows)
    if row[3] == "Done":
        print(f"Skipping Download.. {row[0]} -part- {str(round(row[2]))}.mp4  == Already Downloaded")
        continue
    else:
        content_path = download_dir + "/" +  row[0]
        # print("content path is: ", content_path )
        if not os.path.exists(content_path):
            print(f"Content download dir not exist. Creating.. [ {row[0]} ]")
            print("Downloading...", row[0] + "-part-" + str(round(row[2])) + ".mp4")
            os.mkdir(download_dir + "/" + row[0])
            retcode = 0
            while True:
                if not retcode == 0:
                    download_url = "youtube-dl" + "  " + "-c" + row[1]
                else:
                    download_url = "youtube-dl" + "  " + row[1]
                print(download_url)
                os.chdir(content_path)
                print(content_path)
                retcode =  os.system(download_url)
                if retcode == 0:
                    print(f"{row[0]} -part- {str(round(row[2]))}.mp4 .. Download Successfully.")
                    break
                else:
                    print(f"Resuming download: {row[0]} -part- {str(round(row[2]))}.mp4")
                    
        else:
            # print("Content download dir not exist. Creating.. ", row[0])
            print("Downloading...", row[0] + "-part-" + str(round(row[2])) + ".mp4")
            retcode = 0
            while True:
                if not retcode == 0:
                    download_url = "youtube-dl" + "  " + "-c " + row[1]
                else:
                    download_url = "youtube-dl" + "  " + row[1]
                print(download_url)
                os.chdir(content_path)
                print(content_path)
                retcode =  os.system(download_url)
                if retcode == 0:
                    print(f"{row[0]} -part- {str(round(row[2]))}.mp4 .. Download Successfully.")
                    break
                else:
                    print(f"Resuming download: {row[0]} -part- {str(round(row[2]))}.mp4")
                
