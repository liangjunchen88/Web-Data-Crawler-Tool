# -*- encoding:utf8 -*-
from bs4 import BeautifulSoup
import re
import urllib.request,urllib.error
import xlwt
import sqlite3


def main():
    baseurl="https://movie.douban.com/top250?start="
    #1. Crawl the web
    datalist = getData(baseurl)
    #2。Get the data
    savepath = "Douban Movie Top250.xls"
    #3. Save the data
    saveData(datalist,savepath)
    #askURL("https://movie.douban.com/top250?start=")

#Rules for video detail links

#Create a regular expression object, representing the rule (the pattern of the string)
findLink = re.compile(r'<a href="(.*?)">')
#videoimage
findImgSrc = re.compile(r'<img.*src="(.*?)"',re.S) #re.S make newlines included
#video title
findTitle = re.compile(r'<span class="title">(.*)</span>')
#video rating
findRating = re.compile(r'<span class="rating_num" property="v:average">(.*)</span>')
#The number of reviewers
findJudge = re.compile(r'<span>(\d*)人评价</span>')
#find overview
findInq = re.compile(r'<span class="inq">(.*)</span>')
#find related content of the video
findBd = re.compile(r'<p class="">(.*?)</p>',re.S)

#crawl web pages
def getData(baseurl):
    datalist=[]
    for i in range(0,10):#Call the function to get page information 10 times
        url = baseurl + str(i*25)
        html = askURL(url)    #Save the obtained webpage source code
        # 2. Parse the data one by one
        soup = BeautifulSoup(html,"html.parser")
        for item in soup.find_all('div', class_="item") :  #Find a string that meets the requirements and form a list
            #print(item) #Test, view all information about the movie item
            data = [] #Save all information about a movie
            item = str(item)

            #video details link
            link = re.findall(findLink,item)[0]
            #re library is used to find the specified string through regular expressions

            data.append(link)                   #Add a link

            imgSrc = re.findall(findImgSrc,item)[0]
            data.append(imgSrc)                 #Add pictures

            titles = re.findall(findTitle,item) #The title may only have a Chinese name, no foreign name
            if(len(titles)==2):
                ctitle = titles[0]
                data.append(ctitle)              #Add Chinese name
                otitle = titles[1].replace("/","") # remove irrelevant symbols
                data.append(otitle)              #Add foreign name
            else:
                data.append(titles[0])
                data.append(' ')               #Foreign name leave blank

            rating = re.findall(findRating,item)[0]
            data.append(rating)                #add rating

            judgeNum = re.findall(findJudge,item)[0]
            data.append(judgeNum)              #Add the number of comments

            inq = re.findall(findInq,item)
            if len(inq)!=0:
                inq = inq[0].replace("。","/") # remove the period
                data.append(inq)                 #add overview
            else:
                data.append(" ")   #leave blank

            bd = re.findall(findBd,item)[0]
            bd = re.sub('<br(\s+)?/>(\s+)?'," ",bd)  #remove<br/>
            info = re.findall(r'(\d+.*)',bd)
            infoList = info[0].split('/')
            #print(infoList)
            year = infoList[0].replace(" ","")
            data.append(year)              #add year
            coun = infoList[1].split(' ')
            data.append(coun[0])              #add country
            type = infoList[2]
            data.append(type)              #add type
            #bd = re.sub('/'," ",bd) #replace/
            #bd = re.sub('...', "/", bd) #replace...
            #data.append(bd.strip())  # remove leading and trailing spaces

            datalist.append(data)   #Put the processed movie information into the datalist

    print(datalist)
    return datalist



#Get the content of the web page that specifies a URL
def askURL(url):
    head={#Simulate browser header information: send a message to the Douban server
         "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36"

    }
    #useragent: tells the browser what level of information we can accept
    request = urllib.request.Request(url,headers=head)
    html = ""
    try:
        response = urllib.request.urlopen(request)
        html = response.read().decode("utf-8")
        #print(html)
    except urllib.error.URLError as e:
        if hasattr(e,"code"):
            print(e.code)
        if hasattr(e,"reason"):
            print(e.reason)

    return html


def saveData(datalist,savepath):
    print("save....")
    book = xlwt.Workbook(encoding="utf-8",style_compression=0)  # Create workbook object
    sheet = book.add_sheet('Douban Movie Top250',cell_overwrite_ok=True)  # create worksheet
    col =("Movie details link","image link","Chinese name of the video",
          "Foreign language name of the film","Score","Number of evaluations",
          "Overview","Era","Nation","Type")
    for i in range(0,10):
        sheet.write(0,i,col[i])  #column name
    for i in range(0,250):
        print("%d th line" %(i+1))
        data = datalist[i]
        for j in range(0,10):
            sheet.write(i+1,j,data[j]) #data

    book.save(savepath)  #save data table

if __name__ == '__main__':
    main()
    print("Crawling is complete!")
