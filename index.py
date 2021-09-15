print("+----------------------------------------------------------------+")
print("|     Extract infrmation of given product by link                |")
print("+----------------------------------------------------------------+")
###########################################################################################
#          Import                                                                         #
###########################################################################################
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from time import time,sleep
from os import system
import pandas as pd
#=========================================================================================#
#          important function                                                             #
#=========================================================================================# 
# Use to show program progress
def printC(n,string):
    system("cls")
    print("+----------------------------------------------------------------+")
    print("|     Extract infrmation of given product by link                |")
    print("+----------------------------------------------------------------+")
    print("Progress  : ",str(n)+"/6")
    print(string)
# use to click on button
def click(driver,css,time_ = 60):
    if(time()-initComment >= timeComment and timeComment):
            return None
    try:
        element = WebDriverWait(driver, time_).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR , css))
        )
        element.click()
    except:
        sleep(2)
        click(driver,css)
# use to sum of string in a list
def strSum(li,between=""):
    sumStr = ""
    for x in li:
        sumStr = sumStr+str(x)+between
    if(len(between)):
        return sumStr[:-len(between)]
    return sumStr
# extract all comments from a page
def OnepageComments(driver):
    sleep(4)
    nameS = [x.text for x in driver.find_elements_by_css_selector("#cm_cr-review_list > div > div > div > div:nth-child(1)")]
    ratingS = [x.get_attribute("class")[26] for x in driver.find_elements_by_css_selector("#cm_cr-review_list > div > div > div > div:nth-child(2) > a:nth-child(1) > i")]
    reviewTitleS = [x.text for x in driver.find_elements_by_css_selector("#cm_cr-review_list > div > div > div > div:nth-child(2) > a.a-size-base.a-link-normal.review-title.a-color-base.review-title-content.a-text-bold > span")]
    dateS = [strSum(x.text.split()[-3:]," ") for x in driver.find_elements_by_css_selector("#cm_cr-review_list > div > div > div > span")]
    bodyS = [x.text for x in driver.find_elements_by_css_selector("#cm_cr-review_list > div > div > div > div.a-row.a-spacing-small.review-data > span > span")]
    
    onePageData = []
    for x in range(len(nameS)):
        onePageData.append([nameS[x],ratingS[x],reviewTitleS[x],dateS[x],bodyS[x]])
    return onePageData

# main functions
def amazonProductInfo(link,fileName,timeComment):
    #=========================================================================================#
    #          Amazon website                                                                 #
    #=========================================================================================#
    printC(1,"Started")
    amazonProduct = webdriver.Chrome(ChromeDriverManager().install())
    amazonProduct.get(link)
    amazonProduct.implicitly_wait(30)
    #=========================================================================================#
    #          Basic information                                                              #
    #=========================================================================================#
    printC(2,"Extracting basic information")
    try:
        productName = amazonProduct.find_element_by_css_selector("#title").text
    except :
        return False
    print("title")
    try:
        productCost = amazonProduct.find_element_by_css_selector("#priceblock_ourprice").text
    except:
        try:
            productCost = amazonProduct.find_element_by_css_selector("#priceblock_dealprice").text
        except:
            productCost = "Not available"
    print("Price")
    aboutItems = strSum([x.text for x in amazonProduct.find_elements_by_css_selector("#feature-bullets > ul > li")],"""
""")
    print('About items')
    ratingCount = amazonProduct.find_element_by_css_selector("#reviewsMedley > div > div.a-fixed-left-grid-col.a-col-left > div.a-section.a-spacing-none.a-spacing-top-mini.cr-widget-ACR > div.a-row.a-spacing-medium.averageStarRatingNumerical > span").text.split()[0]
    if(fileName!=""):
        fileName =fileName.replace("/"," or ").replace("&"," or ").replace("|"," or ")
    else:
        fileName =productName.replace("/"," or ").replace("&"," or ").replace("|"," or ")
    #=========================================================================================#
    #          Rating                                                                         #
    #=========================================================================================#
    printC(3,"Extracting rating")
    overallRating = amazonProduct.find_element_by_css_selector("#reviewsMedley > div > div.a-fixed-left-grid-col.a-col-left > div.a-section.a-spacing-none.a-spacing-top-mini.cr-widget-ACR > div.a-fixed-left-grid.AverageCustomerReviews.a-spacing-small > div > div.a-fixed-left-grid-col.aok-align-center.a-col-right > div > span > span").text
    ratingPercantage = [amazonProduct.find_element_by_css_selector(f"#histogramTable > tbody > tr:nth-child({x}) > td.a-text-right.a-nowrap > span.a-size-base").text for x in range(1,6)]
    #=========================================================================================#
    #          Screenshot                                                                     #
    #=========================================================================================#
    printC(4,"Saving screenshot of product")
    amazonProduct.execute_script('document.querySelector("#productTitle").scrollIntoView()')
    amazonProduct.save_screenshot(f"{fileName}.png")
    #=========================================================================================#
    #          Comment                                                                        #
    #=========================================================================================#
    printC(5,"Extracting Comments (takes time)")
    global initComment
    initComment = time()
    click(amazonProduct,"#reviews-medley-footer > div.a-row.a-spacing-medium > a")
    amazonProduct.implicitly_wait(30)
    commentData = []
    page = 1
    while(1==1):
        if(time()-initComment >= timeComment and timeComment):
            print("timeout")
            break
        print(f"Page number : {page}")
        commentData+=OnepageComments(amazonProduct)
        try:
            disable=amazonProduct.find_element_by_css_selector(".a-disabled")
            try:
                if(amazonProduct.find_element_by_css_selector("#cm_cr-pagination_bar > ul > li:nth-child(2)")==disable):
                    break
            except:
                break
        except:
            "Nothing to do"
        click(amazonProduct,"#cm_cr-pagination_bar > ul > li:nth-child(2)")
        amazonProduct.implicitly_wait(20)
        page+=1
    amazonProduct.quit()
    #=========================================================================================#
    #          Saving data                                                                    #
    #=========================================================================================#
    printC(6,"Saving data")
    basicData = [["Product name",productName],["Product cost",productCost],["Rating",overallRating],["Number of global ratings",ratingCount],["About items",aboutItems]]
    ratingData = [[str(x)+' star rating',ratingPercantage[::-1][x-1]] for x in range(5,0,-1)]
    dfBasicData = pd.DataFrame(basicData, columns = ['Feature ', 'Value'])
    dfRatingData = pd.DataFrame(ratingData, columns = ['Rating', 'Percentage of people'])
    dfCommentData = pd.DataFrame(commentData, columns = ['Name', 'Rating',"Review title","Date","Comment"])
    xlsx = pd.ExcelWriter(f'{fileName}.xlsx', engine='xlsxwriter')
    dfBasicData.to_excel(xlsx, sheet_name='Basic information')
    dfRatingData.to_excel(xlsx, sheet_name='Rating Data')
    dfCommentData.to_excel(xlsx, sheet_name='Commnet data')
    xlsx.save()
    return True
link = input("Enter the link address    :    ")
fileName = input("Enter file name without .xlsx    :    ")
timeComment = input("""Specify the specific time (in seconds) for extracting comments/reviews, this will save time
but leave it blank if you want to extract all comments/reviews    :    """)
if(timeComment==""):
    timeComment = False
else:
    timeComment = int(timeComment) 
if(amazonProductInfo(link,fileName,timeComment)):
    print("ok")
    printC(6,"Completed")
else:
    print("Some Error")
input()
