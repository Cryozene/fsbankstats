#basic utilities
import numpy as np
import math
import os.path
from datetime import date, datetime, timedelta
import pytz
import re

#tesseract for reading images
import pytesseract
import cv2 as cv

#Gui & screenshots
import tkinter as tk
import pyautogui

#google API
import gspread
from oauth2client.service_account import ServiceAccountCredentials


#sheet and starting date
###TEST Sheet APITest
#workbook_id = 'sheetID_here'
#sheet_name = 'Sheet1'

###CrybabyDonationsAnon
workbook_id = 'sheetID_here'
sheet_name = 'RawDonationData'
starting_date = datetime.fromisoformat('2020-02-07')-timedelta(hours=2)

#path to pytesseract executable
pytesseract.pytesseract.tesseract_cmd = r'path_to_pytesseract_executable'

#path to google OAuth2 credentials
credential_path = 'credentials.json'

#offset factor to text relative rank symbol
shift_right = 1.5
extend_right = 9.2

class MainGUI:


    def __init__(self):
        #start tkinter instance
        self.root = tk.Tk()
        self.root.wm_attributes("-topmost", 1)

        #initialize membership symbols as templates and donation pairs
        self.temp=r'temp/temp.png'
        self.templates = ['Member.png', 'Leader.png', 'Rookie.png', 'Officer.png']
        self.donation_pairs={}

        #create GUI
        self.createCapturedDonationCounter()
        self.canvas1 = tk.Canvas(self.root, width = 300, height = 250)
        self.canvas1.pack()
        self.createScreenshotButton()
        self.createUploadButton()

        #ready
        self.root.mainloop()

    def createScreenshotButton(self):
        button = tk.Button(text='Take Screenshot', command=self.takeScreenshot, bg='green',fg='white',font= 10)
        self.canvas1.create_window(150, 150, window=button)

    def createUploadButton(self):
        button = tk.Button(text='Upload Data', command=self.uploadData, bg='grey',fg='white',font= 10)
        self.canvas1.create_window(150, 200, window=button)
        
    def createCapturedDonationCounter(self):
        self.capturedLabel = tk.Label(self.root, font='Helvetica 16 bold')
        self.updateConfirmationLabel()
        self.capturedLabel.pack(side=tk.TOP)
        
    def updateConfirmationLabel(self):
        self.capturedLabel['text'] = str(len(self.donation_pairs)) + ' members captured'

    #returns the locations of a given template within a screenshot
    def getTemplateLoc(self, template, threshold=0.8):
        img_rgb = cv.imread(self.temp)
        img_gray = cv.cvtColor(img_rgb, cv.COLOR_BGR2GRAY)
        template = cv.imread(template,0)
        res = cv.matchTemplate(img_gray,template,cv.TM_CCOEFF_NORMED)
        return np.where( res >= threshold)

    #returns a tuple with (name, gc_amount) from a template
    def getImgText(self, loc, template, extend_right=extend_right, shift_right=shift_right,ptdiff=5):
        img_rgb = cv.imread(self.temp)
        template = cv.imread(template,0)
        w, h = template.shape[::-1]
        extend_right = round(w*extend_right)
        shift_right = round(w*shift_right)
        count=0
        out=[]
        for pt in zip(*loc[::-1]):
            if count>0:
                wdiff = math.pow(pt[0]-last_pt[0], 2)
                hdiff = math.pow(pt[1]-last_pt[1], 2)
                diff=math.sqrt(wdiff+hdiff)
            else:
                diff=math.inf
            last_pt=pt
            if ptdiff<=diff:
                crop_img = img_rgb[pt[1]:pt[1]+h, pt[0]+shift_right:pt[0]+extend_right].copy()     
                try: 
                    out.append(pytesseract.image_to_string(crop_img))
                except:
                    pass
            count+=1
        return out

    ##returns a list of tuples with (name, gc_amount)
    def getDonationPairs(self, threshold=0.8, extend_right=1):
        out = []
        temp = []
        for template in self.templates:
            loc = self.getTemplateLoc(template, threshold)
            temp = temp + self.getImgText(loc, template, extend_right=extend_right)
        for line in temp:
            try: 
                print(line)
                #captures any comma-separated number at then end of the string in the format 
                donation = re.search(r'(\d{1,3},)*\d{1,3}$', line, ).group(0)
                #captures any names including multiple words, stopping at & excluding words only consisting of numbers and/or commas, 
                #technical filter at the start of a line: may start with leftover textures that tesseract detects as characters ")" or "|"
                name = re.search(r'^(?:[|)]? ?)\w+(\s\w+[^\d\W,]\w+)*', line).group(0)
                while name[0] in " |)":
                    name = name[1:]
                donation = int(donation.replace(',', ''))
                print([name.lower(), donation])
                out.append([name.lower(), donation])
            except Exception as e:
                print ("Cant'tokenize Line: " + line)
                print (e)
        return out

    #takes a screenshot, reads bank stats and updates self.getDonationPairs
    def takeScreenshot (self):
        screenshot = pyautogui.screenshot()
        screenshot.save(self.temp)
        newPairs = self.getDonationPairs(extend_right=extend_right)
        for pair in newPairs:
            self.donation_pairs[pair[0]] = pair[1]
        self.updateConfirmationLabel()

    #uploads data to google sheets, custom function for CrybabyDonations sheet
    def uploadData(self):
        scopes=['https://www.googleapis.com/auth/drive',
                'https://www.googleapis.com/auth/drive.file'
        ]

        creds = ServiceAccountCredentials.from_json_keyfile_name(credential_path,scopes)
        client = gspread.authorize(creds)
        wb = client.open_by_key(workbook_id)
        sheet= wb.worksheet(sheet_name)
        memberlist=sheet.row_values(1)
        memberlist = [x.lower() for x in memberlist]
        daydelta = datetime.now(pytz.utc)-pytz.utc.localize(starting_date)
        daydelta = round(daydelta.total_seconds()/timedelta(days=1).total_seconds())
        donationRow = daydelta + 4
        #memberMaxIndex=None
        for name, value in self.donation_pairs.items():
            try:
                column = memberlist.index(name)+1
            except ValueError:
                #self.insertNewMemberColumns(wb, sheet, memberlist)
                #print('inserted new member')
                print("member not found: " + name + " Value: " + str(value))
                continue
            sheet.update_cell(donationRow, column, value)
        self.capturedLabel['text'] = 'Done!'

    def insertNewMemberColumns(self, wb, sheet, memberlist):
        try:
            memberMaxIndex = int(memberlist[memberlist.index('maxindex')+1])+2
        except (TypeError, ValueError):
            raise ValueError('failed to find MaxIndex')
        self.insertNewColumn(wb, sheet.id, len(memberlist))
        
        self.insertNewColumn(wb, sheet.id, memberMaxIndex)

    def insertNewColumn(self, wb, sheetid, startindex):
        requests = []
        requests.append({
            "insertDimension": {
                "range": {
                "sheetId": sheetid,
                "dimension": "COLUMNS",
                "startIndex": startindex,
                "endIndex": startindex +1,
                },
                "inheritFromBefore": True
            }
            })

        body = {
            'requests': requests
        }
        wb.batch_update(body)

def main():
    gui = MainGUI()
    

if __name__ == '__main__':
    main()
