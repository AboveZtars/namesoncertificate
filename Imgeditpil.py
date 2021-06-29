from PIL import Image , ImageEnhance
from PIL import ImageFont
from PIL import ImageDraw 
import openpyxl

#Open workbook and sheet with data
doc = openpyxl.load_workbook ("YOURFILE.xlsx")
sheet = doc.get_sheet_by_name("SHEET")

#Save data in list
names=[]
lastnames=[]
levels= []
for row in sheet.iter_rows():
    name = row[0].value
    lastname = row[1].value
    level = row[2].value
    #Verify empty slots
    if name != None and lastname != None:
        names.append(name)
        lastnames.append(lastname)
        levels.append(level)

c=0

for personName,personLastname in zip(names,lastnames):
    #Open image to set fullname to and create a copy to work with
    cert = Image.open("YOURFILETOADDDATA.png")
    cert_copy = cert.copy()

    #font
    draw = ImageDraw.Draw(cert_copy)
    #font = ImageFont.truetype(<font-file>, <font-size>)
    font = ImageFont.truetype("OpenSans-Bold.ttf", 90)
    font2 = ImageFont.truetype("OpenSans-Bold.ttf", 60)
    #draw.text((x, y),"Sample Text",(r,g,b))
    fullname = personName+" "+personLastname
    day = "1"
    month = "April"
    year = "21"
    
    #Calculate position
    W, H = cert.size
    w, h = draw.textsize(fullname, font=font)
    
    #(W-w)/2 => to center in x, (H-h)/2 to center in y
    draw.text(((W-w)/2,405),fullname,(0,0,0),font=font )  
    draw.text((440,575),day,(0,0,0),font=font2 ) #day position 
    draw.text((670,575),month,(0,0,0),font=font2 ) #month position
    draw.text((1090,575),year,(0,0,0),font=font2 ) #year position
    
    

    file = "yourfilename"+fullname
    
    #add a few enhance 
    contrast = ImageEnhance.Contrast(cert_copy)
    contrast.enhance(5)
    color = ImageEnhance.Color(cert_copy)
    color.enhance(2).save(file+".pdf") #you can change the type of file
   
    cert_copy.save(file+".pdf")
    c=c+1

