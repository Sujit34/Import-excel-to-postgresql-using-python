import psycopg2
import xlrd

book = xlrd.open_workbook("proposed.xlsx")
sheet = book.sheet_by_name("Proposed")

database = psycopg2.connect (database = "newdata", user="postgres", password="root", host="localhost", port="5432")

cursor = database.cursor()


query = """INSERT INTO app_basedata (uniqueid, eduid, district, upazilla, s_union, mouza, village, sheltername, northing, easting, distance, expectedpopulation, waterlevel, pucca, semipucca, wooden, bamboo, thatched, shanty, total_house) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"""

for r in range(1, sheet.nrows):
    uniqueid = str(sheet.cell(r,1).value)
    eduid = str(sheet.cell(r,2).value)
    district = str(sheet.cell(r,3).value)
    upazilla = str(sheet.cell(r,4).value)
    s_union = str(sheet.cell(r,5).value)
    mouza = str(sheet.cell(r,6).value)
    village = str(sheet.cell(r,7).value)
    sheltername = str( sheet.cell(r,8).value)  
    northing = str(sheet.cell(r,9).value)
    easting = str(sheet.cell(r,10).value)
    distance = str(sheet.cell(r,11).value)
    expectedpopulation = str(sheet.cell(r,12).value)
    waterlevel = str(sheet.cell(r,13).value)
    pucca = str(sheet.cell(r,14).value)
    semipucca = str(sheet.cell(r,15).value)
    wooden = str(sheet.cell(r,16).value)
    bamboo = str(sheet.cell(r,17).value)
    thatched = str(sheet.cell(r,18).value)
    shanty = str(sheet.cell(r,19).value)
    total_house = str(sheet.cell(r,20).value)
    

    values = (uniqueid, eduid, district, upazilla, s_union, mouza, village, sheltername, northing, easting, distance, expectedpopulation, waterlevel, pucca, semipucca, wooden, bamboo, thatched, shanty,total_house)    					     

    cursor.execute(query, values)

cursor.close()

database.commit()

database.close()

print ("Successfully imported.")
