import psycopg2
import xlrd

book = xlrd.open_workbook("data.xlsx")
sheet = book.sheet_by_name("ECRRP")

database = psycopg2.connect (database = "newdb", user="postgres", password="root", host="localhost", port="5432")

cursor = database.cursor()


query = """INSERT INTO app_basedata (uniqueid, schoolid, district, upazilla, s_union, mouza, village, sheltername ,constructionyear, northing, easting, fundingagencies, usesas, condition, noofvillage, acc_cap_per, acc_cap_livestock, plintharea, landarea, maletoilet, femaletoilet, watersupplystatus, watersupplysource, septictank, ramp, population, distance, expectedpopulation, warningsystem, warningsystemdetails, existing, underconstruction, ground, g_others, first, f_ohter, second, s_others, third, t_others, rs_ground, rs_first, rs_second, rs_roof, rs_beam, rs_column, rs_wall, rs_window, rs_door, rs_toilet, rs_stair, rs_ramp, rs_others, rs_ifothersfillup, communicationsource, cs_others, powerstation, ps_others, waterreservior, foodstoragesystem, additionaltoilet, septiktanksoakwell, others, polderwashedaway, polderrepaired, polderneedrepairs, road1from, road2from, road3from, road4from, road1to, road2to, road3to, road4to, road1name, road2name, road3name, road4name, road1type, road2type, road3type, road4type, road1id, road2id, road3id, road4id, road1length, road2length, road3length, road4length, road1pavement, road2pavement, road3pavement, road4pavement, road1embankment, road2embankment, road3embankment, road4embankment, road1kachcha, road2kachcha, road3kachcha, road4kachcha, road1bc, road2bc, road3bc, road4bc, road1other, road2other, road3other, road4other, road1surfacing, road2surfacing, road3surfacing, road4surfacing, road1basecourse, road2basecourse, road3basecourse, road4basecourse, road1sbase, road2sbase, road3sbase, road4sbase, road1roadcovered, road2roadcovered, road3roadcovered, road4roadcovered, road1avgdepth, road2avgdepth, road3avgdepth, road4avgdepth, e1_a, e1_b, e1_c, e2_a, e2_b, e2_c, e3_a, e3_b, e3_c, e4_a, e4_b, e4_c, e5_a, e5_b, e5_c, e6_a, e6_b, e6_c, e7_a, e7_b, e7_c, e8_a, e8_b, e8_c, e9_a, e9_b, e9_c) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"""

for r in range(1, sheet.nrows):
    uniqueid = str(sheet.cell(r,0).value)
    schoolid = str(sheet.cell(r,1).value)
    district = str(sheet.cell(r,2).value)
    upazilla = str(sheet.cell(r,3).value)
    s_union = str(sheet.cell(r,4).value)
    mouza = str(sheet.cell(r,5).value)
    village = str(sheet.cell(r,6).value)
    sheltername =str( sheet.cell(r,7).value)
    constructionyear = str(sheet.cell(r,8).value)
    northing = str(sheet.cell(r,9).value)
    easting = str(sheet.cell(r,10).value)
    fundingagencies = str(sheet.cell(r,11).value)
    usesas = str(sheet.cell(r,12).value)
    condition = str(sheet.cell(r,13).value)
    noofvillage = str(sheet.cell(r,14).value)
    acc_cap_per = str(sheet.cell(r,15).value)
    acc_cap_livestock = str(sheet.cell(r,16).value)
    plintharea = str(sheet.cell(r,17).value)
    landarea = str(sheet.cell(r,18).value)
    maletoilet = str(sheet.cell(r,19).value)
    femaletoilet =str( sheet.cell(r,20).value)
    watersupplystatus =str( sheet.cell(r,21).value)
    watersupplysource =str( sheet.cell(r,22).value)
    septictank =str( sheet.cell(r,23).value)
    ramp = str(sheet.cell(r,24).value)
    population =str( sheet.cell(r,25).value)
    distance = str(sheet.cell(r,26).value)
    expectedpopulation = str(sheet.cell(r,27).value)
    warningsystem = str(sheet.cell(r,28).value)
    warningsystemdetails =str( sheet.cell(r,29).value)
    existing = str(sheet.cell(r,30).value)
    underconstruction =str( sheet.cell(r,31).value)
    ground = str(sheet.cell(r,32).value)
    g_others = str(sheet.cell(r,33).value)
    first = str(sheet.cell(r,34).value)
    f_ohter =str( sheet.cell(r,35).value)
    second = str(sheet.cell(r,36).value)
    s_others =str( sheet.cell(r,37).value)
    third = str(sheet.cell(r,38).value)
    t_others = str(sheet.cell(r,39).value)
    rs_ground =str( sheet.cell(r,40).value)
    rs_first = str(sheet.cell(r,41).value)
    rs_second = str(sheet.cell(r,42).value)
    rs_roof = str(sheet.cell(r,43).value)
    rs_beam = str(sheet.cell(r,44).value)
    rs_column  = str(sheet.cell(r,45).value)
    rs_wall =str( sheet.cell(r,46).value)
    rs_window = str(sheet.cell(r,47).value)
    rs_door =str( sheet.cell(r,48).value)
    rs_toilet = str(sheet.cell(r,49).value)
    rs_stair = str(sheet.cell(r,50).value)
    s_ramp = str(sheet.cell(r,51).value)
    rs_others = str(sheet.cell(r,52).value)
    rs_ifothersfillup = str(sheet.cell(r,53).value)
    communicationsource  = str(sheet.cell(r,54).value)
    cs_others =str( sheet.cell(r,55).value)
    powerstation  = str(sheet.cell(r,56).value)
    ps_others = str(sheet.cell(r,57).value)
    waterreservior =str( sheet.cell(r,58).value)
    foodstoragesystem  =str( sheet.cell(r,59).value)
    additionaltoilet   = str(sheet.cell(r,60).value)
    septiktanksoakwell = str(sheet.cell(r,61).value)
    others   = str(sheet.cell(r,62).value)
    polderwashedaway = str(sheet.cell(r,63).value)
    polderrepaired  =str( sheet.cell(r,64).value)
    polderneedrepairs  =str( sheet.cell(r,65).value)
    road1from = str(sheet.cell(r,66).value)
    road2from = str(sheet.cell(r,67).value)
    road3from = str(sheet.cell(r,68).value)
    road4from = str(sheet.cell(r,69).value)
    road1to = str(sheet.cell(r,70).value)   
    road2to = str(sheet.cell(r,71).value)  
    road3to = str(sheet.cell(r,72).value)   
    road4to  = str(sheet.cell(r,73).value)
    road1name = str(sheet.cell(r,74).value)
    road2name  = str(sheet.cell(r,75).value)
    road3name  = str(sheet.cell(r,76).value)
    road4name   = str(sheet.cell(r,77).value)
    road1type  = str(sheet.cell(r,78).value)
    road2type  = str(sheet.cell(r,79).value)
    road3type  =str( sheet.cell(r,80).value)
    road4type  = str(sheet.cell(r,81).value)
    road1id  = str(sheet.cell(r,82).value)
    road2id  = str(sheet.cell(r,83).value)
    road3id  =str( sheet.cell(r,84).value)
    road4id  = str(sheet.cell(r,85).value)
    road1length  = str(sheet.cell(r,86).value)
    road2length =str( sheet.cell(r,87).value)
    road3length  =str( sheet.cell(r,88).value)
    road4length  = str(sheet.cell(r,89).value)
    road1pavement =str( sheet.cell(r,90).value) 
    road2pavement =str( sheet.cell(r,91).value)
    road3pavement = str(sheet.cell(r,92).value) 
    road4pavement = str(sheet.cell(r,93).value)
    road1embankment = str(sheet.cell(r,94).value)
    road2embankment  =str( sheet.cell(r,95).value) 
    road3embankment =str( sheet.cell(r,96).value)
    road4embankment =str( sheet.cell(r,97).value)  
    road1kachcha  = str(sheet.cell(r,98).value) 
    road2kachcha  =str( sheet.cell(r,99).value) 
    road3kachcha  =str( sheet.cell(r,100).value)  
    road4kachcha =str( sheet.cell(r,101).value)
    road1bc   = str(sheet.cell(r,102).value)
    road2bc =str( sheet.cell(r,103).value)
    road3bc  =str( sheet.cell(r,104).value) 
    road4bc  = str(sheet.cell(r,105).value)
    road1other = str(sheet.cell(r,106).value)
    road2other = str(sheet.cell(r,107).value)
    road3other  = str(sheet.cell(r,108).value)
    road4other  = str(sheet.cell(r,109).value) 
    road1surfacing = str(sheet.cell(r,110).value)
    road2surfacing = str(sheet.cell(r,111).value)
    road3surfacing = str(sheet.cell(r,112).value)
    road4surfacing   = str(sheet.cell(r,113).value)
    road1basecourse =str( sheet.cell(r,114).value)
    road2basecourse  = str(sheet.cell(r,115).value)
    road3basecourse  = str(sheet.cell(r,116).value) 
    road4basecourse   = str(sheet.cell(r,117).value)
    road1sbase   = str(sheet.cell(r,118).value) 
    road2sbase  = str(sheet.cell(r,119).value) 
    road3sbase  = str(sheet.cell(r,120).value)
    road4sbase   = str(sheet.cell(r,121).value)
    road1roadcovered  =str(sheet.cell(r,122).value)
    road2roadcovered  = str(sheet.cell(r,123).value)
    road3roadcovered = str(sheet.cell(r,124).value)
    road4roadcovered = str(sheet.cell(r,125).value) 
    road1avgdepth  = str(sheet.cell(r,126).value)  
    road2avgdepth  = str(sheet.cell(r,127).value) 
    road3avgdepth  =str( sheet.cell(r,128).value) 
    road4avgdepth  =str( sheet.cell(r,129).value)
    e1_a   = str(sheet.cell(r,130).value)
    e1_b  = str(sheet.cell(r,131).value)
    e1_c = str(sheet.cell(r,132).value)
    e2_a  = str(sheet.cell(r,133).value)
    e2_b = str(sheet.cell(r,134).value)
    e2_c  = str(sheet.cell(r,135).value)
    e3_a  = str(sheet.cell(r,136).value)
    e3_b  = str(sheet.cell(r,137).value)
    e3_c  = str(sheet.cell(r,138).value)
    e4_a   = str(sheet.cell(r,139).value)
    e4_b   = str(sheet.cell(r,140).value)
    e4_c   = str(sheet.cell(r,141).value)
    e5_a   = str(sheet.cell(r,142).value)
    e5_b   = str(sheet.cell(r,143).value) 
    e5_c   =str( sheet.cell(r,144).value) 
    e6_a   = str(sheet.cell(r,145).value)
    e6_b   =str( sheet.cell(r,146).value)
    e6_c   = str(sheet.cell(r,147).value)
    e7_a   =str( sheet.cell(r,148).value)
    e7_b   = str(sheet.cell(r,149).value)
    e7_c   =str( sheet.cell(r,150).value) 
    e8_a   =str( sheet.cell(r,151).value)
    e8_b   =str( sheet.cell(r,152).value)
    e8_c   =str( sheet.cell(r,153).value) 
    e9_a   =str( sheet.cell(r,154).value)
    e9_b   =str( sheet.cell(r,155).value)
    e9_c   = str(sheet.cell(r,156).value)

    values = (uniqueid,schoolid,district,upazilla,s_union,mouza,village,sheltername,constructionyear,northing,easting,fundingagencies,usesas,condition,noofvillage,acc_cap_per,acc_cap_livestock,plintharea,landarea,maletoilet,femaletoilet,watersupplystatus,watersupplysource,septictank,ramp,population,distance,expectedpopulation,warningsystem,warningsystemdetails,existing,underconstruction,ground,g_others,first,f_ohter,second,s_others,third,t_others,rs_ground,rs_first,rs_second,rs_roof,rs_beam,rs_column,rs_wall,rs_window,rs_door,rs_toilet,rs_stair,s_ramp,rs_others,rs_ifothersfillup,communicationsource,cs_others,powerstation,ps_others,waterreservior,foodstoragesystem,additionaltoilet,septiktanksoakwell,others,polderwashedaway,polderrepaired,polderneedrepairs,road1from,road2from,road3from,road4from,road1to,road2to,road3to,road4to,road1name,road2name,road3name,road4name,road1type,road2type,road3type,road4type,road1id,road2id,road3id,road4id,road1length,road2length,road3length,road4length,road1pavement,road2pavement,road3pavement,road4pavement,road1embankment, road2embankment,road3embankment,road4embankment,road1kachcha,road2kachcha,road3kachcha,road4kachcha,road1bc,road2bc,road3bc,road4bc,road1other,road2other,road3other,road4other,road1surfacing,road2surfacing,road3surfacing,road4surfacing,road1basecourse,road2basecourse,road3basecourse,road4basecourse,road1sbase,road2sbase,road3sbase,road4sbase,road1roadcovered,road2roadcovered,road3roadcovered,road4roadcovered,road1avgdepth,road2avgdepth,road3avgdepth,road4avgdepth,e1_a,e1_b,e1_c,e2_a,e2_b,e2_c,e3_a,e3_b,e3_c,e4_a,e4_b,e4_c,e5_a,e5_b,e5_c,e6_a,e6_b,e6_c,e7_a,e7_b,e7_c,e8_a,e8_b,e8_c,e9_a,e9_b,e9_c)    					     
    cursor.execute(query, values)

cursor.close()

database.commit()

database.close()

print ("Successfully imported.")
