//สคริปต์นี้ พัฒนามาจากสคริปต์ประเมิน SDQ นำมาปรับปปรุงให้สามารถใชบันทึกข้อมูลการเยี่ยมบ้านออนไลน์
//พัฒนาต่อยอดให้ส่งอีเมลไปหาครูประจำชั้นเมื่อส่งฟอร์มโดยนายยุทธ ขวัญเมืองแก้ว โรงเรียนฤทธิณรงค์รอน สพม.1
//สามารถส่งอีเมลไปยังครูที่ปรึกษาทันทีที่นักเรียนกรอกฟอร์มเสร็จ
//ต้องสร้างโฟลเดอร์รอเพื่อเก็บไฟล์สำเนา
//ปรับปรุงให้นำไฟล์สำเนาไปเก็บในโฟลเดอร์ที่กำหนด


const GDOC_TEMPLATE_ID =  '1F0VOySqTzZ_cZz4THnT4y8ez5wxxg7eJ644bNzXARss'  
//const GDOC_TEMPLATE_ID =  'template id'
const processSuccess = 'Process success'
const m11 = '12_ssIf-fsWH7zJe5wnDJav5kBKJ4LqqK'
const m12 = '105U1voY7LdcV1VpMT1pA4MtD4PyV_pwo'
const m13 = '1Ee32HCmLTFylhEjsbJ8gqfZoWKAGbm41'
const m14 = '1H4vwJSeAMWjhjGltRT_2TVzZ7yiYWG8p'
const m21 = '14mhwzXljhNz2q6bAt-Ay5PBYm1Ln6r0Y'
const m22 = '1ruldWNPgqjyWumqaz84ii4l772LaVj64'
const m23 = '1vImHK-lTSXK0Be-7YZOxinwx-TbvOUDv'
const m24 = '1RnEPaWM3XY5GUW7tRGsI7q7jhYH84e4v'
const m31 = '1AZLkSW8icAud-i5YL2bdPaC1sc-0PqpC'
const m32 = '1ndUnhwD0JxJSwqcBGuDxken0euwbR8RQ'
const m33 = '12I2E5I7wPnWCc__aO7OuKTA6AkyL0MZA'
const m34 = '1GumRzoOxhmCObb7mRhkhlRwSa93g8RU6'
const m41 = '1xRSvs9utR4NY5K7X4rZqtyG9YPQqq3oN'
const m42 = '1rXzBh3Dh7ttov6dIOkS-f9mZN3eaFdKU'
const m43 = '1pE2y2VJoLP10Zpcfaff8IcrJ_4V_WVgr'
const m51 = '1sI_lZNC3XmgwMxgF5bE61mUjHBCWWncL'
const m52 = '1py1ygGkbK4WkmokeLcbzq-iBHEYIxX4f'
const m53 = '1zmrOtfPNxJ9_rC1V8flTztPoq2oGHsv1'
const m61 = '1a5U_7EzMCCUrXwIE53qiUmVhaVeS72G4'
const m62 = '1wib4UPABbP9a0_aXI9mGFILMZr0W44U9'
const m63 = '1XmQOO9VUPuNNL_kyf-ZZx2XsZpWeXKpw'
const m64 = '1uaMguQ6Kat8EYkxhi2BP8SAPejEjhaBv'
//var destinationFolder = '1OJGVHUjMKjh5vYwT88CjgHhLG3SCgA8B' //folder ม.3/3
//var destinationFolder = ''
var subject = ''
var message = ''

function visitHomeRecord() { 
  //var result_emo = "";
  //var result_bahave = ""
  //var result_med = ""
  //var result_relat = ""
  //var result_socio = ""
  //var result_all = ""
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet()
  var start_row = 2
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var range = sheet.getRange(start_row, 1, lastRow-1 , lastColumn);
  range.activate()
  //var rangeValues = range.getValues() 
  var destinationFolder = ''
  
  for(i = start_row; i< lastRow +1 ; i++){ 
    var checkProcess = sheet.getRange(i,lastColumn).getValue()  //ดึงข้อมูลในคอลัมภ์สุดท้าย
    if(checkProcess!==processSuccess) {       //ถ้าข้อมูลในคอลัมสุดท้ายไม่เท่ากับ Process success จะดึงข้อมูลที่เหลือขึ้นมาทำการประมวลผล             

    //================ ข้อมูลส่วนตัวนักเรียน
    var nameTitle = sheet.getRange(lastRow,2).getValue()
    //var nameTitle = sheet.getRange(i,2).getValue() //คำนำหน้า
    var fName = sheet.getRange(i,3).getValue()     //ชื่อ
    var lName = sheet.getRange(i,4).getValue()        //นามสกุล
    //console.log(nameTitle+fName+' '+lName)
    var nName = sheet.getRange(i,5).getValue()         //ชื่อเล่น
    var stuClass = sheet.getRange(i,6).getValue()         //ระดับชั้น
    var stuRoom = sheet.getRange(i,7).getValue()         //ห้อง
    var stuOrderID = sheet.getRange(i,8).getValue()         //เลขที่
    var stuAge = sheet.getRange(i,9).getValue()         //อายุ
    var stuCitizenID = sheet.getRange(i,10).getValue()         //เลขประจำตัวประชาชน
    var orderOfKids = sheet.getRange(i,11).getValue()        //เป็นบุตรคนที่
    var quantityOfKids = sheet.getRange(i,12).getValue()        //จำนวนพี่น้องทั้งหมด
    //================ ข้อมูลที่อยู่นักเรียน
    var houseID = sheet.getRange(i,13).getValue()        //บ้านเลขที่
    var villegeID = sheet.getRange(i,14).getValue()        //หมู่
    var subRoad = sheet.getRange(i,15).getValue()        //ซอย
    var mainRoad = sheet.getRange(i,16).getValue()        //ถนน
    var subDistrict = sheet.getRange(i,17).getValue()        //แขวง/ตำบล
    var district = sheet.getRange(i,18).getValue()        //เขต/อำเภอ
    var province = sheet.getRange(i,19).getValue()        //จังหวัด
    var zipcode = sheet.getRange(i,20).getValue()        //รหัสไปรษณีย์
    var stuPhone = sheet.getRange(i,21).getValue()        //โทรศัพท์

    //================ ข้อมูลส่วนตัวบิดา
    var hasFather = sheet.getRange(i,22).getValue() //มีบิดาหรือไม่
    var fnameTitle = sheet.getRange(i,23).getValue() //คำนำหน้า
    var fatherFName = sheet.getRange(i,24).getValue()     //ชื่อ
    var fLName = sheet.getRange(i,25).getValue()        //นามสกุล
    var fOccupation = sheet.getRange(i,26).getValue()         //อาชีพ
    var fIncome = sheet.getRange(i,27).getValue()         //รายได้

    //================ ข้อมูลส่วนตัวมารดา
    var hasMother = sheet.getRange(i,28).getValue() //มีมารดาหรือไม่
    var mNameTitle = sheet.getRange(i,29).getValue() //คำนำหน้า
    var mFName = sheet.getRange(i,30).getValue()     //ชื่อ
    var mLName = sheet.getRange(i,31).getValue()        //นามสกุล
    var mOccupation = sheet.getRange(i,32).getValue()         //อาชีพ
    var mIncome = sheet.getRange(i,33).getValue()         //รายได้

    //================ ข้อมูลส่วนตัวผู้ปกครอง
    var parent = sheet.getRange(i,34).getValue()  //บุคคลที่เป็นผู้ปกครอง
    var pNameTitle = sheet.getRange(i,35).getValue() //คำนำหน้า
    var pFName = sheet.getRange(i,36).getValue()     //ชื่อ
    var pLName = sheet.getRange(i,37).getValue()        //นามสกุล
    var relationship2Student = sheet.getRange(i,38).getValue()        //ความสัมพันธ์กับนักเรียน
    var pOccupation = sheet.getRange(i,39).getValue()         //อาชีพ
    var pIncome = sheet.getRange(i,40).getValue()         //รายได้
    var pPhone = sheet.getRange(i,41).getValue()         //โทรศัพท์
    var pCitizenID = sheet.getRange(i,42).getValue()         //เลขประจำตัวประชาชน
    var pEducation = sheet.getRange(i,43).getValue()         //ระดับการศึกษา
    var familyQuantity = sheet.getRange(i,44).getValue()         //สมาชิกในครัวเรือน
    console.log('ชื่อผู้ปกครอง '+pNameTitle+pFName+' '+pLName)

    //================ ข้อมูลเฉพาะด้านครอบครัว
    var familyStatus = sheet.getRange(i,45).getValue()  //สถานภาพของบิดา-มารดา
    var stuResident = sheet.getRange(i,46).getValue() //ปัจจุบันอาศัยอยู่กับ
    var familyIncome = sheet.getRange(i,47).getValue()  //ฐานะทางเศรษฐกิจของครอบครัว
    var familyRelationships = sheet.getRange(i,48).getValue()   //ความสัมพันธ์ของสมาชิกในครอบครัว
    var goodRelationshipto = sheet.getRange(i,49).getValue()    //บุคคลในครอบครัวที่นักเรียนสนิทสนมมากที่สุด
    var studyMethod = sheet.getRange(i,50).getValue()           //วิธีการที่ผู้ปกครองเลี้ยงดูนักเรียน
    var studySupport = sheet.getRange(i,51).getValue()          //การสนับสนุนด้านการเรียน
    var studySupportMethod = sheet.getRange(i,52).getValue()     //ผู้ปกครองสนับสนุนพิเศษด้านการศึกษาด้วยวิธีใด

    //================ ข้อมูลเฉพาะด้านที่อยู่อาศัย
    var residentType = sheet.getRange(i,53).getValue()          //ลักษณะที่อยู่อาศัย
    var houseOwner = sheet.getRange(i,53).getValue()            //ความเป็นเจ้าของบ้านที่อาศัย
    var strength  = sheet.getRange(i,54).getValue()             //สภาพบ้าน
    var houseEnvi = sheet.getRange(i,55).getValue()              //สภาพแวดล้อมของบ้าน
    var communityEnvi = sheet.getRange(i,56).getValue()          //สภาพรอบบ้านหรือในชุมชน

    //================ ข้อมูลกิจกรรมที่บ้าน
    var hasHomeJob = sheet.getRange(i,57).getValue()            //การช่วยพ่อแม่ทำงานบ้าน
    var whatHomeJob = sheet.getRange(i,58).getValue()           //นักเรียนช่วยทำงานบ้านอะไรบ้าง
    var hasPartTime = sheet.getRange(i,59).getValue()           //นักเรียนมีรายได้พิเศษ
    var whatPartTime = sheet.getRange(i,60).getValue()           //นักเรียนมีรายได้พิเศษจากงานใด
    var incomePartTime = sheet.getRange(i,61).getValue()           //รายได้จากงานพิเศษคิดเป็น บาท/วัน
    var goodBehavier = sheet.getRange(i,62).getValue()          //นักเรียนมีพฤติกรรมใดที่น่าพึงพอใจบ้าง
    var noneGoodBehavier = sheet.getRange(i,63).getValue()      //นักเรียนมีพฤติกรรมใดที่ไม่น่าพึงพอใจบ้าง
    var hasStudyPlan = sheet.getRange(i,64).getValue()          //มีการวางแผนด้านการเรียนหรือไม่
    var whatStudyPlan = sheet.getRange(i,65).getValue()         //มีการวางแผนด้านการเรียนในอนาคตอย่างไร
    var parentOpinion = sheet.getRange(i,66).getValue()         //ความคิดเห็นของผู้ปกครอง
    var teacherOpinion = sheet.getRange(i,67).getValue()        //ความคิดเห็นของผู้ครูที่ปรึกษา

    //================ ข้อมูลครูที่ปรึกษา
    var t1NameTitle = sheet.getRange(i,68).getValue()           
    var t1Name = sheet.getRange(i,69).getValue()
    var t1Email = sheet.getRange(i,70).getValue()
    var t2NameTitle = sheet.getRange(i,72).getValue()
    var t2Name = sheet.getRange(i,73).getValue()
    var t2Email = sheet.getRange(i,74).getValue()

    //ข้อมูลภาพถ่าย
    var photo1 = sheet.getRange(i,75).getValue()       //แนบรูปภาพที่ได้จากการสำรวจ
    var photo2 = sheet.getRange(i,76).getValue()       //แนบรูปภาพที่ได้จากการสำรวจ
    var photo3 = sheet.getRange(i,77).getValue()       //แนบรูปภาพที่ได้จากการสำรวจ
    var photo4 = sheet.getRange(i,78).getValue()       //แนบรูปภาพที่ได้จากการสำรวจ

      
      /แยกข้อมูลแต่ละห้องไปใส่ในโฟลเดอร์ที่กำหนด 
      //ม.1
      if(stuClass == 'มัธยมศึกษาปีที่ 1' && stuRoom == 1){
        var destinationFolder = m11   //Folder ม.1/1
      } 
      if(stuClass == 'มัธยมศึกษาปีที่ 1' && stuRoom == 2){
        var destinationFolder = m12   //Folder ม.1/2
      } 
      if(stuClass == 'มัธยมศึกษาปีที่ 1' && stuRoom == 3){
        var destinationFolder = m13   //Folder ม.1/3
      }
      if(stuClass == 'มัธยมศึกษาปีที่ 1' && stuRoom == 4){
        var destinationFolder = m14   //Folder ม.1/4
      }
      //ม.2
      if(stuClass == 'มัธยมศึกษาปีที่ 2' && stuRoom == 1){
        var destinationFolder = m21   //Folder ม.2/1
      } 
      if(stuClass == 'มัธยมศึกษาปีที่ 2' && stuRoom == 2){
        var destinationFolder = m22   //Folder ม.2/2
      } 
      if(stuClass == 'มัธยมศึกษาปีที่ 2' && stuRoom == 3){
        var destinationFolder = m23   //Folder ม.2/3
      }
      if(stuClass == 'มัธยมศึกษาปีที่ 2' && stuRoom == 4){
        var destinationFolder = m24   //Folder ม.2/4
      }
      //ม.3
      if(stuClass == 'มัธยมศึกษาปีที่ 3' && stuRoom == 1){
        var destinationFolder = m31   //Folder ม.3/1
      } 
      if(stuClass == 'มัธยมศึกษาปีที่ 3' && stuRoom == 2){
        var destinationFolder = m32   //Folder ม.3/2
      } 
      if(stuClass == 'มัธยมศึกษาปีที่ 3' && stuRoom == 3){
        var destinationFolder = m33   //Folder ม.3/3
      }
      if(stuClass == 'มัธยมศึกษาปีที่ 3' && stuRoom == 4){
        var destinationFolder = m34   //Folder ม.3/4
      }
      //ม.4
      if(stuClass == 'มัธยมศึกษาปีที่ 4' && stuRoom == 1){
        var destinationFolder = m41   //Folder ม.4/1
      } 
      if(stuClass == 'มัธยมศึกษาปีที่ 4' && stuRoom == 2){
        var destinationFolder = m42   //Folder ม.4/2
      } 
      if(stuClass == 'มัธยมศึกษาปีที่ 4' && stuRoom == 3){
        var destinationFolder = m43   //Folder ม.4/3
      }
      //ม.5
      if(stuClass == 'มัธยมศึกษาปีที่ 5' && stuRoom == 1){
        var destinationFolder = m51   //Folder ม.5/1
      } 
      if(stuClass == 'มัธยมศึกษาปีที่ 5' && stuRoom == 2){
        var destinationFolder = m52   //Folder ม.5/2
      } 
      if(stuClass == 'มัธยมศึกษาปีที่ 5' && stuRoom == 3){
        var destinationFolder = m53   //Folder ม.5/3
      }
      //ม.6
      if(stuClass == 'มัธยมศึกษาปีที่ 6' && stuRoom == 1){
        var destinationFolder = m61   //Folder ม.6/1
      } 
      if(stuClass == 'มัธยมศึกษาปีที่ 6' && stuRoom == 2){
        var destinationFolder = m62   //Folder ม.6/2
      } 
      if(stuClass == 'มัธยมศึกษาปีที่ 6' && stuRoom == 3){
        var destinationFolder = m63   //Folder ม.6/3
      }
      if(stuClass == 'มัธยมศึกษาปีที่ 6' && stuRoom == 4){
        var destinationFolder = m64   //Folder ม.6/4
      }

      var file_name = 'บันทึกการเยี่ยมบ้าน '+stuClass+'/'+stuRoom+' '+stuOrderID + fName +' '+lName ;
      var copyFile = DriveApp.getFileById(GDOC_TEMPLATE_ID).makeCopy(file_name)   
      DriveApp.getFolderById(destinationFolder).addFile(copyFile);
      var fileId = copyFile.getId()
      var slideCopy = SlidesApp.openById(fileId)
            
        slideCopy.replaceAllText('{{name}}',nameTitle+fName+' '+lName)
        //console.log(nameTitle+fName+' '+lName)
        slideCopy.replaceAllText('{{nickname}}',nName)
        //console.log(nName)
        slideCopy.replaceAllText('{{class}}',stuClass+'/'+stuRoom)
        slideCopy.replaceAllText('{{ID}}',stuOrderID)
        slideCopy.replaceAllText('{{age}}',stuAge)
        slideCopy.replaceAllText('{{nationID}}',stuCitizenID)
        slideCopy.replaceAllText('{{sonOrder}}',orderOfKids)
        slideCopy.replaceAllText('{{totalBro}}',quantityOfKids)

        slideCopy.replaceAllText('{{address}}',houseID)
        slideCopy.replaceAllText('{{villID}}',villegeID)
        slideCopy.replaceAllText('{{alley}}',subRoad)
        slideCopy.replaceAllText('{{street}}',mainRoad)
        slideCopy.replaceAllText('{{canton}}',subDistrict)
        slideCopy.replaceAllText('{{district}}',district)
        slideCopy.replaceAllText('{{province}}',province)
        slideCopy.replaceAllText('{{zipcode}}',zipcode)
        slideCopy.replaceAllText('{{phone}}',stuPhone)
        slideCopy.replaceAllText('{{district}}',district)

        if(hasFather == 'ไม่มี/ไม่ทราบข้อมูล'&& i == lastRow) {
          fatherFName = '-'
          fOccupation = '-'
          fIncome = '-'
          slideCopy.replaceAllText('{{fatherName}}',fatherFName)
          slideCopy.replaceAllText('{{fatherOccupation}}',fOccupation)
          slideCopy.replaceAllText('{{fIncome}}',fIncome)
        }else {
          slideCopy.replaceAllText('{{fatherName}}',fnameTitle+fatherFName+' '+fLName)
          slideCopy.replaceAllText('{{fatherOccupation}}',fOccupation)
          slideCopy.replaceAllText('{{fIncome}}',fIncome)
        }
        
        if (hasMother == 'ไม่มี/ไม่ทราบข้อมูล'&& i == lastRow){
          mFName = '-'
          mOccupation = '-'
          mIncome = '-'
          slideCopy.replaceAllText('{{motherName}}',mFName)
          slideCopy.replaceAllText('{{motherOccupation}}',mOccupation)
          slideCopy.replaceAllText('{{mIncome}}',mIncome)
        }else {
          slideCopy.replaceAllText('{{motherName}}',mNameTitle+mFName+' '+mLName)
          slideCopy.replaceAllText('{{motherOccupation}}',mOccupation)
          slideCopy.replaceAllText('{{mIncome}}',mIncome)
        }
        
        
        if(parent=="บิดา" && i == lastRow){
          pNameTitle = fnameTitle
          pFName = fatherFName
          pLName = fLName
          console.log('ชื่อผู้ปกครอง '+i+' '+pNameTitle+pFName+' '+pLName)
        }
        if(parent=="มารดา"&& i == lastRow){ 
            pNameTitle = mNameTitle
            pFName = mFName
            pLName = mLName          
        }
        
        slideCopy.replaceAllText('{{parentName}}',pNameTitle+pFName+' '+pLName)
        
        slideCopy.replaceAllText('{{relative}}',relationship2Student)
        slideCopy.replaceAllText('{{parentOccupation}}',pOccupation)
        slideCopy.replaceAllText('{{pIncome}}',pIncome)
        slideCopy.replaceAllText('{{pPhone}}',pPhone)
        slideCopy.replaceAllText('{{parentNationID}}',pCitizenID)
        slideCopy.replaceAllText('{{education}}',pEducation)
        slideCopy.replaceAllText('{{totalFamily}}',familyQuantity)

        slideCopy.replaceAllText('{{status}}',familyStatus)
        slideCopy.replaceAllText('{{liveWith}}',stuResident)
        slideCopy.replaceAllText('{{familyIncome}}',familyIncome)
        slideCopy.replaceAllText('{{familyRelative}}',familyRelationships)
        slideCopy.replaceAllText('{{bestRelative}}',goodRelationshipto)
        slideCopy.replaceAllText('{{care}}',studyMethod)
        
        if(studySupport=="ไม่มี"){
          studySupportMethod = "-"
          slideCopy.replaceAllText('{{support}}',studySupportMethod)
        } else slideCopy.replaceAllText('{{support}}',studySupportMethod)

        slideCopy.replaceAllText('{{house}}',residentType)
        slideCopy.replaceAllText('{{ownership}}',houseOwner)
        slideCopy.replaceAllText('{{houseStatus}}',strength)
        slideCopy.replaceAllText('{{area}}',houseEnvi)
        slideCopy.replaceAllText('{{environment}}',communityEnvi)

        if(hasHomeJob == "ไม่ช่วยงานบ้าน"){
          whatHomeJob = "-"
          slideCopy.replaceAllText('{{activity}}',whatHomeJob)
        } else slideCopy.replaceAllText('{{activity}}',whatHomeJob)

        if(hasPartTime == "ไม่มี"){
          whatPartTime = "-"
          incomePartTime = "-"
          slideCopy.replaceAllText('{{parttime}}',whatPartTime)
          slideCopy.replaceAllText('{{parttimeIncome}}',incomePartTime)
        } else {
          slideCopy.replaceAllText('{{parttime}}',whatPartTime)
          slideCopy.replaceAllText('{{parttimeIncome}}',incomePartTime)
        }

        slideCopy.replaceAllText('{{goodBehavior}}',goodBehavier)
        slideCopy.replaceAllText('{{noGoodBehavior}}',noneGoodBehavier)

        if(hasStudyPlan == "ไม่มีการวางแผน"){
          whatStudyPlan = "-"
          slideCopy.replaceAllText('{{plan}}',whatStudyPlan)
        } else slideCopy.replaceAllText('{{plan}}',whatStudyPlan)

        slideCopy.replaceAllText('{{pComment}}',parentOpinion)
        slideCopy.replaceAllText('{{tComment}}',teacherOpinion)
        slideCopy.replaceAllText('{{teacher1}}',t1NameTitle+t1Name)
        slideCopy.replaceAllText('{{teacher2}}',t2NameTitle+t2Name)

        var slides = slideCopy.getSlides()
        var slide2 = slides[1]
        var IMAGE_URL_PIC1 =  photo1.toString().replace("https://drive.google.com/open?", "https://doc.google.com/uc?export=view&")
        slide2.insertImage(IMAGE_URL_PIC1, 30, 135, 265, 120).sendToBack
        var IMAGE_URL_PIC2 =  photo2.toString().replace("https://drive.google.com/open?", "https://doc.google.com/uc?export=view&")
        slide2.insertImage(IMAGE_URL_PIC2, 285, 135, 265, 120).sendToBack   
        var IMAGE_URL_PIC3 =  photo3.toString().replace("https://drive.google.com/open?", "https://doc.google.com/uc?export=view&")
        slide2.insertImage(IMAGE_URL_PIC3, 30, 320, 265, 120).sendToBack   
        var IMAGE_URL_PIC4 =  photo4.toString().replace("https://drive.google.com/open?", "https://doc.google.com/uc?export=view&")
        slide2.insertImage(IMAGE_URL_PIC4, 285, 320, 265, 120).sendToBack        
        //var image = slideCopy.getPageElementById('ge7f69ff345_0_0').asImage()
        //image.replace(photo1) //***แทรกรูปภาพที่มาจากฟอร์ม แทนที่ภาพเดิมที่อยู่ในสไลด์
        slideCopy.saveAndClose(); //***บันทึกและปิดไฟล์สไลด์
      
      copyFile.setTrashed(true)
      var pdf_file = DriveApp.createFile(copyFile.getAs("application/pdf"))
      var save_pdf_folder = DriveApp.getFolderById(destinationFolder);
      save_pdf_folder.addFile(pdf_file);
      //sheet.getRange(lastRow, 42).setValue(processSuccess) //พิมพ์ข้อความ Process success ในคอลัมภ์สุดท้าย
      //SpreadsheetApp.flush()
      subject = 'เอกสารการเยี่ยมบ้านของ '+nameTitle+fName+' '+lName
      message = 'เรียนครูประจำชั้น'+stuClass+'/'+stuRoom+'\n\tเนื่องด้วยงานระบบดูแลช่วยเหลือนักเรียนโรงเรียนฤทธิณรงค์รอน ได้จัดทำบันทึกการเยี่ยมบ้านของ'+nameTitle+fName+' '+lName+' ชั้น'+stuClass+'/'+stuRoom+' เสร็จแล้ว และได้แนบแบบบันทึกการเยี่ยมบ้านมาในอีเมลฉบับนี้แล้ว ทีมงานขอขอบคุณที่ท่านให้ความร่วมมือในการจัดทำระบบดูแลช่วยเหลือนักเรียนของโรงเรียนฤทธิณรงค์รอนให้สำเร็จลุล่วงไปด้วยดี ทางทีมงานจะได้นำข้อมูลนี้ไปประมวลผลเพื่อให้การช่วยเหลือนักเรียนต่อไป\n\nงานรับบดูแลช่วยเหลือ\nกิจกรรมพัฒนาผู้เรียน'
      //MailApp.sendEmail(t1Email,subject,message,{attachments:[pdf_file],cc:t2Email})
      sheet.getRange(lastRow, 79).setValue(processSuccess) //พิมพ์ข้อความ Process success ในคอลัมภ์สุดท้าย
      SpreadsheetApp.flush()
    }
  

  }

}
