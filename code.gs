//สคริปต์นี้ พัฒนามาจากสคริปต์ประเมิน SDQ นำมาปรับปปรุงให้สามารถใชบันทึกข้อมูลการเยี่ยมบ้านออนไลน์
//พัฒนาต่อยอดให้ส่งอีเมลไปหาครูประจำชั้นเมื่อส่งฟอร์มโดยนายยุทธ ขวัญเมืองแก้ว โรงเรียนฤทธิณรงค์รอน สพม.1
//สามารถส่งอีเมลไปยังครูที่ปรึกษาทันทีที่นักเรียนกรอกฟอร์มเสร็จ
//ต้องสร้างโฟลเดอร์รอเพื่อเก็บไฟล์สำเนา
//ปรับปรุงให้นำไฟล์สำเนาไปเก็บในโฟลเดอร์ที่กำหนด


const GDOC_TEMPLATE_ID =  '1F0VOySqTzZ_cZz4THnT4y8ez5wxxg7eJ644bNzXARss'  
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
var subject = ''
var message = ''

function visitHomeRecord() { 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet()
  var start_row = 2
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var range = sheet.getRange(start_row, 1, lastRow-1 , lastColumn);
  range.activate()
  var destinationFolder = ''

    var checkProcess = sheet.getRange(lastRow,lastColumn).getValue()  //ดึงข้อมูลในคอลัมภ์สุดท้าย
    if(checkProcess!==processSuccess) {       //ถ้าข้อมูลในคอลัมสุดท้ายไม่เท่ากับ Process success จะดึงข้อมูลที่เหลือขึ้นมาทำการประมวลผล             

    //================ ข้อมูลส่วนตัวนักเรียน
    var nameTitle = sheet.getRange(lastRow,2).getValue()    //คำนำหน้า
    var fName = sheet.getRange(lastRow,3).getValue()     //ชื่อ
    var lName = sheet.getRange(lastRow,4).getValue()        //นามสกุล
    var nName = sheet.getRange(lastRow,5).getValue()         //ชื่อเล่น
    var stuClass = sheet.getRange(lastRow,6).getValue()         //ระดับชั้น
    var stuRoom = sheet.getRange(lastRow,7).getValue()         //ห้อง
    var stuOrderID = sheet.getRange(lastRow,8).getValue()         //เลขที่
    var stuAge = sheet.getRange(lastRow,9).getValue()         //อายุ
    var stuCitizenID = sheet.getRange(lastRow,10).getValue()         //เลขประจำตัวประชาชน
    var orderOfKids = sheet.getRange(lastRow,11).getValue()        //เป็นบุตรคนที่
    var quantityOfKids = sheet.getRange(lastRow,12).getValue()        //จำนวนพี่น้องทั้งหมด
    //================ ข้อมูลที่อยู่นักเรียน
    var houseID = sheet.getRange(lastRow,13).getValue()        //บ้านเลขที่
    var villegeID = sheet.getRange(lastRow,14).getValue()        //หมู่
    var subRoad = sheet.getRange(lastRow,15).getValue()        //ซอย
    var mainRoad = sheet.getRange(lastRow,16).getValue()        //ถนน
    var subDistrict = sheet.getRange(lastRow,17).getValue()        //แขวง/ตำบล
    var district = sheet.getRange(lastRow,18).getValue()        //เขต/อำเภอ
    var province = sheet.getRange(lastRow,19).getValue()        //จังหวัด
    var zipcode = sheet.getRange(lastRow,20).getValue()        //รหัสไปรษณีย์
    var stuPhone = sheet.getRange(lastRow,21).getValue()        //โทรศัพท์

    //================ ข้อมูลส่วนตัวบิดา
    var hasFather = sheet.getRange(lastRow,22).getValue() //มีบิดาหรือไม่
    var fnameTitle = sheet.getRange(lastRow,23).getValue() //คำนำหน้า
    var fatherFName = sheet.getRange(lastRow,24).getValue()     //ชื่อ
    var fLName = sheet.getRange(lastRow,25).getValue()        //นามสกุล
    var fOccupation = sheet.getRange(lastRow,26).getValue()         //อาชีพ
    var fIncome = sheet.getRange(lastRow,27).getValue()         //รายได้

    //================ ข้อมูลส่วนตัวมารดา
    var hasMother = sheet.getRange(lastRow,28).getValue() //มีมารดาหรือไม่
    var mNameTitle = sheet.getRange(lastRow,29).getValue() //คำนำหน้า
    var mFName = sheet.getRange(lastRow,30).getValue()     //ชื่อ
    var mLName = sheet.getRange(lastRow,31).getValue()        //นามสกุล
    var mOccupation = sheet.getRange(lastRow,32).getValue()         //อาชีพ
    var mIncome = sheet.getRange(lastRow,33).getValue()         //รายได้

    //================ ข้อมูลส่วนตัวผู้ปกครอง
    var parent = sheet.getRange(lastRow,34).getValue()  //บุคคลที่เป็นผู้ปกครอง
    var pNameTitle = sheet.getRange(lastRow,35).getValue() //คำนำหน้า
    var pFName = sheet.getRange(lastRow,36).getValue()     //ชื่อ
    var pLName = sheet.getRange(lastRow,37).getValue()        //นามสกุล
    var relationship2Student = sheet.getRange(lastRow,38).getValue()        //ความสัมพันธ์กับนักเรียน
    var pOccupation = sheet.getRange(lastRow,39).getValue()         //อาชีพ
    var pIncome = sheet.getRange(lastRow,40).getValue()         //รายได้
    var pPhone = sheet.getRange(lastRow,41).getValue()         //โทรศัพท์
    var pCitizenID = sheet.getRange(lastRow,42).getValue()         //เลขประจำตัวประชาชน
    var pEducation = sheet.getRange(lastRow,43).getValue()         //ระดับการศึกษา
    var familyQuantity = sheet.getRange(lastRow,44).getValue()         //สมาชิกในครัวเรือน

    //================ ข้อมูลเฉพาะด้านครอบครัว
    var familyStatus = sheet.getRange(lastRow,45).getValue()  //สถานภาพของบิดา-มารดา
    var stuResident = sheet.getRange(lastRow,46).getValue() //ปัจจุบันอาศัยอยู่กับ
    var familyIncome = sheet.getRange(lastRow,47).getValue()  //ฐานะทางเศรษฐกิจของครอบครัว
    var familyRelationships = sheet.getRange(lastRow,48).getValue()   //ความสัมพันธ์ของสมาชิกในครอบครัว
    var goodRelationshipto = sheet.getRange(lastRow,49).getValue()    //บุคคลในครอบครัวที่นักเรียนสนิทสนมมากที่สุด
    var studyMethod = sheet.getRange(lastRow,50).getValue()           //วิธีการที่ผู้ปกครองเลี้ยงดูนักเรียน
    var studySupport = sheet.getRange(lastRow,51).getValue()          //การสนับสนุนด้านการเรียน
    var studySupportMethod = sheet.getRange(lastRow,52).getValue()     //ผู้ปกครองสนับสนุนพิเศษด้านการศึกษาด้วยวิธีใด

    //================ ข้อมูลเฉพาะด้านที่อยู่อาศัย
    var residentType = sheet.getRange(lastRow,53).getValue()          //ลักษณะที่อยู่อาศัย
    var houseOwner = sheet.getRange(lastRow,54).getValue()            //ความเป็นเจ้าของบ้านที่อาศัย
    var strength  = sheet.getRange(lastRow,55).getValue()             //สภาพบ้าน
    var houseEnvi = sheet.getRange(lastRow,56).getValue()              //สภาพแวดล้อมของบ้าน
    var communityEnvi = sheet.getRange(lastRow,57).getValue()          //สภาพรอบบ้านหรือในชุมชน

    //================ ข้อมูลกิจกรรมที่บ้าน
    var hasHomeJob = sheet.getRange(lastRow,58).getValue()            //การช่วยพ่อแม่ทำงานบ้าน
    var whatHomeJob = sheet.getRange(lastRow,59).getValue()           //นักเรียนช่วยทำงานบ้านอะไรบ้าง
    var hasPartTime = sheet.getRange(lastRow,60).getValue()           //นักเรียนมีรายได้พิเศษ
    var whatPartTime = sheet.getRange(lastRow,61).getValue()           //นักเรียนมีรายได้พิเศษจากงานใด
    var incomePartTime = sheet.getRange(lastRow,62).getValue()           //รายได้จากงานพิเศษคิดเป็น บาท/วัน
    var goodBehavier = sheet.getRange(lastRow,63).getValue()          //นักเรียนมีพฤติกรรมใดที่น่าพึงพอใจบ้าง
    var noneGoodBehavier = sheet.getRange(lastRow,64).getValue()      //นักเรียนมีพฤติกรรมใดที่ไม่น่าพึงพอใจบ้าง
    var hasStudyPlan = sheet.getRange(lastRow,65).getValue()          //มีการวางแผนด้านการเรียนหรือไม่
    var whatStudyPlan = sheet.getRange(lastRow,66).getValue()         //มีการวางแผนด้านการเรียนในอนาคตอย่างไร
    var parentOpinion = sheet.getRange(lastRow,67).getValue()         //ความคิดเห็นของผู้ปกครอง
    var teacherOpinion = sheet.getRange(lastRow,68).getValue()        //ความคิดเห็นของผู้ครูที่ปรึกษา

    //================ ข้อมูลสถานการณ์โควิด 19
    var risk = sheet.getRange(lastRow,69).getValue()                  //สถานการณ์โควิด 19
    var device = sheet.getRange(lastRow,70).getValue()                //อุปกรณ์การเรียนออนไลน์
    var network = sheet.getRange(lastRow,71).getValue()               //เครือข่ายที่ใช้ในการเรียนออนไลน์

    //================ ข้อมูลครูที่ปรึกษา
    var t1NameTitle = sheet.getRange(lastRow,72).getValue()           
    var t1Name = sheet.getRange(lastRow,73).getValue()
    var t1Email = sheet.getRange(lastRow,74).getValue()
    var t2NameTitle = sheet.getRange(lastRow,75).getValue()
    var t2Name = sheet.getRange(lastRow,76).getValue()
    var t2Email = sheet.getRange(lastRow,77).getValue()

    //ข้อมูลภาพถ่าย
    var photo1 = sheet.getRange(lastRow,78).getValue()       //แนบรูปภาพที่ได้จากการสำรวจ
    var photo2 = sheet.getRange(lastRow,79).getValue()       //แนบรูปภาพที่ได้จากการสำรวจ
    var photo3 = sheet.getRange(lastRow,80).getValue()       //แนบรูปภาพที่ได้จากการสำรวจ
    var photo4 = sheet.getRange(lastRow,81).getValue()       //แนบรูปภาพที่ได้จากการสำรวจ

      
      //แยกข้อมูลแต่ละห้องไปใส่ในโฟลเดอร์ที่กำหนด 
      //ม.1
      if(stuClass == 'ม 1' && stuRoom == 1){
        destinationFolder = m11   //Folder ม.1/1
      } 
      if(stuClass == 'ม 1' && stuRoom == 2){
        destinationFolder = m12   //Folder ม.1/2
      } 
      if(stuClass == 'ม 1' && stuRoom == 3){
        destinationFolder = m13   //Folder ม.1/3
      }
      if(stuClass == 'ม 1' && stuRoom == 4){
        destinationFolder = m14   //Folder ม.1/4
      }
      //ม.2
      if(stuClass == 'ม 2' && stuRoom == 1){
        destinationFolder = m21   //Folder ม.2/1
      } 
      if(stuClass == 'ม 2' && stuRoom == 2){
        destinationFolder = m22   //Folder ม.2/2
      } 
      if(stuClass == 'ม 2' && stuRoom == 3){
        destinationFolder = m23   //Folder ม.2/3
      }
      if(stuClass == 'ม 2' && stuRoom == 4){
        destinationFolder = m24   //Folder ม.2/4
      }
      //ม.3
      if(stuClass == 'ม 3' && stuRoom == 1){
        destinationFolder = m31   //Folder ม.3/1
      } 
      if(stuClass == 'ม 3' && stuRoom == 2){
        destinationFolder = m32   //Folder ม.3/2
      } 
      if(stuClass == 'ม 3' && stuRoom == 3){
        destinationFolder = m33   //Folder ม.3/3
      }
      if(stuClass == 'ม 3' && stuRoom == 4){
        destinationFolder = m34   //Folder ม.3/4
      }
      //ม.4
      if(stuClass == 'ม 4' && stuRoom == 1){
        destinationFolder = m41   //Folder ม.4/1
      } 
      if(stuClass == 'ม 4' && stuRoom == 2){
        destinationFolder = m42   //Folder ม.4/2
      } 
      if(stuClass == 'ม 4' && stuRoom == 3){
        destinationFolder = m43   //Folder ม.4/3
      }
      //ม.5
      if(stuClass == 'ม 5' && stuRoom == 1){
        destinationFolder = m51   //Folder ม.5/1
      } 
      if(stuClass == 'ม 5' && stuRoom == 2){
        destinationFolder = m52   //Folder ม.5/2
      } 
      if(stuClass == 'ม 5' && stuRoom == 3){
        destinationFolder = m53   //Folder ม.5/3
      }
      //ม.6
      if(stuClass == 'ม 6' && stuRoom == 1){
        destinationFolder = m61   //Folder ม.6/1
      } 
      if(stuClass == 'ม 6' && stuRoom == 2){
        destinationFolder = m62   //Folder ม.6/2
      } 
      if(stuClass == 'ม 6' && stuRoom == 3){
        destinationFolder = m63   //Folder ม.6/3
      }
      if(stuClass == 'ม 6' && stuRoom == 4){
        destinationFolder = m64   //Folder ม.6/4
      }

      var file_name = 'บันทึกการเยี่ยมบ้าน '+stuClass+'/'+stuRoom+' '+stuOrderID + fName +' '+lName ;
      var copyFile = DriveApp.getFileById(GDOC_TEMPLATE_ID).makeCopy(file_name) 
      DriveApp.getFolderById(destinationFolder).addFile(copyFile);
      var fileId = copyFile.getId()
      var slideCopy = SlidesApp.openById(fileId)
            
        slideCopy.replaceAllText('{{name}}',nameTitle+fName+' '+lName)
        slideCopy.replaceAllText('{{nickname}}',nName)
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

        if(hasFather == 'ไม่มี/ไม่ทราบข้อมูล') {
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
        
        if (hasMother == 'ไม่มี/ไม่ทราบข้อมูล'){
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
        
        
        if(parent=="บิดา"){
          pNameTitle = fnameTitle
          pFName = fatherFName
          pLName = fLName
          pOccupation = fOccupation
          pIncome = fIncome
          relationship2Student = parent          
        }
        if(parent=="มารดา"){ 
          pNameTitle = mNameTitle
          pFName = mFName
          pLName = mLName             
          pOccupation = fOccupation
          pIncome = fIncome  
          relationship2Student = parent      
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
        slideCopy.replaceAllText('{{risk}}',risk)
        slideCopy.replaceAllText('{{device}}',device)
        slideCopy.replaceAllText('{{network}}',network)
        slideCopy.replaceAllText('{{teacher1}}',t1NameTitle+t1Name)
        slideCopy.replaceAllText('{{teacher2}}',t2NameTitle+t2Name)

        var slides = slideCopy.getSlides()
        var slide2 = slides[1]
        var IMAGE_URL_PIC1 =  photo1.toString().replace("https://drive.google.com/open?", "https://doc.google.com/uc?export=view&")
        slide2.insertImage(IMAGE_URL_PIC1, 35, 205, 265, 120).sendToBack
        var IMAGE_URL_PIC2 =  photo2.toString().replace("https://drive.google.com/open?", "https://doc.google.com/uc?export=view&")
        slide2.insertImage(IMAGE_URL_PIC2, 290, 205, 265, 120).sendToBack   
        var IMAGE_URL_PIC3 =  photo3.toString().replace("https://drive.google.com/open?", "https://doc.google.com/uc?export=view&")
        slide2.insertImage(IMAGE_URL_PIC3, 35, 390, 265, 120).sendToBack   
        var IMAGE_URL_PIC4 =  photo4.toString().replace("https://drive.google.com/open?", "https://doc.google.com/uc?export=view&")
        slide2.insertImage(IMAGE_URL_PIC4, 290, 390, 265, 120).sendToBack        
        slideCopy.saveAndClose(); //***บันทึกและปิดไฟล์สไลด์
      
      copyFile.setTrashed(true)
      var pdf_file = DriveApp.createFile(copyFile.getAs("application/pdf"))
      var save_pdf_folder = DriveApp.getFolderById(destinationFolder);
      save_pdf_folder.addFile(pdf_file);
      subject = 'เอกสารการเยี่ยมบ้านของ '+nameTitle+fName+' '+lName
      message = 'เรียนครูประจำชั้น'+stuClass+'/'+stuRoom+'\n\tเนื่องด้วยงานระบบดูแลช่วยเหลือนักเรียนโรงเรียนฤทธิณรงค์รอน ได้จัดทำบันทึกการเยี่ยมบ้านของ'+nameTitle+fName+' '+lName+' ชั้น'+stuClass+'/'+stuRoom+' เสร็จแล้ว และได้แนบแบบบันทึกการเยี่ยมบ้านมาในอีเมลฉบับนี้แล้ว ทีมงานขอขอบคุณที่ท่านให้ความร่วมมือในการจัดทำระบบดูแลช่วยเหลือนักเรียนของโรงเรียนฤทธิณรงค์รอนให้สำเร็จลุล่วงไปด้วยดี ทางทีมงานจะได้นำข้อมูลนี้ไปประมวลผลเพื่อให้การช่วยเหลือนักเรียนต่อไป\n\nงานรับบดูแลช่วยเหลือ\nกิจกรรมพัฒนาผู้เรียน'
      MailApp.sendEmail(t1Email,subject,message,{attachments:[pdf_file],cc:t2Email})
      sheet.getRange(lastRow, lastColumn+1).setValue(processSuccess) //พิมพ์ข้อความ Process success ในคอลัมภ์สุดท้าย
      SpreadsheetApp.flush()
    }
}
