/**
 * Google Apps Script สำหรับบันทึกข้อมูลทีมลง Google Sheets
 * 
 * วิธีการติดตั้ง:
 * 1. เปิด Google Sheets ที่ต้องการบันทึกข้อมูล
 * 2. ไปที่ Extensions > Apps Script
 * 3. ลบโค้ดเดิมทั้งหมดแล้ววางโค้ดนี้แทน
 * 4. บันทึกและตั้งชื่อโปรเจค
 * 5. สร้างชีท "การตั้งค่า" และใส่รหัสผ่านที่ A1 (ค่าเริ่มต้น: abcd)
 * 6. Deploy > New deployment
 * 7. เลือก Type: Web app
 * 8. Execute as: Me
 * 9. Who has access: Anyone
 * 10. Deploy และคัดลอก URL ที่ได้
 * 11. นำ URL ไปใส่ในตัวแปร gasUrl ในไฟล์ midhigh.html
 */

/**
 * ฟังก์ชันหลักสำหรับรับ HTTP GET request
 */
function doGet(e) {
  try {
    // ตรวจสอบว่ามี parameters หรือไม่
    if (!e || !e.parameter) {
      return createJSONPResponse(e.parameter.callback, {
        success: false,
        message: 'No parameters provided'
      });
    }
    
    const params = e.parameter;
    const action = params.action;
    const callback = params.callback;
    
    // ตรวจสอบ action
    if (action === 'addTeam') {
      const result = addTeamToSheet(params);
      return createJSONPResponse(callback, result);
    } else if (action === 'getAllTeams') {
      const result = getAllTeams();
      return createJSONPResponse(callback, result);
    } else if (action === 'updateScore') {
      const result = updateScore(params);
      return createJSONPResponse(callback, result);
    } else if (action === 'getPassword') {
      const result = getPassword();
      return createJSONPResponse(callback, result);
    } else {
      return createJSONPResponse(callback, {
        success: false,
        message: 'Invalid action'
      });
    }
    
  } catch (error) {
    Logger.log('Error in doGet: ' + error.toString());
    return createJSONPResponse(e.parameter.callback, {
      success: false,
      message: 'Server error: ' + error.toString()
    });
  }
}

/**
 * ฟังก์ชันสำหรับสร้าง JSONP response
 */
function createJSONPResponse(callback, data) {
  const jsonString = JSON.stringify(data);
  const output = callback + '(' + jsonString + ')';
  return ContentService.createTextOutput(output)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

/**
 * ฟังก์ชันสำหรับเพิ่มข้อมูลทีมลง Google Sheets
 */
function addTeamToSheet(params) {
  try {
    // เปิด Spreadsheet (ใช้ active spreadsheet หรือระบุ ID)
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // หาชีทที่ชื่อ "ข้อมูล" หรือสร้างใหม่ถ้ายังไม่มี
    let sheet = ss.getSheetByName('ข้อมูล');
    if (!sheet) {
      sheet = ss.insertSheet('ข้อมูล');
      // เพิ่มหัวตาราง
      const headers = [
        'school_id',
        'โรงเรียน',
        'เขต',
        'ระดับ',
        'รายการแข่ง',
        'สนาม',
        '# score1',
        '# retry1',
        'time1',
        '# score2',
        '# retry2',
        'time2',
        'timestamp'
      ];
      sheet.appendRow(headers);
      
      // จัดรูปแบบหัวตาราง
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#1976D2');
      headerRange.setFontColor('#FFFFFF');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
    }
    
    // ตรวจสอบว่ามีข้อมูลที่จำเป็นครบหรือไม่
    if (!params.school_id || !params.school_name) {
      return {
        success: false,
        message: 'Missing required fields'
      };
    }
    
    // เตรียมข้อมูลที่จะบันทึก
    const rowData = [
      params.school_id || '',
      params.school_name || '',
      params.zone || '',
      params.level || '',
      params.competition || '',
      params.field || '',
      0,  // score1
      0,  // retry1
      '0:00:00',  // time1
      0,  // score2
      0,  // retry2
      '0:00:00',  // time2
      params.timestamp || new Date().toISOString()
    ];
    
    // ตรวจสอบว่ามีข้อมูลซ้ำหรือไม่ (ตรวจจาก school_id)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const schoolIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < schoolIds.length; i++) {
        if (schoolIds[i][0] === params.school_id) {
          return {
            success: false,
            message: 'รหัสโรงเรียนนี้มีอยู่แล้วในระบบ'
          };
        }
      }
    }
    
    // เพิ่มข้อมูลลงในชีท
    sheet.appendRow(rowData);
    
    // จัดรูปแบบแถวที่เพิ่ม
    const newRow = sheet.getLastRow();
    const newRange = sheet.getRange(newRow, 1, 1, rowData.length);
    
    // สลับสีแถว
    if (newRow % 2 === 0) {
      newRange.setBackground('#F5F5F5');
    }
    
    // จัดให้อยู่กึ่งกลางสำหรับคอลัมน์ตัวเลข
    sheet.getRange(newRow, 7, 1, 6).setHorizontalAlignment('center');
    
    return {
      success: true,
      message: 'บันทึกข้อมูลสำเร็จ',
      data: {
        row: newRow,
        school_id: params.school_id,
        school_name: params.school_name
      }
    };
    
  } catch (error) {
    Logger.log('Error in addTeamToSheet: ' + error.toString());
    return {
      success: false,
      message: 'เกิดข้อผิดพลาด: ' + error.toString()
    };
  }
}

/**
 * ฟังก์ชันสำหรับสร้างโครงสร้างชีททั้งหมดครั้งแรก
 * รันฟังก์ชันนี้ใน Apps Script Editor โดยตรง (ไม่ต้องผ่าน URL)
 * 
 * วิธีใช้งาน:
 * 1. เปิด Apps Script Editor
 * 2. เลือกฟังก์ชัน initializeSheets จากเมนู dropdown
 * 3. คลิก Run (▶️)
 * 4. อนุญาตการเข้าถึงครั้งแรก
 */
function initializeSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // ===== 1. สร้างชีท "การตั้งค่า" =====
    let settingsSheet = ss.getSheetByName('การตั้งค่า');
    if (!settingsSheet) {
      settingsSheet = ss.insertSheet('การตั้งค่า');
      settingsSheet.getRange('A1').setValue('abcd');
      settingsSheet.getRange('B1').setValue('← รหัสผ่านสำหรับยืนยันตัวตน (แก้ไขได้)');
      settingsSheet.getRange('A2').setValue('');
      settingsSheet.getRange('B2').setValue('← รายการแข่ง default (ว่าง = แสดงทั้งหมด)');
      
      // จัดรูปแบบ
      settingsSheet.getRange('A1').setBackground('#FFE599');
      settingsSheet.getRange('A1').setFontWeight('bold');
      settingsSheet.getRange('A1').setFontSize(14);
      settingsSheet.getRange('B1').setFontStyle('italic');
      settingsSheet.getRange('B1').setFontColor('#666666');
      
      settingsSheet.getRange('A2').setBackground('#E3F2FD');
      settingsSheet.getRange('A2').setFontWeight('bold');
      settingsSheet.getRange('A2').setFontSize(14);
      settingsSheet.getRange('B2').setFontStyle('italic');
      settingsSheet.getRange('B2').setFontColor('#666666');
      
      // ปรับขนาดคอลัมน์
      settingsSheet.setColumnWidth(1, 150);
      settingsSheet.setColumnWidth(2, 400);
      
      Logger.log('✓ สร้างชีท "การตั้งค่า" สำเร็จ');
    } else {
      Logger.log('ℹ️ ชีท "การตั้งค่า" มีอยู่แล้ว');
    }
    
    // ===== 2. สร้างชีท "ข้อมูล" สำหรับข้อมูลทีม =====
    let dataSheet = ss.getSheetByName('ข้อมูล');
    if (!dataSheet) {
      dataSheet = ss.insertSheet('ข้อมูล');
      
      // หัวตาราง
      const headers = [
        'school_id',
        'โรงเรียน',
        'เขต',
        'ระดับ',
        'รายการแข่ง',
        'สนาม',
        '# score1',
        '# retry1',
        'time1',
        '# score2',
        '# retry2',
        'time2',
        'timestamp'
      ];
      
      // ใส่หัวตาราง
      const headerRange = dataSheet.getRange(1, 1, 1, headers.length);
      headerRange.setValues([headers]);
      
      // จัดรูปแบบหัวตาราง
      headerRange.setBackground('#1976D2');
      headerRange.setFontColor('#FFFFFF');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
      headerRange.setVerticalAlignment('middle');
      headerRange.setFontSize(11);
      
      // ปรับขนาดคอลัมน์
      dataSheet.setColumnWidth(1, 100);  // school_id
      dataSheet.setColumnWidth(2, 200);  // โรงเรียน
      dataSheet.setColumnWidth(3, 80);   // เขต
      dataSheet.setColumnWidth(4, 80);   // ระดับ
      dataSheet.setColumnWidth(5, 120);  // รายการแข่ง
      dataSheet.setColumnWidth(6, 70);   // สนาม
      dataSheet.setColumnWidth(7, 90);   // score1
      dataSheet.setColumnWidth(8, 90);   // retry1
      dataSheet.setColumnWidth(9, 100);  // time1
      dataSheet.getRange(2, 9, dataSheet.getMaxRows(), 1).setNumberFormat('@'); // time1 เป็น text
      dataSheet.setColumnWidth(10, 90);  // score2
      dataSheet.setColumnWidth(11, 90);  // retry2
      dataSheet.setColumnWidth(12, 100); // time2
      dataSheet.getRange(2, 12, dataSheet.getMaxRows(), 1).setNumberFormat('@'); // time2 เป็น text
      dataSheet.setColumnWidth(13, 150); // timestamp
      
      // ตรึงแถวหัวตาราง
      dataSheet.setFrozenRows(1);
      
      Logger.log('✓ สร้างชีท "ข้อมูล" สำเร็จ');
    } else {
      Logger.log('ℹ️ ชีท "ข้อมูล" มีอยู่แล้ว');
    }
    
    // ===== 3. สร้างชีท "ตัวอย่างข้อมูล" =====
    let sampleSheet = ss.getSheetByName('ตัวอย่างข้อมูล');
    if (!sampleSheet) {
      sampleSheet = ss.insertSheet('ตัวอย่างข้อมูล');
      
      // หัวตาราง
      const headers = [
        'school_id',
        'โรงเรียน',
        'เขต',
        'ระดับ',
        'รายการแข่ง',
        'สนาม',
        '# score1',
        '# retry1',
        'time1',
        '# score2',
        '# retry2',
        'time2',
        'timestamp'
      ];
      
      // ข้อมูลตัวอย่าง
      const sampleData = [
        headers,
        ['s001', 'สมศุภนาการวิทยา', '4', 'ม.ปลาย', 'สพม.นศ68', 'A', 100, 0, '0:43:84', 75, 0, '1:00:00', '2025-01-08T10:00:00.000Z'],
        ['s002', 'สตรีปากพนัง', '3', 'ม.ปลาย', 'สพม.นศ68', 'B', 0, 1, '0:00:00', 0, 4, '0:00:00', '2025-01-08T10:01:00.000Z'],
        ['s003', 'ทุ่งสง', '2', 'ม.ปลาย', 'สพม.นศ68', 'A', 30, 3, '0:00:00', 60, 2, '0:00:00', '2025-01-08T10:02:00.000Z']
      ];
      
      // ใส่ข้อมูล
      sampleSheet.getRange(1, 1, sampleData.length, headers.length).setValues(sampleData);
      
      // จัดรูปแบบหัวตาราง
      const headerRange = sampleSheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#4CAF50');
      headerRange.setFontColor('#FFFFFF');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
      
      // จัดรูปแบบข้อมูล
      const dataRange = sampleSheet.getRange(2, 1, sampleData.length - 1, headers.length);
      dataRange.setHorizontalAlignment('center');
      dataRange.setBorder(true, true, true, true, true, true);
      
      // สลับสีแถว
      for (let i = 2; i <= sampleData.length; i++) {
        if (i % 2 === 0) {
          sampleSheet.getRange(i, 1, 1, headers.length).setBackground('#F5F5F5');
        }
      }
      
      // ปรับขนาดคอลัมน์
      sampleSheet.setColumnWidth(1, 100);
      sampleSheet.setColumnWidth(2, 200);
      sampleSheet.setColumnWidth(3, 80);
      sampleSheet.setColumnWidth(4, 80);
      sampleSheet.setColumnWidth(5, 120);
      sampleSheet.setColumnWidth(6, 70);
      sampleSheet.setColumnWidth(7, 90);
      sampleSheet.setColumnWidth(8, 90);
      sampleSheet.setColumnWidth(9, 100);
      sampleSheet.setColumnWidth(10, 90);
      sampleSheet.setColumnWidth(11, 90);
      sampleSheet.setColumnWidth(12, 100);
      sampleSheet.setColumnWidth(13, 150);
      
      // ตรึงแถวหัวตาราง
      sampleSheet.setFrozenRows(1);
      
      Logger.log('✓ สร้างชีท "ตัวอย่างข้อมูล" สำเร็จ');
    } else {
      Logger.log('ℹ️ ชีท "ตัวอย่างข้อมูล" มีอยู่แล้ว');
    }
    
    // ===== 4. เรียงลำดับชีท =====
    const sheetOrder = ['การตั้งค่า', 'ข้อมูล', 'ตัวอย่างข้อมูล'];
    sheetOrder.forEach((sheetName, index) => {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        ss.setActiveSheet(sheet);
        ss.moveActiveSheet(index + 1);
      }
    });
    
    // ตั้งชีท "ข้อมูล" เป็นชีทที่เปิดอยู่
    ss.setActiveSheet(ss.getSheetByName('ข้อมูล'));
    
    Logger.log('=================================');
    Logger.log('✅ สร้างโครงสร้างชีททั้งหมดสำเร็จ!');
    Logger.log('=================================');
    Logger.log('ชีทที่สร้าง:');
    Logger.log('1. การตั้งค่า - รหัสผ่าน: abcd');
    Logger.log('2. ข้อมูล - พร้อมหัวตารางและรอข้อมูล');
    Logger.log('3. ตัวอย่างข้อมูล - มีข้อมูลตัวอย่าง 3 ทีม');
    Logger.log('=================================');
    
    // แสดง dialog แจ้งเตือน
    SpreadsheetApp.getUi().alert(
      'สำเร็จ! ✅\n\n' +
      'สร้างโครงสร้างชีททั้งหมดเรียบร้อยแล้ว\n\n' +
      '✓ การตั้งค่า - รหัสผ่าน: abcd\n' +
      '✓ ข้อมูล - พร้อมใช้งาน\n' +
      '✓ ตัวอย่างข้อมูล - มีข้อมูลตัวอย่าง\n\n' +
      'ตอนนี้สามารถ Deploy เป็น Web App และใช้งานได้เลย'
    );
    
    return {
      success: true,
      message: 'สร้างโครงสร้างชีททั้งหมดสำเร็จ',
      sheets: ['การตั้งค่า', 'ข้อมูล', 'ตัวอย่างข้อมูล']
    };
    
  } catch (error) {
    Logger.log('❌ เกิดข้อผิดพลาด: ' + error.toString());
    SpreadsheetApp.getUi().alert('เกิดข้อผิดพลาด!\n\n' + error.toString());
    return {
      success: false,
      message: 'เกิดข้อผิดพลาด: ' + error.toString()
    };
  }
}

/**
 * ฟังก์ชันสำหรับดึงรหัสผ่านจาก Google Sheets
 */
function getPassword() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let settingsSheet = ss.getSheetByName('การตั้งค่า');
    
    if (!settingsSheet) {
      // สร้างชีทการตั้งค่าใหม่ถ้ายังไม่มี
      settingsSheet = ss.insertSheet('การตั้งค่า');
      settingsSheet.getRange('A1').setValue('abcd'); // รหัสผ่านเริ่มต้น
      settingsSheet.getRange('A1').setBackground('#FFE599');
      settingsSheet.getRange('A1').setFontWeight('bold');
      settingsSheet.getRange('B1').setValue('← รหัสผ่านสำหรับยืนยันตัวตน (แก้ไขได้)');
      settingsSheet.getRange('A2').setValue(''); // รายการแข่ง default
      settingsSheet.getRange('A2').setBackground('#E3F2FD');
      settingsSheet.getRange('A2').setFontWeight('bold');
      settingsSheet.getRange('B2').setValue('← รายการแข่ง default (ว่าง = แสดงทั้งหมด)');
    }
    
    const password = settingsSheet.getRange('A1').getValue() || 'abcd';
    const defaultCompetition = settingsSheet.getRange('A2').getValue() || '';
    
    return {
      success: true,
      password: password,
      defaultCompetition: defaultCompetition
    };
  } catch (error) {
    Logger.log('Error in getPassword: ' + error.toString());
    return {
      success: false,
      message: 'เกิดข้อผิดพลาด: ' + error.toString()
    };
  }
}

/**
 * ฟังก์ชันสำหรับดึงข้อมูลทีมทั้งหมด
 */
function getAllTeams() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ข้อมูล');
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return {
        success: true,
        data: []
      };
    }
    
    // ใช้ getDisplayValues() เพื่อดึงค่าที่แสดงใน Sheet โดยตรง (รวมถึง text format)
    const data = sheet.getDataRange().getDisplayValues();
    const headers = data[0];
    const teams = [];
    
    // สร้าง mapping สำหรับแปลงชื่อคอลัมน์
    const fieldMapping = {
      '# score1': 'score1',
      '# retry1': 'retry1',
      '# score2': 'score2',
      '# retry2': 'retry2',
      'โรงเรียน': 'school_name',
      'เขต': 'zone',
      'ระดับ': 'level',
      'รายการแข่ง': 'competition',
      'สนาม': 'field'
    };
    
    for (let i = 1; i < data.length; i++) {
      const team = {};
      for (let j = 0; j < headers.length; j++) {
        const header = headers[j];
        // ใช้ชื่อที่แปลงแล้ว หรือชื่อเดิมถ้าไม่มีใน mapping
        const key = fieldMapping[header] || header;
        let value = data[i][j];
        
        // แปลงค่าเป็น string และใช้ตามที่แสดงใน Sheet
        if (key === 'time1' || key === 'time2') {
          // ใช้ค่าที่แสดงใน Sheet โดยตรง (เป็น string แล้ว)
          if (value && value.trim() !== '') {
            team[key] = value.trim();
          } else {
            team[key] = '0:00:00';
          }
        } else if (key === 'score1' || key === 'score2' || key === 'retry1' || key === 'retry2') {
          // แปลงเป็นตัวเลข
          team[key] = parseInt(value) || 0;
        } else {
          team[key] = value;
        }
      }
      teams.push(team);
    }
    
    return {
      success: true,
      data: teams
    };
    
  } catch (error) {
    Logger.log('Error in getAllTeams: ' + error.toString());
    return {
      success: false,
      message: 'เกิดข้อผิดพลาด: ' + error.toString()
    };
  }
}

/**
 * ฟังก์ชันสำหรับอัพเดตคะแนน
 */
function updateScore(params) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ข้อมูล');
    
    if (!sheet) {
      return {
        success: false,
        message: 'ไม่พบชีทข้อมูล'
      };
    }
    
    // หาแถวที่ตรงกับ school_id
    const lastRow = sheet.getLastRow();
    const schoolIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    
    for (let i = 0; i < schoolIds.length; i++) {
      if (schoolIds[i][0] === params.school_id) {
        const rowNum = i + 2; // เริ่มจากแถว 2 (แถว 1 เป็นหัวตาราง)
        
        // อัพเดตคะแนนทั้ง 2 รอบ
        sheet.getRange(rowNum, 7).setValue(params.score1 || 0); // score1
        sheet.getRange(rowNum, 8).setValue(params.retry1 || 0); // retry1
        // บันทึก time1 เป็น text เพื่อป้องกันการแปลงค่าอัตโนมัติ
        const time1Cell = sheet.getRange(rowNum, 9);
        time1Cell.setNumberFormat('@');
        time1Cell.setValue(params.time1 || '0:00:00');
        sheet.getRange(rowNum, 10).setValue(params.score2 || 0); // score2
        sheet.getRange(rowNum, 11).setValue(params.retry2 || 0); // retry2
        // บันทึก time2 เป็น text เพื่อป้องกันการแปลงค่าอัตโนมัติ
        const time2Cell = sheet.getRange(rowNum, 12);
        time2Cell.setNumberFormat('@');
        time2Cell.setValue(params.time2 || '0:00:00');
        // อัพเดต timestamp
        sheet.getRange(rowNum, 13).setValue(new Date().toISOString());
        
        return {
          success: true,
          message: 'อัพเดตคะแนนสำเร็จ'
        };
      }
    }
    
    return {
      success: false,
      message: 'ไม่พบข้อมูลทีม'
    };
    
  } catch (error) {
    Logger.log('Error in updateScore: ' + error.toString());
    return {
      success: false,
      message: 'เกิดข้อผิดพลาด: ' + error.toString()
    };
  }
}
