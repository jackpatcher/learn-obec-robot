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
    } else if (action === 'authenticate') {
      const result = authenticate(params);
      return createJSONPResponse(callback, result);
    } else if (action === 'saveDefaultFilter') {
      const result = saveDefaultFilter(params);
      return createJSONPResponse(callback, result);
    } else if (action === 'getDefaultFilter') {
      const result = getDefaultFilter();
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
    
    // หาชีทที่ชื่อ "ระดับกลาง" หรือสร้างใหม่ถ้ายังไม่มี
    let sheet = ss.getSheetByName('ระดับกลาง');
    if (!sheet) {
      sheet = ss.insertSheet('ระดับกลาง');
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
 * ฟังก์ชันสำหรับสร้างชีทระดับกลาง
 */
function initMidSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    let sheet = ss.getSheetByName('ระดับกลาง');
    if (!sheet) {
      sheet = ss.insertSheet('ระดับกลาง');
      
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
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setValues([headers]);
      
      // จัดรูปแบบหัวตาราง
      headerRange.setBackground('#1976D2');
      headerRange.setFontColor('#FFFFFF');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
      headerRange.setVerticalAlignment('middle');
      headerRange.setFontSize(11);
      
      // ปรับขนาดคอลัมน์
      sheet.setColumnWidth(1, 100);  // school_id
      sheet.setColumnWidth(2, 200);  // โรงเรียน
      sheet.setColumnWidth(3, 80);   // เขต
      sheet.setColumnWidth(4, 80);   // ระดับ
      sheet.setColumnWidth(5, 120);  // รายการแข่ง
      sheet.setColumnWidth(6, 70);   // สนาม
      sheet.setColumnWidth(7, 90);   // score1
      sheet.setColumnWidth(8, 90);   // retry1
      sheet.setColumnWidth(9, 100);  // time1
      sheet.getRange(2, 9, sheet.getMaxRows(), 1).setNumberFormat('@'); // time1 เป็น text
      sheet.setColumnWidth(10, 90);  // score2
      sheet.setColumnWidth(11, 90);  // retry2
      sheet.setColumnWidth(12, 100); // time2
      sheet.getRange(2, 12, sheet.getMaxRows(), 1).setNumberFormat('@'); // time2 เป็น text
      sheet.setColumnWidth(13, 150); // timestamp
      
      // ตรึงแถวหัวตาราง
      sheet.setFrozenRows(1);
      
      Logger.log('✓ สร้างชีท "ระดับกลาง" สำเร็จ');
      return { success: true, message: 'สร้างชีท "ระดับกลาง" สำเร็จ' };
    } else {
      Logger.log('ℹ️ ชีท "ระดับกลาง" มีอยู่แล้ว');
      return { success: true, message: 'ชีท "ระดับกลาง" มีอยู่แล้ว' };
    }
  } catch (error) {
    Logger.log('❌ เกิดข้อผิดพลาดในการสร้างชีทระดับกลาง: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * ฟังก์ชันสำหรับสร้างชีทระดับสูง
 */
function initHighSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    let sheet = ss.getSheetByName('ระดับสูง');
    if (!sheet) {
      sheet = ss.insertSheet('ระดับสูง');
      
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
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setValues([headers]);
      
      // จัดรูปแบบหัวตาราง
      headerRange.setBackground('#4CAF50');
      headerRange.setFontColor('#FFFFFF');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
      headerRange.setVerticalAlignment('middle');
      headerRange.setFontSize(11);
      
      // ปรับขนาดคอลัมน์
      sheet.setColumnWidth(1, 100);  // school_id
      sheet.setColumnWidth(2, 200);  // โรงเรียน
      sheet.setColumnWidth(3, 80);   // เขต
      sheet.setColumnWidth(4, 80);   // ระดับ
      sheet.setColumnWidth(5, 120);  // รายการแข่ง
      sheet.setColumnWidth(6, 70);   // สนาม
      sheet.setColumnWidth(7, 90);   // score1
      sheet.setColumnWidth(8, 90);   // retry1
      sheet.setColumnWidth(9, 100);  // time1
      sheet.getRange(2, 9, sheet.getMaxRows(), 1).setNumberFormat('@'); // time1 เป็น text
      sheet.setColumnWidth(10, 90);  // score2
      sheet.setColumnWidth(11, 90);  // retry2
      sheet.setColumnWidth(12, 100); // time2
      sheet.getRange(2, 12, sheet.getMaxRows(), 1).setNumberFormat('@'); // time2 เป็น text
      sheet.setColumnWidth(13, 150); // timestamp
      
      // ตรึงแถวหัวตาราง
      sheet.setFrozenRows(1);
      
      Logger.log('✓ สร้างชีท "ระดับสูง" สำเร็จ');
      return { success: true, message: 'สร้างชีท "ระดับสูง" สำเร็จ' };
    } else {
      Logger.log('ℹ️ ชีท "ระดับสูง" มีอยู่แล้ว');
      return { success: true, message: 'ชีท "ระดับสูง" มีอยู่แล้ว' };
    }
  } catch (error) {
    Logger.log('❌ เกิดข้อผิดพลาดในการสร้างชีทระดับสูง: ' + error.toString());
    return { success: false, message: error.toString() };
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
      
      // สร้างหัวตาราง
      settingsSheet.getRange('A1').setValue('Key');
      settingsSheet.getRange('B1').setValue('Value');
      
      // จัดรูปแบบหัวตาราง
      const headerRange = settingsSheet.getRange('A1:B1');
      headerRange.setBackground('#1976D2');
      headerRange.setFontColor('#FFFFFF');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
      
      // เพิ่มข้อมูลเริ่มต้น
      settingsSheet.appendRow(['password', 'abcd']);
      settingsSheet.appendRow(['defaultFilter', '']);
      
      // จัดรูปแบบข้อมูล
      settingsSheet.getRange('A2:A3').setBackground('#FFE599');
      settingsSheet.getRange('A2:A3').setFontWeight('bold');
      
      // ปรับขนาดคอลัมน์
      settingsSheet.setColumnWidth(1, 200);
      settingsSheet.setColumnWidth(2, 400);
      
      // ตรึงแถวหัวตาราง
      settingsSheet.setFrozenRows(1);
      
      Logger.log('✓ สร้างชีท "การตั้งค่า" สำเร็จ (Key-Value System)');
    } else {
      Logger.log('ℹ️ ชีท "การตั้งค่า" มีอยู่แล้ว');
    }
    
    // ===== 2. สร้างชีท "ระดับกลาง" และ "ระดับสูง" =====
    initMidSheet();
    initHighSheet();
    
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
    const sheetOrder = ['การตั้งค่า', 'ระดับกลาง', 'ระดับสูง', 'ตัวอย่างข้อมูล'];
    sheetOrder.forEach((sheetName, index) => {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        ss.setActiveSheet(sheet);
        ss.moveActiveSheet(index + 1);
      }
    });
    
    // ตั้งชีท "ระดับกลาง" เป็นชีทที่เปิดอยู่
    ss.setActiveSheet(ss.getSheetByName('ระดับกลาง'));
    
    Logger.log('=================================');
    Logger.log('✅ สร้างโครงสร้างชีททั้งหมดสำเร็จ!');
    Logger.log('=================================');
    Logger.log('ชีทที่สร้าง:');
    Logger.log('1. การตั้งค่า - ระบบ Key-Value (password: abcd)');
    Logger.log('2. ระดับกลาง - พร้อมหัวตารางและรอข้อมูล');
    Logger.log('3. ระดับสูง - พร้อมหัวตารางและรอข้อมูล');
    Logger.log('4. ตัวอย่างข้อมูล - มีข้อมูลตัวอย่าง 3 ทีม');
    Logger.log('=================================');
    
    // แสดง dialog แจ้งเตือน
    SpreadsheetApp.getUi().alert(
      'สำเร็จ! ✅\n\n' +
      'สร้างโครงสร้างชีททั้งหมดเรียบร้อยแล้ว\n\n' +
      '✓ การตั้งค่า - ระบบ Key-Value (รหัสผ่าน: abcd)\n' +
      '✓ ระดับกลาง - พร้อมใช้งาน\n' +
      '✓ ระดับสูง - พร้อมใช้งาน\n' +
      '✓ ตัวอย่างข้อมูล - มีข้อมูลตัวอย่าง\n\n' +
      'ตอนนี้สามารถ Deploy เป็น Web App และใช้งานได้เลย'
    );
    
    return {
      success: true,
      message: 'สร้างโครงสร้างชีททั้งหมดสำเร็จ',
      sheets: ['การตั้งค่า', 'ระดับกลาง', 'ระดับสูง', 'ตัวอย่างข้อมูล']
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
 * ฟังก์ชันสำหรับตรวจสอบ authentication (รองรับ multi-user)
 * @param {string} username - ชื่อผู้ใช้
 * @param {string} password - รหัสผ่าน
 * @param {string} level - ระดับ ('mid' หรือ 'high')
 * @return {object} - ผลการตรวจสอบ
 */
function authenticate(params) {
  try {
    const username = params.username || '';
    const password = params.password || '';
    const level = params.level || 'mid';
    
    // ตรวจสอบ superadmin (เข้าถึงได้ทุกระดับ)
    const superadminUser = getConfigValue('superadmin', 'superadmin');
    const superadminPass = getConfigValue('superadmin_password', 'amp');
    
    if (username === superadminUser && password === superadminPass) {
      return {
        success: true,
        role: 'superadmin',
        level: 'all',
        username: username
      };
    }
    
    // ตรวจสอบ admin ตามระดับ
    if (level === 'mid') {
      // admin_mid1
      const admin1User = getConfigValue('admin_mid1', 'admin1');
      const admin1Pass = getConfigValue('admin_mid1_password', '1234');
      if (username === admin1User && password === admin1Pass) {
        return {
          success: true,
          role: 'admin',
          level: 'mid',
          username: username
        };
      }
      
      // admin_mid2
      const admin2User = getConfigValue('admin_mid2', 'admin2');
      const admin2Pass = getConfigValue('admin_mid2_password', '1234');
      if (username === admin2User && password === admin2Pass) {
        return {
          success: true,
          role: 'admin',
          level: 'mid',
          username: username
        };
      }
    } else if (level === 'high') {
      // admin_high1
      const admin1User = getConfigValue('admin_high1', 'admin1');
      const admin1Pass = getConfigValue('admin_high1_password', '1234');
      if (username === admin1User && password === admin1Pass) {
        return {
          success: true,
          role: 'admin',
          level: 'high',
          username: username
        };
      }
      
      // admin_high2
      const admin2User = getConfigValue('admin_high2', 'admin2');
      const admin2Pass = getConfigValue('admin_high2_password', '1234');
      if (username === admin2User && password === admin2Pass) {
        return {
          success: true,
          role: 'admin',
          level: 'high',
          username: username
        };
      }
    }
    
    return {
      success: false,
      message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง'
    };
    
  } catch (error) {
    Logger.log('Error in authenticate: ' + error.toString());
    return {
      success: false,
      message: 'เกิดข้อผิดพลาด: ' + error.toString()
    };
  }
}

/**
 * ฟังก์ชันสำหรับบันทึกค่า default filter ลงชีท "การตั้งค่า"
 */
function saveDefaultFilter(params) {
  try {
    const level = params.level || 'mid';
    const key = level === 'mid' ? 'defaultFilter_mid' : 'defaultFilter_high';
    
    // เพิ่ม logging เพื่อ debug
    Logger.log('saveDefaultFilter called with level: ' + level);
    Logger.log('Using key: ' + key);
    Logger.log('Filter data: ' + params.filterData);
    
    setConfigValue(key, params.filterData);
    
    return {
      success: true,
      message: 'บันทึก Default Filter สำเร็จ (key: ' + key + ')'
    };
    
  } catch (error) {
    Logger.log('Error in saveDefaultFilter: ' + error.toString());
    return {
      success: false,
      message: 'เกิดข้อผิดพลาด: ' + error.toString()
    };
  }
}

/**
 * ฟังก์ชันสำหรับอ่านค่า default filter จากชีท "การตั้งค่า"
 */
function getDefaultFilter(params) {
  try {
    // อ่านค่าทั้ง mid และ high พร้อมกัน
    const filterDataMid = getConfigValue('defaultFilter_mid', null);
    const filterDataHigh = getConfigValue('defaultFilter_high', null);
    
    return {
      success: true,
      data: {
        mid: filterDataMid,
        high: filterDataHigh
      }
    };
    
  } catch (error) {
    Logger.log('Error in getDefaultFilter: ' + error.toString());
    return {
      success: false,
      message: 'เกิดข้อผิดพลาด: ' + error.toString()
    };
  }
}

/**
 * ฟังก์ชัน Helper สำหรับอ่านค่าจากชีทการตั้งค่า (Key-Value)
 */
/**
 * ฟังก์ชัน Helper สำหรับอ่านค่า config ทั้งหมดจากชีทการตั้งค่า (Key-Value)
 */
function getAllConfigValues() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let settingsSheet = ss.getSheetByName('การตั้งค่า');
    
    if (!settingsSheet) {
      Logger.log('ไม่พบชีท "การตั้งค่า"');
      return {};
    }
    
    const lastRow = settingsSheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('ชีท "การตั้งค่า" ไม่มีข้อมูล');
      return {};
    }
    
    // อ่านข้อมูลทั้งหมดจากชีท (เริ่มแถว 2 เพราะแถว 1 เป็นหัวตาราง)
    const data = settingsSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    
    const config = {};
    for (let i = 0; i < data.length; i++) {
      const key = String(data[i][0]).trim();
      const value = data[i][1];
      
      if (key) {  // ถ้ามี key
        // Trim whitespace ถ้าเป็น string
        config[key] = typeof value === 'string' ? value.trim() : value;
      }
    }
    
    Logger.log('✓ อ่าน settings ทั้งหมด: ' + Object.keys(config).length + ' keys');
    return config;
    
  } catch (error) {
    Logger.log('Error in getAllConfigValues: ' + error.toString());
    return {};
  }
}

/**
 * ฟังก์ชัน Helper สำหรับอ่านค่า config แบบทีละตัว (backward compatible)
 */
function getConfigValue(key, defaultValue = null) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let settingsSheet = ss.getSheetByName('การตั้งค่า');
    
    if (!settingsSheet) {
      return defaultValue;
    }
    
    const lastRow = settingsSheet.getLastRow();
    if (lastRow < 2) {
      return defaultValue;
    }
    
    // อ่านข้อมูลทั้งหมดจากชีท
    const data = settingsSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    
    // หา key ที่ตรงกัน (trim ทั้งสองฝั่งเพื่อเปรียบเทียบ)
    for (let i = 0; i < data.length; i++) {
      const sheetKey = String(data[i][0]).trim();
      const searchKey = String(key).trim();
      
      if (sheetKey === searchKey) {
        const value = data[i][1];
        
        // ถ้าค่าเป็น empty string หรือ null/undefined ให้ใช้ default
        if (value === null || value === undefined || value === '') {
          return defaultValue;
        }
        // Trim whitespace จากค่าที่อ่านได้
        const finalValue = typeof value === 'string' ? value.trim() : value;
        return finalValue;
      }
    }
    
    return defaultValue;
    
  } catch (error) {
    Logger.log('Error in getConfigValue: ' + error.toString());
    return defaultValue;
  }
}

/**
 * ฟังก์ชัน Helper สำหรับบันทึกค่าลงชีทการตั้งค่า (Key-Value)
 */
function setConfigValue(key, value) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let settingsSheet = ss.getSheetByName('การตั้งค่า');
    
    if (!settingsSheet) {
      // สร้างชีทใหม่ถ้ายังไม่มี
      settingsSheet = ss.insertSheet('การตั้งค่า');
      
      // สร้างหัวตาราง
      settingsSheet.getRange('A1').setValue('Key');
      settingsSheet.getRange('B1').setValue('Value');
      
      // จัดรูปแบบหัวตาราง
      const headerRange = settingsSheet.getRange('A1:B1');
      headerRange.setBackground('#1976D2');
      headerRange.setFontColor('#FFFFFF');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
      
      // ปรับขนาดคอลัมน์
      settingsSheet.setColumnWidth(1, 200);
      settingsSheet.setColumnWidth(2, 400);
      
      // ตรึงแถวหัวตาราง
      settingsSheet.setFrozenRows(1);
    }
    
    const lastRow = settingsSheet.getLastRow();
    let found = false;
    
    // หา key ที่มีอยู่แล้ว
    if (lastRow > 1) {
      const data = settingsSheet.getRange(2, 1, lastRow - 1, 2).getValues(); // อ่านทั้ง Key และ Value
      for (let i = 0; i < data.length; i++) {
        // Trim key ก่อนเปรียบเทียบเพื่อหลีกเลี่ยง whitespace
        if (String(data[i][0]).trim() === String(key).trim()) {
          // อัพเดตค่าเดิม (trim value ก่อนเขียน)
          const trimmedValue = typeof value === 'string' ? value.trim() : value;
          settingsSheet.getRange(i + 2, 2).setValue(trimmedValue);
          found = true;
          break;
        }
      }
    }
    
    // ถ้าไม่เจอ ให้เพิ่มแถวใหม่ (trim ทั้ง key และ value)
    if (!found) {
      const trimmedKey = typeof key === 'string' ? key.trim() : key;
      const trimmedValue = typeof value === 'string' ? value.trim() : value;
      settingsSheet.appendRow([trimmedKey, trimmedValue]);
    }
    
    return true;
    
  } catch (error) {
    Logger.log('Error in setConfigValue: ' + error.toString());
    return false;
  }
}
function getAllTeams() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetMid = ss.getSheetByName('ระดับกลาง');
    const sheetHigh = ss.getSheetByName('ระดับสูง');
    
    // โหลดการตั้งค่าทั้งหมดจากชีท "การตั้งค่า" (key-value pairs)
    const settings = getAllConfigValues();
    
    // ฟังก์ชันช่วยสำหรับแปลง sheet เป็น array of teams
    function processSheet(sheet) {
      if (!sheet || sheet.getLastRow() <= 1) {
        return [];
      }
      
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
          const key = fieldMapping[header] || header;
          let value = data[i][j];
          
          if (key === 'time1' || key === 'time2') {
            if (value && value.trim() !== '') {
              team[key] = value.trim();
            } else {
              team[key] = '0:00:00';
            }
          } else if (key === 'score1' || key === 'score2' || key === 'retry1' || key === 'retry2') {
            team[key] = parseInt(value) || 0;
          } else {
            team[key] = value;
          }
        }
        teams.push(team);
      }
      
      return teams;
    }
    
    // ประมวลผลทั้งสอง sheet
    const midTeams = processSheet(sheetMid);
    const highTeams = processSheet(sheetHigh);
    
    // แยก defaultFilter_mid และ defaultFilter_high ออกจาก settings
    const defaultFilter_mid = settings.defaultFilter_mid || null;
    const defaultFilter_high = settings.defaultFilter_high || null;
    
    return {
      success: true,
      data: midTeams,
      highTeams: highTeams,
      settings: settings,
      defaultFilter_mid: defaultFilter_mid,
      defaultFilter_high: defaultFilter_high
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
    
    // เลือก sheet ตาม level parameter (default เป็น mid)
    const level = params.level || 'mid';
    const sheetName = level === 'high' ? 'ระดับสูง' : 'ระดับกลาง';
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return {
        success: false,
        message: 'ไม่พบชีทข้อมูล: ' + sheetName
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
          message: 'อัพเดตคะแนนสำเร็จ (' + sheetName + ')'
        };
      }
    }
    
    return {
      success: false,
      message: 'ไม่พบข้อมูลทีมใน ' + sheetName
    };
    
  } catch (error) {
    Logger.log('Error in updateScore: ' + error.toString());
    return {
      success: false,
      message: 'เกิดข้อผิดพลาด: ' + error.toString()
    };
  }
}
