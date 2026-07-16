/**
 * Excel Parsing and Data Normalization Module
 */

/**
 * Utility to match and retrieve values from row object by case-insensitive key names
 * @param {Object} row - Raw JSON row object from SheetJS
 * @param {Array<string>} keyVariants - Acceptable header names
 * @returns {*} The value of the matched key or null
 */
function getVal(row, keyVariants) {
  for (const k in row) {
    const kClean = k.toLowerCase().replace(/[^a-z0-9]/g, '');
    for (const variant of keyVariants) {
      const vClean = variant.toLowerCase().replace(/[^a-z0-9]/g, '');
      if (kClean === vClean) {
        return row[k];
      }
    }
  }
  return null;
}

/**
 * Format date value to stable YYYY-MM-DD string
 * @param {*} val - Date object, serial number, or string
 * @returns {string} Formatted date string
 */
export function formatDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    // Compensate for Excel midnight floating-point rounding errors (e.g. 23:59:50 of previous day) by shifting to midday
    const adjusted = new Date(val.getTime() + 12 * 60 * 60 * 1000);
    const y = adjusted.getFullYear();
    const m = String(adjusted.getMonth() + 1).padStart(2, '0');
    const d = String(adjusted.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  if (typeof val === 'number') {
    // SheetJS parses serial numbers sometimes
    const date = new Date(Math.round((val - 25569) * 86400 * 1000));
    const adjusted = new Date(date.getTime() + 12 * 60 * 60 * 1000);
    const y = adjusted.getFullYear();
    const m = String(adjusted.getMonth() + 1).padStart(2, '0');
    const d = String(adjusted.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  
  // If it's a string, try parsing it or clean it
  const str = String(val).trim();
  // If string contains date in format like "2026-03-06T00:00:00"
  if (str.includes('T')) {
    return str.split('T')[0];
  }
  return str;
}

/**
 * Parses Student List Excel workbook.
 * Reads data from Sheet 2 (which is index 1 of the worksheets array).
 * @param {ArrayBuffer} arrayBuffer - The uploaded Excel file buffer
 * @returns {Promise<Array<Object>>} Parsed and normalized student objects
 */
export function parseStudentExcel(arrayBuffer) {
  return new Promise((resolve, reject) => {
    try {
      const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
      const combinedStudents = [];
      let globalIndex = 0;
      console.log('parseStudentExcel: sheet names in workbook:', workbook.SheetNames);

      for (const sheetName of workbook.SheetNames) {
        // Skip metadata/room sheets
        if (sheetName.toLowerCase() === 'room list' || sheetName.toLowerCase() === 'dashboard') {
          console.log(`parseStudentExcel: skipping sheet ${sheetName}`);
          continue;
        }

        const worksheet = workbook.Sheets[sheetName];
        const rawData = XLSX.utils.sheet_to_json(worksheet, { defval: null });
        console.log(`parseStudentExcel: sheet ${sheetName} raw rows:`, rawData.length);

        const sheetStudents = rawData.map(row => {
          const regNo = getVal(row, ['Registration number', 'Registration No', 'RegNo', 'Reg No', 'RegistrationNumber']);
          const courseCode = getVal(row, ['Course code', 'Course Code', 'SubCode', 'Sub Code', 'Subject Code']);
          const courseName = getVal(row, ['Course name', 'Course Name', 'SubName', 'Sub Name', 'Subject Name']);
          const branch = getVal(row, ['Branch', 'Dept', 'Department']);
          const semester = getVal(row, ['SEMESTER', 'Semester', 'Sem']);
          const dateRaw = getVal(row, ['Date', 'Exam Date']);
          const session = getVal(row, ['Session', 'Exam Session', 'Session (FN/AN)']);

          // If no registration number, it's not a valid student row in this sheet
          if (!regNo) return null;

          return {
            id: globalIndex++,
            registrationNumber: String(regNo).trim(),
            courseCode: courseCode ? String(courseCode).trim() : 'UNKNOWN',
            courseName: courseName ? String(courseName).trim() : 'UNKNOWN',
            branch: branch ? String(branch).trim() : 'GENERAL',
            semester: semester ? String(semester).trim() : 'I',
            date: formatDate(dateRaw),
            session: session ? String(session).trim().toUpperCase() : 'FN'
          };
        }).filter(s => s !== null);

        console.log(`parseStudentExcel: sheet ${sheetName} parsed students:`, sheetStudents.length);
        combinedStudents.push(...sheetStudents);
      }

      console.log('parseStudentExcel: total combined students:', combinedStudents.length);
      resolve(combinedStudents);
    } catch (err) {
      console.error('parseStudentExcel error:', err);
      reject(new Error('Failed to parse Student Excel: ' + err.message));
    }
  });
}

/**
 * Parses Room List Excel workbook.
 * Reads data from Sheet 1 (index 0).
 * @param {ArrayBuffer} arrayBuffer - The uploaded Excel file buffer
 * @returns {Promise<Array<Object>>} Parsed and normalized room objects
 */
export function parseRoomExcel(arrayBuffer) {
  return new Promise((resolve, reject) => {
    try {
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      
      // Header starts on row 2, data on row 3 (we use range header offsets to skip Row 1 description)
      const rawData = XLSX.utils.sheet_to_json(worksheet, { range: 1, defval: null });
      
      const parsedRooms = rawData.map(row => {
        const roomNo = getVal(row, ['Room No.', 'Room No', 'Room Number', 'RoomNo']);
        const building = getVal(row, ['Building', 'Bldg', 'Block']);
        const floor = getVal(row, ['Classroom Floor', 'Floor']);
        const c1 = getVal(row, ['C1 / No. of Benches', 'C1', 'Column 1']);
        const c2 = getVal(row, ['C2 / No. of Benches', 'C2', 'Column 2']);
        const c3 = getVal(row, ['C3 / No. of Benches', 'C3', 'Column 3']);
        const c4 = getVal(row, ['C4 / No. of Benches', 'C4', 'Column 4']);
        const c5 = getVal(row, ['C5 / No. of Benches', 'C5', 'Column 5']);
        const c6 = getVal(row, ['C6 / No. of Benches', 'C6', 'Column 6']);
        
        // Skip header lines or lines without a room number
        if (!roomNo || String(roomNo).toLowerCase().includes('room') || String(roomNo).includes('S.No.')) {
          return null;
        }
        
        const numC1 = c1 !== null && c1 !== '' ? parseInt(c1) : 0;
        const numC2 = c2 !== null && c2 !== '' ? parseInt(c2) : 0;
        const numC3 = c3 !== null && c3 !== '' ? parseInt(c3) : 0;
        const numC4 = c4 !== null && c4 !== '' ? parseInt(c4) : 0;
        const numC5 = c5 !== null && c5 !== '' ? parseInt(c5) : 0;
        const numC6 = c6 !== null && c6 !== '' ? parseInt(c6) : 0;
        
        const totalBenches = numC1 + numC2 + numC3 + numC4 + numC5 + numC6;
        
        return {
          roomNumber: String(roomNo).trim(),
          building: building ? String(building).trim() : 'MAIN',
          floor: floor ? String(floor).trim() : 'Ground',
          c1: numC1,
          c2: numC2,
          c3: numC3,
          c4: numC4,
          c5: numC5,
          c6: numC6,
          totalBenches: totalBenches
        };
      }).filter(r => r !== null && r.totalBenches > 0);
      
      resolve(parsedRooms);
    } catch (err) {
      reject(new Error('Failed to parse Room Excel: ' + err.message));
    }
  });
}
