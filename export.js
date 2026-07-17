/**
 * Excel Generation and Export Module
 */

import { formatDate } from './excel.js';

// Cache for parsed template workbook to avoid redundant network fetches and parsing
let cachedTemplateWb = null;


// Deep clones a SheetJS worksheet object including all cell styles and structures
function cloneSheet(ws) {
  const newSheet = {};
  for (const key in ws) {
    if (key.startsWith('!')) {
      if (key === '!merges') {
        newSheet[key] = ws[key].map(m => ({
          s: { r: m.s.r, c: m.s.c },
          e: { r: m.e.r, c: m.e.c }
        }));
      } else if (key === '!cols' || key === '!rows') {
        newSheet[key] = ws[key].map(x => (x ? { ...x } : null));
      } else {
        newSheet[key] = JSON.parse(JSON.stringify(ws[key]));
      }
    } else {
      // Copy cell object and its style properties
      newSheet[key] = { ...ws[key] };
      if (ws[key].s) {
        newSheet[key].s = { ...ws[key].s };
      }
    }
  }
  return newSheet;
}

// Matches a room number to its template worksheet name
function findRoomSheetName(workbook, roomNumber, building) {
  const cleanRoom = roomNumber.trim().replace(/\s+/g, '').replace(/-/g, '').toUpperCase();
  const cleanBldg = building.trim().replace(/\s+/g, '').replace(/-/g, '').toUpperCase();
  
  const targets = [
    cleanRoom,
    cleanBldg + cleanRoom,
    cleanBldg.slice(0, 3) + cleanRoom,
    cleanBldg.slice(0, 2) + cleanRoom,
    cleanRoom + cleanBldg
  ];

  for (const name of workbook.SheetNames) {
    const cleanName = name.replace(/\s+/g, '').replace(/-/g, '').toUpperCase();
    if (targets.includes(cleanName) || cleanName === cleanRoom) {
      return name;
    }
  }
  return null;
}

/**
 * Generates the final seating plan Excel workbook and downloads it.
 * @param {Object} seatingByRoom - Generated seating arrangement
 * @param {string} examDate - Selected Date
 * @param {string} examSession - Selected Session
 * @param {string} examTiming - Selected Session Custom timing
 * @param {number} studentsPerBench - 2 or 3 students per bench
 * @returns {Promise<void>}
 */
export function exportSeatingWorkbook(seatingByRoom, examDate, examSession, examTiming, studentsPerBench) {
  return new Promise(async (resolve, reject) => {
    try {
      // 1. Fetch the master template Excel file (use cache if available)
      let templateWb;
      if (cachedTemplateWb) {
        templateWb = cachedTemplateWb;
      } else {
        const response = await fetch('Seating Plan Master copy 3-AN-S1.xlsx');
        if (!response.ok) {
          throw new Error('Failed to load master template Excel. Make sure Seating Plan Master copy 3-AN-S1.xlsx is present.');
        }
        const arrayBuffer = await response.arrayBuffer();
        // Read with style option enabled
        templateWb = XLSX.read(arrayBuffer, { type: 'array', cellStyles: true });
        cachedTemplateWb = templateWb;
      }
      
      // 2. Prepare new workbook
      const outputWb = XLSX.utils.book_new();
      
      // Copy Room List and student lists sheets as is
      const roomListSheet = templateWb.Sheets['Room List'];
      if (roomListSheet) {
        XLSX.utils.book_append_sheet(outputWb, cloneSheet(roomListSheet), 'Room List');
      }

      // 3. Process and write seating data to each room sheet
      const roomKeys = Object.keys(seatingByRoom).filter(k => seatingByRoom[k].benches.some(b => b.seats.some(s => s !== null)));
      
      for (const roomNum of roomKeys) {
        const roomData = seatingByRoom[roomNum];
        const room = roomData.room;
        const benches = roomData.benches;
        
        // Find existing sheet name or clone default
        const origSheetName = findRoomSheetName(templateWb, room.roomNumber, room.building);
        let ws;
        if (origSheetName && templateWb.Sheets[origSheetName]) {
          ws = cloneSheet(templateWb.Sheets[origSheetName]);
        } else {
          // If no worksheet matches, clone VBA 101A as the generic layout template
          const fallbackName = templateWb.Sheets['VBA 101A'] ? 'VBA 101A' : templateWb.SheetNames[3];
          ws = cloneSheet(templateWb.Sheets[fallbackName]);
        }

        // Set Date, Timing, Room No in header
        ws['B2'] = { v: formatDate(examDate), t: 's', s: ws['B2'] ? ws['B2'].s : null };
        ws['E2'] = { v: examTiming, t: 's', s: ws['E2'] ? ws['E2'].s : null };
        ws['K2'] = { v: `${room.building} - ${room.roomNumber}`, t: 's', s: ws['K2'] ? ws['K2'].s : null };

        // Group benches by column
        const colsData = {};
        benches.forEach(b => {
          if (!colsData[b.colIndex]) colsData[b.colIndex] = [];
          colsData[b.colIndex].push(b);
        });

        // Track statistics
        const semPresent = {};
        let roomSeatedCount = 0;

        // Clean out default template formulas and write static values
        const colIndices = Object.keys(colsData).sort((a, b) => parseInt(a) - parseInt(b));
        
        colIndices.forEach(colIdx => {
          const colBenches = colsData[colIdx];
          
          // Clear any dynamic array references in header cells (row 8 and 10)
          const seatColsSheet = [];
          if (studentsPerBench === 2) {
            seatColsSheet.push((colIdx - 1) * 3, (colIdx - 1) * 3 + 1);
          } else {
            seatColsSheet.push((colIdx - 1) * 4, (colIdx - 1) * 4 + 1, (colIdx - 1) * 4 + 2);
          }

          // Gather unique courses and semesters for the column
          const coursesInCol = [new Set(), new Set(), new Set()];
          const semsInCol = [new Set(), new Set(), new Set()];

          colBenches.forEach(bench => {
            bench.seats.forEach((seat, seatIdx) => {
              if (seat) {
                coursesInCol[seatIdx].add(seat.courseCode);
                semsInCol[seatIdx].add(seat.semester);
                
                semPresent[seat.semester] = (semPresent[seat.semester] || 0) + 1;
                roomSeatedCount++;
              }
            });
          });

          // Write course code (row 8) and semester (row 10) headers
          seatColsSheet.forEach((colIdxSheet, seatIdx) => {
            const courseRef = XLSX.utils.encode_cell({ r: 7, c: colIdxSheet }); // Row 8
            const semRef = XLSX.utils.encode_cell({ r: 9, c: colIdxSheet });    // Row 10
            
            const origCourseStyle = ws[courseRef] ? ws[courseRef].s : null;
            const origSemStyle = ws[semRef] ? ws[semRef].s : null;
            
            const coursesStr = [...coursesInCol[seatIdx]].join('/');
            const semsStr = [...semsInCol[seatIdx]].join('/');
            
            ws[courseRef] = { v: coursesStr || '', t: 's', s: origCourseStyle };
            ws[semRef] = { v: semsStr ? `Sem: ${semsStr}` : '', t: 's', s: origSemStyle };
            
            // Clean up any old formula descriptors
            delete ws[courseRef].f;
            delete ws[semRef].f;
          });

          // Populate student rows (Row 11 to 16)
          colBenches.forEach(bench => {
            const rowIdxSheet = 10 + bench.benchIndex; // 1-based Row Index in Excel (11 for Bench 1)
            
            bench.seats.forEach((seat, seatIdx) => {
              const colIdxSheet = seatColsSheet[seatIdx];
              const cellRef = XLSX.utils.encode_cell({ r: rowIdxSheet - 1, c: colIdxSheet });
              
              let cellStyle = null;
              if (ws[cellRef] && ws[cellRef].s) {
                cellStyle = ws[cellRef].s;
              } else {
                // If cell does not exist (like Seat C), copy style from Seat B in same row
                const seatBCellRef = XLSX.utils.encode_cell({ r: rowIdxSheet - 1, c: seatColsSheet[1] });
                cellStyle = ws[seatBCellRef] ? ws[seatBCellRef].s : null;
              }

              ws[cellRef] = { v: seat ? seat.registrationNumber : '', t: 's', s: cellStyle };
              delete ws[cellRef].f; // Ensure formulas are removed
            });
          });

          // Clear remaining unused template rows in this column
          for (let rIdx = 10 + colBenches.length + 1; rIdx <= 16; rIdx++) {
            seatColsSheet.forEach(colIdxSheet => {
              const cellRef = XLSX.utils.encode_cell({ r: rIdx - 1, c: colIdxSheet });
              if (ws[cellRef]) {
                ws[cellRef] = { v: '', t: 's', s: ws[cellRef].s };
                delete ws[cellRef].f;
              }
            });
          }

          // Calculate student index ranges for this column (Rows 12 and 13 in columns T to Y)
          const studentIndices = [];
          colBenches.forEach(bench => {
            bench.seats.forEach(seat => {
              if (seat) {
                studentIndices.push(seat.id + 1); // 1-based index in the filtered student list
              }
            });
          });

          const numericColIdx = parseInt(colIdx);
          const rangeColSheet = 18 + numericColIdx; // colIdx is 1-based (e.g. 1 -> Col 19 / T)
          const startRef = XLSX.utils.encode_cell({ r: 11, c: rangeColSheet }); // Row 12 (Student Index Start)
          const endRef = XLSX.utils.encode_cell({ r: 12, c: rangeColSheet });   // Row 13 (Student Index End)
          
          const startStyle = ws[startRef] ? ws[startRef].s : null;
          const endStyle = ws[endRef] ? ws[endRef].s : null;
          
          if (studentIndices.length > 0) {
            const minIdx = Math.min(...studentIndices);
            const maxIdx = Math.max(...studentIndices);
            ws[startRef] = { v: minIdx, t: 'n', s: startStyle };
            ws[endRef] = { v: maxIdx, t: 'n', s: endStyle };
          } else {
            ws[startRef] = { v: '', t: 's', s: startStyle };
            ws[endRef] = { v: '', t: 's', s: endStyle };
          }
          if (ws[startRef]) delete ws[startRef].f;
          if (ws[endRef]) delete ws[endRef].f;
        });

        // Populate Present/Absent counts in footer (starts at row 18)
        let fRow = 18;
        const semestersList = Object.keys(semPresent).sort();
        semestersList.forEach(sem => {
          const semCell = `A${fRow}`;
          const absentCell = `J${fRow}`;
          const presentCell = `K${fRow}`;
          
          ws[semCell] = { v: `Sem: ${sem}`, t: 's', s: ws[semCell] ? ws[semCell].s : null };
          ws[absentCell] = { v: 0, t: 'n', s: ws[absentCell] ? ws[absentCell].s : null };
          ws[presentCell] = { v: semPresent[sem], t: 'n', s: ws[presentCell] ? ws[presentCell].s : null };
          
          delete ws[semCell].f;
          delete ws[absentCell].f;
          delete ws[presentCell].f;
          fRow++;
        });

        // Populate Total Capacity and Total Present summaries
        const totalCapacity = benches.length * studentsPerBench;
        
        ws[`A${fRow}`] = { v: 'Total Present', t: 's', s: ws[`A18`] ? ws[`A18`].s : null };
        ws[`K${fRow}`] = { v: roomSeatedCount, t: 'n', s: ws[`K18`] ? ws[`K18`].s : null };
        delete ws[`A${fRow}`].f;
        delete ws[`K${fRow}`].f;
        fRow++;

        ws[`A${fRow}`] = { v: 'Total Capacity', t: 's', s: ws[`A18`] ? ws[`A18`].s : null };
        ws[`K${fRow}`] = { v: totalCapacity, t: 'n', s: ws[`K18`] ? ws[`K18`].s : null };
        delete ws[`A${fRow}`].f;
        delete ws[`K${fRow}`].f;

        // Clear remaining footer table cells in range to avoid residual template data
        for (let r = fRow + 1; r <= 20; r++) {
          if (ws[`A${r}`]) delete ws[`A${r}`];
          if (ws[`J${r}`]) delete ws[`J${r}`];
          if (ws[`K${r}`]) delete ws[`K${r}`];
        }

        // Adjust column widths if 3 students per bench is chosen (insert column width)
        if (studentsPerBench === 3 && ws['!cols']) {
          const newCols = [];
          for (let i = 0; i < ws['!cols'].length; i++) {
            newCols.push(ws['!cols'][i]);
            // If it's Seat B column, insert a new width for Seat C
            if (i % 3 === 1 && i <= 10) {
              newCols.push({ ...ws['!cols'][i] });
            }
          }
          ws['!cols'] = newCols;
        }

        // Append to output workbook
        XLSX.utils.book_append_sheet(outputWb, ws, room.roomNumber);
      }

      // 4. Generate Summary Sheet
      const summaryData = [['Room', 'Students', 'Capacity', 'Empty Seats', 'Course Count', 'Semester Count']];
      roomKeys.forEach(roomNum => {
        const roomData = seatingByRoom[roomNum];
        const capacity = roomData.benches.length * studentsPerBench;
        const occupied = roomData.benches.reduce((sum, b) => sum + b.seats.filter(s => s !== null).length, 0);
        const empty = capacity - occupied;
        
        const courses = new Set();
        const semesters = new Set();
        roomData.benches.forEach(b => {
          b.seats.forEach(s => {
            if (s) {
              courses.add(s.courseCode);
              semesters.add(s.semester);
            }
          });
        });
        
        summaryData.push([
          roomNum,
          occupied,
          capacity,
          empty,
          courses.size,
          semesters.size
        ]);
      });
      const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);
      
      // Style Summary Sheet headers
      for (let c = 0; c < 6; c++) {
        const cellRef = XLSX.utils.encode_cell({ r: 0, c });
        if (summarySheet[cellRef]) {
          summarySheet[cellRef].s = {
            font: { bold: true, color: { rgb: 'FFFFFF' } },
            fill: { fgColor: { rgb: '4F46E5' } }, // Indigo Header
            alignment: { horizontal: 'center' }
          };
        }
      }
      XLSX.utils.book_append_sheet(outputWb, summarySheet, 'Seating Summary');

      // 5. Generate Used Rooms Report Sheet
      const usedRoomsData = [['Room', 'Capacity', 'Occupied', 'Remaining', 'Building', 'Floor']];
      roomKeys.forEach(roomNum => {
        const roomData = seatingByRoom[roomNum];
        const room = roomData.room;
        const capacity = roomData.benches.length * studentsPerBench;
        const occupied = roomData.benches.reduce((sum, b) => sum + b.seats.filter(s => s !== null).length, 0);
        const remaining = capacity - occupied;
        
        usedRoomsData.push([
          roomNum,
          capacity,
          occupied,
          remaining,
          room.building,
          room.floor
        ]);
      });
      const usedRoomsSheet = XLSX.utils.aoa_to_sheet(usedRoomsData);
      
      // Style Used Rooms Report headers
      for (let c = 0; c < 6; c++) {
        const cellRef = XLSX.utils.encode_cell({ r: 0, c });
        if (usedRoomsSheet[cellRef]) {
          usedRoomsSheet[cellRef].s = {
            font: { bold: true, color: { rgb: 'FFFFFF' } },
            fill: { fgColor: { rgb: '10B981' } }, // Emerald Header
            alignment: { horizontal: 'center' }
          };
        }
      }
      XLSX.utils.book_append_sheet(outputWb, usedRoomsSheet, 'Used Rooms Report');

      // 6. Write and trigger browser download (using binary array buffer and Blob for high performance)
      const wbout = XLSX.write(outputWb, { bookType: 'xlsx', type: 'array', cellStyles: true });
      const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const filename = `Seating_Plan_${formatDate(examDate)}_${examSession}.xlsx`;

      if (typeof saveAs !== 'undefined') {
        saveAs(blob, filename);
      } else {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        setTimeout(() => {
          document.body.removeChild(a);
          URL.revokeObjectURL(url);
        }, 100);
      }

      resolve();
    } catch (err) {
      reject(new Error('Export failed: ' + err.message));
    }
  });
}
