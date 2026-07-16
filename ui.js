/**
 * UI Rendering and Visualization Module
 */

// Local cache for course color mappings to keep them consistent
const courseColors = {};

/**
 * Generate or retrieve a consistent, premium HSL color for a course code
 * @param {string} courseCode - The course code to color-code
 * @returns {string} HSL color string
 */
function getCourseColor(courseCode) {
  if (!courseCode) return 'rgba(255, 255, 255, 0.05)';
  const key = courseCode.trim().toUpperCase();
  if (courseColors[key]) {
    return courseColors[key];
  }
  
  // Hash the course code to get a deterministic hue
  let hash = 0;
  for (let i = 0; i < key.length; i++) {
    hash = key.charCodeAt(i) + ((hash << 5) - hash);
  }
  const hue = Math.abs(hash) % 360;
  
  // Use HSL for highly readable dark-themed badge background
  courseColors[key] = `hsl(${hue}, 60%, 32%)`;
  return courseColors[key];
}

/**
 * Update top statistics cards on the dashboard
 * @param {Object} stats - Statistics object containing students, subjects, rooms, capacity, and empty
 */
export function updateStats(stats) {
  document.getElementById('stat-total-students').textContent = stats.totalStudents || 0;
  document.getElementById('stat-total-subjects').textContent = stats.totalSubjects || 0;
  document.getElementById('stat-rooms-used').textContent = stats.roomsUsed || 0;
  document.getElementById('stat-total-capacity').textContent = stats.totalCapacity || 0;
  document.getElementById('stat-empty-seats').textContent = stats.emptySeats || 0;
}

/**
 * Show a warning/error banner on the dashboard
 * @param {string} text - Message text to display
 */
export function showWarning(text) {
  const banner = document.getElementById('warning-banner');
  const bannerText = document.getElementById('warning-text');
  bannerText.textContent = text;
  banner.classList.remove('hidden');
}

/**
 * Hide the warning banner
 */
export function hideWarning() {
  const banner = document.getElementById('warning-banner');
  banner.classList.add('hidden');
}

/**
 * Render parsed student list table preview
 * @param {Array<Object>} students - List of students
 */
export function renderStudentsTable(students) {
  const tbody = document.getElementById('student-table-body');
  if (!students || students.length === 0) {
    tbody.innerHTML = `<tr><td colspan="7" class="empty-table">No student data imported. Upload Student List.</td></tr>`;
    return;
  }
  
  tbody.innerHTML = students.map(s => `
    <tr>
      <td><strong>${s.registrationNumber}</strong></td>
      <td><span class="badge" style="background-color: ${getCourseColor(s.courseCode)}">${s.courseCode}</span></td>
      <td>${s.courseName}</td>
      <td>${s.branch}</td>
      <td>Sem ${s.semester}</td>
      <td>${s.date}</td>
      <td>${s.session}</td>
    </tr>
  `).join('');
}

/**
 * Render parsed room list table preview
 * @param {Array<Object>} rooms - List of rooms
 */
export function renderRoomsTable(rooms) {
  const tbody = document.getElementById('room-table-body');
  if (!rooms || rooms.length === 0) {
    tbody.innerHTML = `<tr><td colspan="10" class="empty-table">No room data imported. Upload Room List.</td></tr>`;
    return;
  }
  
  tbody.innerHTML = rooms.map(r => `
    <tr>
      <td><strong>${r.roomNumber}</strong></td>
      <td>${r.building}</td>
      <td>${r.floor}</td>
      <td>${r.c1 || '-'}</td>
      <td>${r.c2 || '-'}</td>
      <td>${r.c3 || '-'}</td>
      <td>${r.c4 || '-'}</td>
      <td>${r.c5 || '-'}</td>
      <td>${r.c6 || '-'}</td>
      <td><span class="badge" style="background-color: var(--accent-color)">${r.totalBenches} benches</span></td>
    </tr>
  `).join('');
}

/**
 * Render visual cards for room seating plan preview
 * @param {Object} seatingByRoom - Seating plan grouped by room
 * @param {number} studentsPerBench - 2 or 3 students
 */
export function renderSeatingCards(seatingByRoom, studentsPerBench) {
  const container = document.getElementById('room-cards-container');
  const roomKeys = Object.keys(seatingByRoom).filter(k => seatingByRoom[k].benches.some(b => b.seats.some(s => s !== null)));
  
  if (roomKeys.length === 0) {
    container.innerHTML = `
      <div class="empty-state">
        <i class="fa-solid fa-chair-office"></i>
        <h4>No Seating Plan Generated</h4>
        <p>Configure exam parameters and click "Generate Seating Plan" to view seating layouts.</p>
      </div>
    `;
    return;
  }
  
  container.innerHTML = roomKeys.map(roomNum => {
    const roomData = seatingByRoom[roomNum];
    const room = roomData.room;
    const benches = roomData.benches;
    
    // Group benches by column
    const columns = {};
    benches.forEach(bench => {
      if (!columns[bench.colIndex]) {
        columns[bench.colIndex] = [];
      }
      columns[bench.colIndex].push(bench);
    });
    
    // Calculate statistics for this room
    let totalCapacity = benches.length * studentsPerBench;
    let occupied = benches.reduce((sum, b) => sum + b.seats.filter(s => s !== null).length, 0);
    
    const colKeys = Object.keys(columns).sort((a, b) => parseInt(a) - parseInt(b));
    
    const columnsHtml = colKeys.map(colIdx => {
      const colBenches = columns[colIdx];
      // Get the semesters active in this column (e.g. from seats)
      const semesters = [...new Set(colBenches.flatMap(b => b.seats.filter(s => s !== null).map(s => s.semester)))];
      const semText = semesters.length > 0 ? semesters.join('/') : '-';
      
      const benchesHtml = colBenches.map(bench => {
        const seatsHtml = bench.seats.map((seat, seatIdx) => {
          const seatChar = String.fromCharCode(65 + seatIdx); // A, B, C
          if (seat) {
            const color = getCourseColor(seat.courseCode);
            return `
              <div class="seat occupied" style="background-color: ${color}" title="${seat.courseName} | Semester ${seat.semester}">
                <span class="seat-label">${seatChar}:</span>${seat.registrationNumber}
                <div style="font-size: 8px; opacity: 0.8; margin-top: 1px;">${seat.courseCode} (S${seat.semester})</div>
              </div>
            `;
          } else {
            return `
              <div class="seat empty">
                <span class="seat-label">${seatChar}:</span>Empty
              </div>
            `;
          }
        }).join('');
        
        return `
          <div class="seating-bench">
            <span class="bench-num">Bench ${bench.benchIndex}</span>
            ${seatsHtml}
          </div>
        `;
      }).join('');
      
      return `
        <div class="seating-column">
          <div class="column-header">Col ${colIdx}</div>
          <div class="column-semester" title="Semesters in this column">Sem: ${semText}</div>
          ${benchesHtml}
        </div>
      `;
    }).join('');
    
    return `
      <div class="room-card">
        <div class="room-card-header">
          <div class="room-card-title">
            <i class="fa-solid fa-door-open text-indigo"></i>
            <span>Room ${roomNum} (${room.building})</span>
          </div>
          <div class="room-card-stats">
            Occupied: ${occupied} / ${totalCapacity}
          </div>
        </div>
        <div class="room-card-body">
          ${columnsHtml}
        </div>
      </div>
    `;
  }).join('');
}

/**
 * Render Used Rooms Report and Seating Summary tables
 * @param {Object} seatingByRoom - Seating plan grouped by room
 * @param {number} studentsPerBench - 2 or 3 students
 */
export function renderReports(seatingByRoom, studentsPerBench) {
  const usedRoomsBody = document.getElementById('used-rooms-table-body');
  const summaryBody = document.getElementById('summary-table-body');
  
  const roomKeys = Object.keys(seatingByRoom).filter(k => seatingByRoom[k].benches.some(b => b.seats.some(s => s !== null)));
  
  if (roomKeys.length === 0) {
    usedRoomsBody.innerHTML = `<tr><td colspan="6" class="empty-table">No seating plan generated yet.</td></tr>`;
    summaryBody.innerHTML = `<tr><td colspan="6" class="empty-table">No seating plan generated yet.</td></tr>`;
    return;
  }
  
  // Render Used Rooms Report
  usedRoomsBody.innerHTML = roomKeys.map(roomNum => {
    const roomData = seatingByRoom[roomNum];
    const room = roomData.room;
    const totalCapacity = roomData.benches.length * studentsPerBench;
    const occupied = roomData.benches.reduce((sum, b) => sum + b.seats.filter(s => s !== null).length, 0);
    const remaining = totalCapacity - occupied;
    
    return `
      <tr>
        <td><strong>${roomNum}</strong></td>
        <td>${room.building}</td>
        <td>${room.floor}</td>
        <td>${totalCapacity}</td>
        <td><span class="badge" style="background-color: var(--accent-color); padding: 2px 6px; border-radius: 8px;">${occupied}</span></td>
        <td>${remaining}</td>
      </tr>
    `;
  }).join('');
  
  // Render Seating Summary
  summaryBody.innerHTML = roomKeys.map(roomNum => {
    const roomData = seatingByRoom[roomNum];
    const totalCapacity = roomData.benches.length * studentsPerBench;
    const occupied = roomData.benches.reduce((sum, b) => sum + b.seats.filter(s => s !== null).length, 0);
    const emptySeats = totalCapacity - occupied;
    
    // Find count of unique courses and semesters
    const courses = new Set();
    const semesters = new Set();
    
    roomData.benches.forEach(bench => {
      bench.seats.forEach(seat => {
        if (seat) {
          courses.add(seat.courseCode);
          semesters.add(seat.semester);
        }
      });
    });
    
    return `
      <tr>
        <td><strong>${roomNum}</strong></td>
        <td>${occupied}</td>
        <td>${totalCapacity}</td>
        <td>${emptySeats}</td>
        <td><span class="badge" style="background-color: var(--color-success); padding: 2px 6px; border-radius: 8px;">${courses.size} Courses</span></td>
        <td><span class="badge" style="background-color: var(--color-warning); padding: 2px 6px; border-radius: 8px; color: #000; font-weight: 700;">${semesters.size} Semesters</span></td>
      </tr>
    `;
  }).join('');
}
