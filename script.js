/**
 * Main Controller Module
 */

import { parseStudentExcel, parseRoomExcel } from './excel.js';
import { SeatingArranger } from './algorithm.js';
import * as ui from './ui.js';
import { exportSeatingWorkbook } from './export.js';

// Application State
const state = {
  students: [],
  rooms: [],
  filteredStudents: [],
  seatingResult: null,
  selectedDate: '',
  selectedSession: '',
  studentsPerBench: 2
};

// DOM Elements
const elements = {
  studentInput: document.getElementById('student-file-input'),
  studentUploader: document.getElementById('student-uploader'),
  studentStatus: document.getElementById('student-file-status'),
  
  roomInput: document.getElementById('room-file-input'),
  roomUploader: document.getElementById('room-uploader'),
  roomStatus: document.getElementById('room-file-status'),
  
  examDateSelect: document.getElementById('exam-date'),
  examSessionSelect: document.getElementById('exam-session'),
  customTimingInput: document.getElementById('custom-timing'),
  
  btnLoadDemo: document.getElementById('btn-load-demo'),
  btnGenerate: document.getElementById('btn-generate'),
  btnExport: document.getElementById('btn-export'),
  btnPreviewTab: document.getElementById('btn-preview-tab'),
  btnClear: document.getElementById('btn-clear'),
  
  engineStatus: document.getElementById('engine-status-text'),
  warningBanner: document.getElementById('warning-banner'),
  closeWarning: document.getElementById('close-warning')
};

// Initialize listeners
document.addEventListener('DOMContentLoaded', () => {
  setupNavigation();
  setupUploaders();
  setupConfigListeners();
  setupActionButtons();
});

// Setup sidebar tab navigation
function setupNavigation() {
  const navItems = document.querySelectorAll('.nav-item');
  const panels = document.querySelectorAll('.tab-panel');
  const pageTitle = document.getElementById('page-title');
  const pageSubtitle = document.getElementById('page-subtitle');
  
  const subtitles = {
    dashboard: 'Overview and exam seating controls',
    students: 'Review list of imported student registry',
    rooms: 'Inspect configured room lists and bench capacities',
    seating: 'Interactive seating arrangements for each classroom',
    reports: 'Export summaries and room occupation lists'
  };

  navItems.forEach(item => {
    item.addEventListener('click', (e) => {
      e.preventDefault();
      const tabId = item.getAttribute('data-tab');
      
      // Update sidebar styling
      navItems.forEach(nav => nav.classList.remove('active'));
      item.classList.add('active');
      
      // Update tab panels visibility
      panels.forEach(panel => panel.classList.remove('active'));
      document.getElementById(`tab-${tabId}`).classList.add('active');
      
      // Update page title
      pageTitle.textContent = item.innerText.trim();
      pageSubtitle.textContent = subtitles[tabId] || '';
    });
  });

  // Enable visual preview transition button
  elements.btnPreviewTab.addEventListener('click', () => {
    const seatingNav = document.querySelector('.nav-item[data-tab="seating"]');
    if (seatingNav) seatingNav.click();
  });
}

// Setup upload interactions
function setupUploaders() {
  // Bind Student Uploader clicks
  elements.studentUploader.addEventListener('click', () => elements.studentInput.click());
  elements.studentInput.addEventListener('change', handleStudentUpload);
  
  // Drag and Drop for Student
  setupDragAndDrop(elements.studentUploader, elements.studentInput, handleStudentUpload);

  // Bind Room Uploader clicks
  elements.roomUploader.addEventListener('click', () => elements.roomInput.click());
  elements.roomInput.addEventListener('change', handleRoomUpload);

  // Drag and Drop for Room
  setupDragAndDrop(elements.roomUploader, elements.roomInput, handleRoomUpload);
}

// Drag & drop utility helper
function setupDragAndDrop(dropArea, fileInput, uploadHandler) {
  ['dragenter', 'dragover'].forEach(eventName => {
    dropArea.addEventListener(eventName, (e) => {
      e.preventDefault();
      dropArea.style.borderColor = 'var(--accent-color)';
    }, false);
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropArea.addEventListener(eventName, (e) => {
      e.preventDefault();
      dropArea.style.borderColor = '';
    }, false);
  });

  dropArea.addEventListener('drop', (e) => {
    const dt = e.dataTransfer;
    const files = dt.files;
    if (files.length > 0) {
      fileInput.files = files;
      // Trigger event manually
      const event = new Event('change');
      fileInput.dispatchEvent(event);
    }
  }, false);
}

// Handle student sheet import
async function handleStudentUpload(e) {
  const files = e.target.files;
  if (!files || files.length === 0) return;

  elements.studentStatus.textContent = 'Loading...';
  elements.studentUploader.classList.remove('success');

  try {
    let allStudents = [];
    for (const file of files) {
      const buffer = await file.arrayBuffer();
      const fileStudents = await parseStudentExcel(buffer);
      allStudents.push(...fileStudents);
    }
    
    // De-duplicate / index sequential students
    state.students = allStudents.map((s, idx) => ({ ...s, id: idx }));
    
    elements.studentStatus.textContent = `${state.students.length} students loaded`;
    elements.studentUploader.classList.add('success');
    
    // Fill Exam Dates select
    const uniqueDates = [...new Set(state.students.map(s => s.date))].filter(Boolean).sort();
    
    elements.examDateSelect.innerHTML = '<option value="">Select Date</option>' + 
      uniqueDates.map(date => `<option value="${date}">${date}</option>`).join('');
    
    // Enable configuration selectors
    elements.examDateSelect.removeAttribute('disabled');
    elements.examSessionSelect.removeAttribute('disabled');
    elements.customTimingInput.removeAttribute('disabled');

    ui.renderStudentsTable(state.students);
    checkEnableGenerate();
    
    elements.engineStatus.textContent = `Imported students from ${files.length} file(s). Now configure exam timings.`;
  } catch (err) {
    elements.studentStatus.textContent = 'Failed to load';
    ui.showWarning(err.message);
  }
}

// Handle room list import
async function handleRoomUpload(e) {
  const file = e.target.files[0];
  if (!file) return;

  elements.roomStatus.textContent = 'Loading...';
  elements.roomUploader.classList.remove('success');

  try {
    const buffer = await file.arrayBuffer();
    state.rooms = await parseRoomExcel(buffer);
    
    elements.roomStatus.textContent = `${state.rooms.length} rooms loaded`;
    elements.roomUploader.classList.add('success');
    
    ui.renderRoomsTable(state.rooms);
    checkEnableGenerate();
    
    elements.engineStatus.textContent = 'Room list imported. Ready to calculate.';
  } catch (err) {
    elements.roomStatus.textContent = 'Failed to load';
    ui.showWarning(err.message);
  }
}

// Handle Exam configuration adjustments
function setupConfigListeners() {
  // Timing defaults based on sessions
  elements.examSessionSelect.addEventListener('change', (e) => {
    state.selectedSession = e.target.value;
    if (state.selectedSession === 'FN') {
      elements.customTimingInput.value = '10:15 AM - 12:15 PM';
    } else if (state.selectedSession === 'AN') {
      elements.customTimingInput.value = '02:00 PM - 04:00 PM';
    } else {
      elements.customTimingInput.value = '';
    }
    filterStudentsBySelection();
  });

  elements.examDateSelect.addEventListener('change', (e) => {
    state.selectedDate = e.target.value;
    filterStudentsBySelection();
  });

  // Toggle students per bench
  const radioButtons = document.querySelectorAll('input[name="students-per-bench"]');
  radioButtons.forEach(radio => {
    radio.addEventListener('change', (e) => {
      state.studentsPerBench = parseInt(e.target.value);
    });
  });

  // Warning banner close action
  elements.closeWarning.addEventListener('click', () => {
    ui.hideWarning();
  });
}

// Filter students in real time as Date/Session selection changes
function filterStudentsBySelection() {
  if (!state.selectedDate || !state.selectedSession) {
    state.filteredStudents = [];
    ui.updateStats({
      totalStudents: 0,
      totalSubjects: 0,
      roomsUsed: 0,
      totalCapacity: 0,
      emptySeats: 0
    });
    checkEnableGenerate();
    return;
  }

  state.filteredStudents = state.students.filter(student => 
    student.date === state.selectedDate && 
    student.session === state.selectedSession
  );

  const uniqueSubjects = new Set(state.filteredStudents.map(s => s.courseCode));

  // Update statistics for selection overview
  ui.updateStats({
    totalStudents: state.filteredStudents.length,
    totalSubjects: uniqueSubjects.size,
    roomsUsed: 0,
    totalCapacity: 0,
    emptySeats: 0
  });

  elements.engineStatus.textContent = `Filtered ${state.filteredStudents.length} students across ${uniqueSubjects.size} subjects for selected exam slot.`;
  checkEnableGenerate();
}

// Check whether Seating engine is ready to run
function checkEnableGenerate() {
  if (state.filteredStudents.length > 0 && state.rooms.length > 0) {
    elements.btnGenerate.removeAttribute('disabled');
  } else {
    elements.btnGenerate.setAttribute('disabled', 'true');
  }
}

// Action button handlers
function setupActionButtons() {
  // Load Demo Data Action
  elements.btnLoadDemo.addEventListener('click', async () => {
    try {
      elements.engineStatus.textContent = 'Fetching and parsing template workbook...';
      const response = await fetch('Seating Plan Master copy 3-AN-S1.xlsx');
      if (!response.ok) {
        throw new Error('Failed to fetch Seating Plan Master copy 3-AN-S1.xlsx from server.');
      }
      const buffer = await response.arrayBuffer();
      
      // Parse rooms from Sheet 1
      state.rooms = await parseRoomExcel(buffer);
      elements.roomStatus.textContent = `${state.rooms.length} rooms loaded (Demo)`;
      elements.roomUploader.classList.add('success');
      ui.renderRoomsTable(state.rooms);
      
      // Parse students from Sheet 2 (S1)
      state.students = await parseStudentExcel(buffer);
      elements.studentStatus.textContent = `${state.students.length} students loaded (Demo)`;
      elements.studentUploader.classList.add('success');
      ui.renderStudentsTable(state.students);
      
      // Fill Exam Dates select
      const uniqueDates = [...new Set(state.students.map(s => s.date))].filter(Boolean).sort();
      elements.examDateSelect.innerHTML = '<option value="">Select Date</option>' + 
        uniqueDates.map(date => `<option value="${date}">${date}</option>`).join('');
        
      elements.examDateSelect.removeAttribute('disabled');
      elements.examSessionSelect.removeAttribute('disabled');
      elements.customTimingInput.removeAttribute('disabled');
      
      // Select first date and session by default for demo
      if (uniqueDates.length > 0) {
        elements.examDateSelect.value = uniqueDates[0];
        state.selectedDate = uniqueDates[0];
      }
      elements.examSessionSelect.value = 'FN';
      state.selectedSession = 'FN';
      elements.customTimingInput.value = '10:15 AM - 12:15 PM';
      
      filterStudentsBySelection();
      checkEnableGenerate();
      
      elements.engineStatus.textContent = 'Demo data successfully loaded! You can now click "Generate Seating Plan".';
    } catch (err) {
      ui.showWarning('Failed to load demo data: ' + err.message);
    }
  });

  // Generate Seating Action
  elements.btnGenerate.addEventListener('click', () => {
    try {
      elements.engineStatus.textContent = 'Seating allocation engine running...';
      const start = performance.now();
      
      const arranger = new SeatingArranger(
        state.filteredStudents,
        state.rooms,
        state.studentsPerBench
      );
      
      state.seatingResult = arranger.generate();
      
      const duration = ((performance.now() - start) / 1000).toFixed(2);
      
      // Handle warnings
      if (state.seatingResult.warnings.length > 0) {
        ui.showWarning(state.seatingResult.warnings.join('\n'));
      } else {
        ui.hideWarning();
      }

      // Compute statistics for the generated seating
      const usedRooms = Object.keys(state.seatingResult.seatingByRoom).filter(k => 
        state.seatingResult.seatingByRoom[k].benches.some(b => b.seats.some(s => s !== null))
      );
      
      let totalBenchesUsed = 0;
      let totalStudentsSeated = 0;
      
      usedRooms.forEach(roomNum => {
        const roomData = state.seatingResult.seatingByRoom[roomNum];
        totalBenchesUsed += roomData.benches.length;
        totalStudentsSeated += roomData.benches.reduce((sum, b) => sum + b.seats.filter(s => s !== null).length, 0);
      });
      
      const totalCapacity = totalBenchesUsed * state.studentsPerBench;
      const emptySeats = totalCapacity - totalStudentsSeated;
      const uniqueSubjects = new Set(state.filteredStudents.map(s => s.courseCode));

      ui.updateStats({
        totalStudents: totalStudentsSeated,
        totalSubjects: uniqueSubjects.size,
        roomsUsed: usedRooms.length,
        totalCapacity: totalCapacity,
        emptySeats: emptySeats
      });

      // Render Visual layout previews and reports
      ui.renderSeatingCards(state.seatingResult.seatingByRoom, state.studentsPerBench);
      ui.renderReports(state.seatingResult.seatingByRoom, state.studentsPerBench);
      
      // Enable export controls
      elements.btnExport.removeAttribute('disabled');
      elements.btnPreviewTab.removeAttribute('disabled');
      
      elements.engineStatus.textContent = `Seating plan generated in ${duration} seconds. ${totalStudentsSeated} students seated across ${usedRooms.length} rooms.`;
    } catch (err) {
      ui.showWarning('Seating Arrangement engine failed: ' + err.message);
    }
  });

  // Export Excel workbook Action
  elements.btnExport.addEventListener('click', async () => {
    if (!state.seatingResult) return;
    
    elements.engineStatus.textContent = 'Preparing Excel download packages...';
    try {
      const customTiming = elements.customTimingInput.value.trim() || '10:15 AM - 12:15 PM';
      await exportSeatingWorkbook(
        state.seatingResult.seatingByRoom,
        state.selectedDate,
        state.selectedSession,
        customTiming,
        state.studentsPerBench
      );
      elements.engineStatus.textContent = 'Excel workbook exported successfully!';
    } catch (err) {
      ui.showWarning(err.message);
    }
  });

  // Clear data Action
  elements.btnClear.addEventListener('click', () => {
    // Reset state variables
    state.students = [];
    state.rooms = [];
    state.filteredStudents = [];
    state.seatingResult = null;
    state.selectedDate = '';
    state.selectedSession = '';
    state.studentsPerBench = 2;

    // Reset inputs
    elements.studentInput.value = '';
    elements.roomInput.value = '';
    elements.studentStatus.textContent = 'No file chosen';
    elements.roomStatus.textContent = 'No file chosen';
    elements.studentUploader.classList.remove('success');
    elements.roomUploader.classList.remove('success');
    
    // Reset configuration
    elements.examDateSelect.innerHTML = '<option value="">Select Date (Import Student List First)</option>';
    elements.examDateSelect.setAttribute('disabled', 'true');
    elements.examSessionSelect.value = '';
    elements.examSessionSelect.setAttribute('disabled', 'true');
    elements.customTimingInput.value = '';
    elements.customTimingInput.setAttribute('disabled', 'true');
    document.querySelector('input[name="students-per-bench"][value="2"]').checked = true;

    // Disable triggers
    elements.btnGenerate.setAttribute('disabled', 'true');
    elements.btnExport.setAttribute('disabled', 'true');
    elements.btnPreviewTab.setAttribute('disabled', 'true');

    // Clear views
    ui.hideWarning();
    ui.updateStats({
      totalStudents: 0,
      totalSubjects: 0,
      roomsUsed: 0,
      totalCapacity: 0,
      emptySeats: 0
    });
    ui.renderStudentsTable([]);
    ui.renderRoomsTable([]);
    ui.renderSeatingCards({}, 2);
    ui.renderReports({}, 2);
    
    // Go to dashboard tab
    const dashboardNav = document.querySelector('.nav-item[data-tab="dashboard"]');
    if (dashboardNav) dashboardNav.click();

    elements.engineStatus.textContent = 'Workspace cleared. Upload lists to begin.';
  });
}
