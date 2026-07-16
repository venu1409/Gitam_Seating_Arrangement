/**
 * Seating Arrangement Algorithm Module
 */

export class SeatingArranger {
  /**
   * @param {Array<Object>} students - Array of filtered student objects
   * @param {Array<Object>} rooms - Array of room objects
   * @param {number} studentsPerBench - 2 or 3 students per bench
   */
  constructor(students, rooms, studentsPerBench = 2) {
    this.students = students;
    this.rooms = rooms;
    this.studentsPerBench = parseInt(studentsPerBench);
  }

  /**
   * Runs the seating plan generation engine.
   * @returns {Object} The generation results containing seatingByRoom, unseatedStudents, and warnings
   */
  generate() {
    if (!this.students || this.students.length === 0) {
      return { seatingByRoom: {}, unseatedStudents: [], warnings: ['No student data to seat.'] };
    }
    if (!this.rooms || this.rooms.length === 0) {
      return { seatingByRoom: {}, unseatedStudents: this.students, warnings: ['No rooms available for seating.'] };
    }

    // 1. Group students by Course Code
    // Within each course group, sub-group by Branch and sort to keep them contiguous
    const courseGroups = {};
    for (const student of this.students) {
      const courseKey = student.courseCode.trim().toUpperCase();
      if (!courseGroups[courseKey]) {
        courseGroups[courseKey] = {
          courseCode: courseKey,
          courseName: student.courseName.trim(),
          studentsByBranch: {}
        };
      }
      
      const branchKey = student.branch.trim().toUpperCase();
      if (!courseGroups[courseKey].studentsByBranch[branchKey]) {
        courseGroups[courseKey].studentsByBranch[branchKey] = [];
      }
      courseGroups[courseKey].studentsByBranch[branchKey].push(student);
    }

    // Convert to flat list of courses with sorted, contiguous branch cohorts
    const courses = [];
    for (const courseCode in courseGroups) {
      const group = courseGroups[courseCode];
      const flatStudents = [];
      
      // Sort branches alphabetically for stable output
      const sortedBranches = Object.keys(group.studentsByBranch).sort();
      for (const branch of sortedBranches) {
        flatStudents.push(...group.studentsByBranch[branch]);
      }
      
      courses.push({
        courseCode: group.courseCode,
        courseName: group.courseName,
        students: flatStudents,
        total: flatStudents.length
      });
    }

    // 2. Partition Course Groups into K tracks (K = studentsPerBench)
    // We sort the courses by total student count descending and distribute greedily using LPT.
    courses.sort((a, b) => b.total - a.total);
    const K = this.studentsPerBench;
    const tracks = Array.from({ length: K }, () => []);
    const trackSizes = Array.from({ length: K }, () => 0);

    for (const course of courses) {
      // Find the track with the smallest current student size
      let minTrackIdx = 0;
      let minSize = trackSizes[0];
      for (let i = 1; i < K; i++) {
        if (trackSizes[i] < minSize) {
          minSize = trackSizes[i];
          minTrackIdx = i;
        }
      }
      
      tracks[minTrackIdx].push(course);
      trackSizes[minTrackIdx] += course.total;
    }

    // 3. Flatten each track into a flat pool of students (preserving course/branch contiguity)
    const trackStudentPools = tracks.map(trackCourses => {
      const pool = [];
      for (const course of trackCourses) {
        pool.push(...course.students);
      }
      return pool;
    });

    // 4. Create a sequential list of all benches across all rooms
    // We order them room-by-room, column-by-column, and bench-by-bench (top-to-bottom)
    const allBenches = [];
    for (const room of this.rooms) {
      const cols = [room.c1, room.c2, room.c3, room.c4, room.c5, room.c6];
      for (let colIdx = 0; colIdx < cols.length; colIdx++) {
        const numBenches = cols[colIdx];
        if (numBenches && numBenches > 0) {
          for (let bIdx = 0; bIdx < numBenches; bIdx++) {
            allBenches.push({
              room: room,
              colIndex: colIdx + 1,
              benchIndex: bIdx + 1,
              seats: Array(K).fill(null)
            });
          }
        }
      }
    }

    // 5. Place students onto benches track by track
    // For each bench, draw from the tracks sequentially.
    // If drawing causes a course-code conflict on the bench, look ahead in the track pool.
    const activeStudentPointers = Array(K).fill(0);

    for (let bIdx = 0; bIdx < allBenches.length; bIdx++) {
      const bench = allBenches[bIdx];
      const seatedCourses = []; // Track courses on this specific bench

      for (let trackIdx = 0; trackIdx < K; trackIdx++) {
        const pool = trackStudentPools[trackIdx];
        
        // Find the first student in the pool that does not conflict with already seated courses
        let foundIdx = -1;
        for (let i = 0; i < pool.length; i++) {
          const student = pool[i];
          if (!seatedCourses.includes(student.courseCode.toUpperCase())) {
            foundIdx = i;
            break;
          }
        }

        if (foundIdx !== -1) {
          const student = pool[foundIdx];
          bench.seats[trackIdx] = student;
          seatedCourses.push(student.courseCode.toUpperCase());
          
          // Remove from track pool as they are seated
          pool.splice(foundIdx, 1);
        } else {
          // Leave blank temporarily if no compatible student is available in this track
          bench.seats[trackIdx] = null;
        }
      }
    }

    // Gather leftover unseated students
    const unseatedStudents = [];
    for (let trackIdx = 0; trackIdx < K; trackIdx++) {
      unseatedStudents.push(...trackStudentPools[trackIdx]);
    }

    // 6. Post-processing: Try to seat leftover students in remaining blank seats
    // Scan all benches for blanks, and see if any unseated student can fit there.
    if (unseatedStudents.length > 0) {
      for (let bIdx = 0; bIdx < allBenches.length; bIdx++) {
        const bench = allBenches[bIdx];
        
        for (let trackIdx = 0; trackIdx < K; trackIdx++) {
          if (bench.seats[trackIdx] === null && unseatedStudents.length > 0) {
            // Find current courses on this bench
            const currentCourses = bench.seats
              .filter(s => s !== null)
              .map(s => s.courseCode.toUpperCase());
              
            const matchIdx = unseatedStudents.findIndex(student => 
              !currentCourses.includes(student.courseCode.toUpperCase())
            );

            if (matchIdx !== -1) {
              bench.seats[trackIdx] = unseatedStudents[matchIdx];
              unseatedStudents.splice(matchIdx, 1);
            }
          }
        }
      }
    }

    // 7. Group the final seated benches back by room for output and visualization
    const seatingByRoom = {};
    for (const room of this.rooms) {
      seatingByRoom[room.roomNumber] = {
        room: room,
        benches: []
      };
    }

    for (const bench of allBenches) {
      const roomNum = bench.room.roomNumber;
      if (seatingByRoom[roomNum]) {
        seatingByRoom[roomNum].benches.push(bench);
      }
    }

    // Generate warning reports if necessary
    const warnings = [];
    if (unseatedStudents.length > 0) {
      warnings.push(`Warning: ${unseatedStudents.length} students could not be seated due to capacity limits or course code conflicts.`);
    }

    return {
      seatingByRoom,
      unseatedStudents,
      warnings
    };
  }
}
