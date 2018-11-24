//-------------------------- CONSTANTS --------------------------
// Names
const SCHEDULE_SHEET = 'Generated Schedule';

// Sheet properties
const STARTING_ROW = 2;
const COLUMNS = ['B','C','D','E','F','G'];

// Schedule properties
const DAYS_OF_THE_WEEK = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday'];
const ALL_SHIFTS = ['9-10','10-11','11-12','12-1','1-2','2-3','3-4','4-5','5-6','6-7','7-8','8-9'];
const MAX_TUTORS = 4;

//-------------------------- UI-related --------------------------
/**
 * Creates a menu entry in the Google Sheets UI when the document is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Schedule Helper')
    .addItem('Generate schedule', 'generateSchedule')
    .addToUi();
}

//-------------------------- Schedule automation --------------------------
/**
 * Creates a new schedule draft from form responses given in SURVEY_SHEET
 * and prints result in SCHEDULE_SHEET. 
 * Acts as the main schedule automation function and called when user selects
 * 'Generate Schedule' option in add-on menu.
 */
function generateSchedule() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  try {
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName(SCHEDULE_SHEET));
  } catch(e) {
    // insert after form responses
    spreadsheet.insertSheet(SCHEDULE_SHEET, 1);
  }
  writeBlankSchedule();
}

/** 
 * Pre-formats a blank schedule with the days of the week, shift hours, and 
 * an empty grid. Also populates the array for all shift ranges.
 */
function writeBlankSchedule(): void {
  let sheet = SpreadsheetApp.getActiveSheet();
  sheet.clear();
  // top row
  sheet.getRange('B1:G1')
    .setValues([['Sunday','Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']])
    .setFontWeight('bold')
  // left column
  let leftColumn = 'A';
  let shiftRow = STARTING_ROW;
  ALL_SHIFTS.forEach((shift) => {
    sheet.getRange(leftColumn+shiftRow)
      .setValue(shift)
      .setFontWeight('bold')
      .setHorizontalAlignment('right');
    shiftRow += MAX_TUTORS;  
  });
  // per shift per day
  let shiftBlocks = getAllShiftRanges();
  shiftBlocks.forEach((block) => {
    sheet.getRange(block.range)
      .setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  });
}

//-------------------------- General purpose --------------------------
class ShiftBlock {
  range: string
  day: string
  time: string
}

/**
 * Returns an array containing ranges in A1 format, each of which represent 
 * a shift block for the given schedule.
 */
function getAllShiftRanges(): ShiftBlock[] {
  let allShiftRanges = [];
  for (let i = 0; i < DAYS_OF_THE_WEEK.length; i++) {
    var column = COLUMNS[i];
    var startRow = STARTING_ROW;
    var endRow = startRow + MAX_TUTORS-1; // subtract one because the group of cells is inclusive
    for (var j = 0; j < ALL_SHIFTS.length; j++) {
      var shift = new ShiftBlock();
      shift.range = column+startRow + ':' + column+endRow;
      shift.day = DAYS_OF_THE_WEEK[i];
      shift.time = ALL_SHIFTS[j];
      // store the range of the current shift block
      allShiftRanges.push(shift);
      startRow += MAX_TUTORS;
      endRow += MAX_TUTORS;
    }
  }
  return allShiftRanges;
}