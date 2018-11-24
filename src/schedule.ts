//-------------------------- CONSTANTS --------------------------
// Names
const SCHEDULE_SHEET = 'Generated Schedule';
const SURVEY_SHEET = 'Form Responses 1';

// Sheet properties
const STARTING_ROW = 2;
const COLUMNS = ['B','C','D','E','F','G'];

// Schedule properties
const DAYS_OF_THE_WEEK = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday'];
const ALL_SHIFTS = ['9-10','10-11','11-12','12-1','1-2','2-3','3-4','4-5','5-6','6-7','7-8','8-9'];
const MAX_TUTORS = 4;
const DAYS_WITH_INDV_AND_GROUP = 4;
const FRIDAY_HOURS_CELL = 4;
const SUNDAY_HOURS_CELL = 9;

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
  let sheet = spreadsheet.getSheetByName(SCHEDULE_SHEET);
  if (sheet === null) { // oh sheet, it doesn't exist
    // insert after form responses
    sheet = spreadsheet.insertSheet(SCHEDULE_SHEET, 1);
  }
  spreadsheet.setActiveSheet(sheet);
  writeBlankSchedule();
  let tutors = fetchSurveyData();
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

/**
 * Constructs a tutor object for each survey response and returns a list of 
 * all tutor objects in the order his/her response was submitted.
 */
function fetchSurveyData(): Tutor[] {
  let tutors = []; // reset data
  const survey = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SURVEY_SHEET);
  const lastRow = survey.getLastRow();
  const infoRange = survey.getRange('A2:F' + lastRow);
  const basicInfo = infoRange.getValues() as string[][];
  const timeRange = survey.getRange('G2:P' + lastRow);
  const hoursInfo = timeRange.getValues() as string[][];
  if (!basicInfo) {
    Logger.log('No data found.');
  } else {
    for (let row = 0; row < basicInfo.length; row++) {
      let tutor = new Tutor();
      tutor.name = basicInfo[row][1]; // B
      tutor.email = basicInfo[row][2] // C
      tutor.major = basicInfo[row][3]; // D
      tutor.level = basicInfo[row][4]; // E
      tutor.courses = basicInfo[row][5].split(" "); // F
      
      // combine individual and group hours.
      let totalHours: string[] = [];
      // only required for Monday to Thursday.
      for (let day = 0; day < DAYS_WITH_INDV_AND_GROUP; day++) {
        let individual = hoursInfo[row][day];
        // individual and group responses for the same day are 5 cells apart.
        let group = hoursInfo[row][day+5];
        totalHours.push(mergeHours_(individual, group));
      }
      // add the hours for Friday and Sunday as they are.
      totalHours.push(hoursInfo[row][FRIDAY_HOURS_CELL]);
      // push Sunday hours to front
      totalHours.splice(0, 0, hoursInfo[row][SUNDAY_HOURS_CELL]); 
      tutor.setShifts(totalHours);
      tutors.push(tutor);
    }
  }
  return tutors;
}

/**
 * Combines the responses for individual and group hours per day into
 * one single string. 
 */
function mergeHours_(individual: string, group: string): string {
  if (!individual && !group) { // both are empty
    return '';
  } else if (!group) { // only works individual hours
    return individual;
  } else if (!individual) { // only works group hours
    return group;
  }
  return individual + ', ' + group;
}
//-------------------------- General purpose --------------------------
class Tutor {
  name: string
  email: string
  major: string
  level: string   // class level: freshman, sophomore, junior, senior
  courses: string[]
  shifts: WeeklyShifts;

  /**
   * Converts the responses given for hours available per day into an object
   * containing each day of the the week --> an array of shifts the tutor
   * can work for that day. If the tutor cannot work at all for a given day,
   * the value for that day is empty.
   *
   * @param {array} Available hours per work day from a tutor's form response.
   * @return {object} {workday1: [hours], workday2: [hours], ...}
   */
  public setShifts(hoursPerDay: string[]) {
    for (let day = 0; day < hoursPerDay.length; day++) {
      if (!hoursPerDay[day]) { // tutor does not work this day.
        this.shifts[DAYS_OF_THE_WEEK[day]] = [];
      } else {
        this.shifts[DAYS_OF_THE_WEEK[day]] = this.formatHours_(hoursPerDay[day]);
      }
    }
  }

  /**
   * Removes the meridian suffix from each shift duration in the input.
   *
   * @param {string} shifts in the format "9-10 AM, 10-11 AM,..."
   * @return {array} shifts in the format [9-10, 10-11,...]
   */
  private formatHours_(hours: string): string[] {
    let array = hours.split(', ');
    for (let i = 0; i < array.length; i++) {
      let noSuffix = array[i].split(' ')[0];
      array[i] = noSuffix;
    }
    return array;
  }
}

/**
 * Has a property representing each day of the week.
 * Each property is an array of strings, where each element is a shift that can be worked.
 * ex) monday = ['9-10', '3-4']
 */
class WeeklyShifts {
  sunday: string[]
  monday: string[]
  tuesday: string[]
  wednesday: string[]
  thursday: string[]
  friday: string[]
  [key: string]: string[]; // for index signature
}

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
    let column = COLUMNS[i];
    let startRow = STARTING_ROW;
    let endRow = startRow + MAX_TUTORS-1; // subtract one because the group of cells is inclusive
    for (let j = 0; j < ALL_SHIFTS.length; j++) {
      let shift = new ShiftBlock();
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