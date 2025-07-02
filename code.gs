/**
 *
 * @OnlyCurrentDoc
 * 
 **/


// --- Teams Configuration for WEDNESDAY --- //
const TEAMS_WEDNESDAY = {
  'Equipe A': ['member 1', 'member 2', 'member 3', 'member 4'],
  'Equipe B': ['member 5', 'member 6', 'member 7', 'member 8', 'member 9'],
  'Equipe C': ['member 10', 'member 11', 'member 12', 'member 13'],
  'Equipe D': ['member 14', 'member 15', 'member 16', 'member 17'],
  'Equipe E': ['member 18', 'member 19', 'member 20', 'member 21', 'member 22']
};

// --- Teams Configuration for FRIDAY --- //
const TEAMS_FRIDAY = {
  'Equipe A': ['member 1', 'member 2', 'member 3', 'member 4', 'member 22'],
  'Equipe B': ['member 5', 'member 6', 'member 7', 'member 8', 'member 9', 'member 20'],
  'Equipe C': ['member 10', 'member 11', 'member 12', 'member 13', 'member 19', 'member 21'],
  'Equipe D': ['member 14', 'member 15', 'member 16', 'member 17', 'member 18']
};

// --- Responsibilities --- //
const RESPONSIBILITIES = [
  'Organisation des tables',
  'Organisation des chaises',
  'Vaisselle assiettes et cuillères',
  'Débarassage de table'
];

// --- Excluded Members --- //
const EXCLUDED_MEMBERS = ['abdellah', 'SAAD'];

// --- Automation Settings --- //
const DAY_OF_MONTH_TO_CREATE_SHEET = 20;

// --- Sheet style --- //
const STYLES = {
  WEEK_HEADER: {
    BACKGROUND: '#000000', // Darkish Blue
    FONT_COLOR: '#ffffff', // White
    FONT_SIZE: 10,
  },
  DATE_HEADER: {
    BACKGROUND: '#21124c',
    FONT_COLOR: '#ffffff',
    FONT_SIZE: 10,
  },
  TEAM_HEADER: {
    BACKGROUND: '#1c4587',
    FONT_COLOR: '#ffffff',
    FONT_SIZE: 10,
  },
  RESPONSIBILITIES_COLUMN: {
    BACKGROUND: '#c9dbf9',
    FONT_COLOR: '#000000',
    FONT_SIZE: 10,
  },
  TABLE_BORDER_COLOR: '#000000', // Gray
};

const COLUMN_PADDING = 15;

const FRENCH_MONTHS = {
  'janvier': 0, 'février': 1, 'mars': 2, 'avril': 3, 'mai': 4, 'juin': 5,
  'juillet': 6, 'août': 7, 'septembre': 8, 'octobre': 9, 'novembre': 10, 'décembre': 11
};

// ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ //
// ~~~~~~ SCRIPT LOGIC ~~~~~~~~~~ //
// ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ //


/**
 * Creates a menu in the spreadsheet UI to manually trigger the script.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Organisation déjeuner')
    .addItem('Créer les plannings du mois prochain', 'createScheduleSheetsForNextMonth')
    .addToUi();
}

/**
 * Sets up a trigger to automatically run the script to create the new sheets.
 */
function setupTrigger() {
  const allTriggers = ScriptApp.getProjectTriggers();
  for (const trigger of allTriggers) {
    if (trigger.getHandlerFunction() === 'createScheduleSheetsForNextMonth') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  ScriptApp.newTrigger('createScheduleSheetsForNextMonth')
    .timeBased()
    .onMonthDay(DAY_OF_MONTH_TO_CREATE_SHEET)
    .create();
}

/**
 * Main function that creates two separate sheets for the next month,
 * one for Wednesdays and one for Fridays.
 */
function createScheduleSheetsForNextMonth() {
  // Create Wednesday's sheet
  createSheetForDay('mercredi', 3, TEAMS_WEDNESDAY);
  // Create Friday's sheet
  createSheetForDay('vendredi', 5, TEAMS_FRIDAY);
  // Clean up the oldest month's sheets
  cleanupOldSheets();
}


/**
 * Creates a new sheet for a specific day of the week with its corresponding teams.
 * @param {string} dayName - The name of the day (e.g., "Wednesday").
 * @param {number} dayOfWeek - The day number (Sunday=0, Monday=1, ..., Saturday=6).
 * @param {Object} teams - The team configuration object to use for this day.
 */
function createSheetForDay(dayName, dayOfWeek, teams) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const nextMonth = new Date();
  nextMonth.setMonth(nextMonth.getMonth() + 1);
  const monthName = Utilities.formatDate(nextMonth, Session.getScriptTimeZone(), 'MMMM yyyy');
  var monthNameFr = LanguageApp.translate(monthName, 'en', 'fr');
  const sheetName = `${monthNameFr} (${dayName})`;

  if (spreadsheet.getSheetByName(sheetName)) {
    console.log(`Sheet "${sheetName}" already exists. Skipping creation.`);
    return;
  }

  const newSheet = spreadsheet.insertSheet(sheetName);
  if (dayName === 'mercredi') {
    newSheet.setTabColor('#4a86e8');
  } else {
    newSheet.setTabColor('#6aa84f');
  }

  const lunchDays = getDaysForMonth(nextMonth.getFullYear(), nextMonth.getMonth(), dayOfWeek);

  if (lunchDays.length === 0) {
    newSheet.getRange(1, 1).setValue(`No ${dayName}s found for ${monthName}.`);
    return;
  }
  
  let currentRow = 1;
  let responsibilityCounter = 0;
  const maxTeamSize = Math.max(...Object.values(teams).map(team => team.length));
  const numColumns = 1 + maxTeamSize;

  for (const day of lunchDays) {
    const formattedDate = Utilities.formatDate(day, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    const weekNumber = Math.ceil(day.getDate() / 7);

    const weekHeaderRange = newSheet.getRange(currentRow, 1, 1, numColumns);
    weekHeaderRange.setValue(`Semaine ${weekNumber}`)
      .merge()
      .setBackground(STYLES.WEEK_HEADER.BACKGROUND)
      .setFontColor(STYLES.WEEK_HEADER.FONT_COLOR)
      .setFontSize(STYLES.WEEK_HEADER.FONT_SIZE)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

    currentRow++;

    for (const teamName in teams) {
      const teamMembers = teams[teamName];
      const eligibleMembers = teamMembers.filter(member => !EXCLUDED_MEMBERS.includes(member));
      const teamHeaderRow = currentRow;

      // --- Write Team and Responsibility Data ---
      newSheet.getRange(teamHeaderRow, 1).setValue(teamName);

      if (teamMembers.length > 0) {
        newSheet.getRange(teamHeaderRow, 2, 1, teamMembers.length).setValues([teamMembers]);
      }
      const teamHeaderRange = newSheet.getRange(teamHeaderRow, 1, 1,numColumns);
      teamHeaderRange.setBorder(true, true, true, true, true, true, STYLES.TABLE_BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
      currentRow+= 1;

        // --- Write Date Header ---
      const dateHeaderRange = newSheet.getRange(currentRow, 1, 1, numColumns);
      dateHeaderRange.setValue(formattedDate)
        .merge()
        .setBackground(STYLES.DATE_HEADER.BACKGROUND)
        .setFontColor(STYLES.DATE_HEADER.FONT_COLOR)
        .setFontSize(STYLES.DATE_HEADER.FONT_SIZE)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      currentRow+= 1;
      
      // for (let i = 0; i < RESPONSIBILITIES.length; i++) {
      //   const responsibility = RESPONSIBILITIES[i];
      //   const row = currentRow + i;
      //   newSheet.getRange(row, 1).setValue(responsibility)
      //   .merge()
      //   .setBackground(STYLES.RESPONSIBILITIES_COLUMN.BACKGROUND)
      //   .setFontColor(STYLES.RESPONSIBILITIES_COLUMN.FONT_COLOR)
      //   .setFontSize(STYLES.RESPONSIBILITIES_COLUMN.FONT_SIZE)
      //   .setFontWeight('bold');

      //   const assignments = new Array(teamMembers.length).fill(false);

      //   if (eligibleMembers.length > 0) {
      //     const responsibleMember = eligibleMembers[responsibilityCounter % eligibleMembers.length];
      //     const responsibleMemberIndex = teamMembers.indexOf(responsibleMember);
      //     if (responsibleMemberIndex !== -1) {
      //       assignments[responsibleMemberIndex] = true;
      //     }
      //     responsibilityCounter++;
      //   }
      //   if (assignments.length > 0) {
      //     const assignmentRange = newSheet.getRange(row, 2, 1, teamMembers.length);
      //     // Create a checkbox data validation rule.
      //     const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
      //     // Apply the checkbox rule and set the checked/unchecked state.
      //     assignmentRange.setDataValidation(rule).setValues([assignments]);
      //   }
      // }

      let i = 0;
      while (i < RESPONSIBILITIES.length) {
        const responsibility = RESPONSIBILITIES[i];
        const firstResponsibilityRow = currentRow
        // Determine who is responsible for the next task(s)
        const responsibleMember = eligibleMembers.length > 0 ? eligibleMembers[responsibilityCounter % eligibleMembers.length] : null;
        const responsibleMemberIndex = responsibleMember ? teamMembers.indexOf(responsibleMember) : -1;

        // Helper function to set checkboxes for a given row
        const setCheckboxes = (row, responsibleIdx) => {
          const assignments = new Array(teamMembers.length).fill(false);
          if (responsibleIdx !== -1) {
            assignments[responsibleIdx] = true;
          }
          if (teamMembers.length > 0) {
            const assignmentRange = newSheet.getRange(row, 2, 1, teamMembers.length);
            const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
            assignmentRange.setDataValidation(rule).setValues([assignments]);
          }
        };
        
        // Helper function to style the responsibility cell
        const styleResponsibilityCell = (row, text) => {
          newSheet.getRange(row, 1).setValue(text)
            .setBackground(STYLES.RESPONSIBILITIES_COLUMN.BACKGROUND)
            .setFontColor(STYLES.RESPONSIBILITIES_COLUMN.FONT_COLOR)
            .setFontSize(STYLES.RESPONSIBILITIES_COLUMN.FONT_SIZE)
            .setFontWeight('bold');
        };

        // Check for the combined task
        if (responsibility === 'Organisation des tables' && i + 1 < RESPONSIBILITIES.length && RESPONSIBILITIES[i+1] === 'Organisation des chaises') {
          // Assign both tasks to the same person
          styleResponsibilityCell(firstResponsibilityRow + i, responsibility);
          setCheckboxes(firstResponsibilityRow + i, responsibleMemberIndex);
          i++; // Move to the next responsibility
          
          styleResponsibilityCell(firstResponsibilityRow + i, RESPONSIBILITIES[i]);
          setCheckboxes(firstResponsibilityRow + i, responsibleMemberIndex); // Assign to the same person
          i++; // Move past the second task
          
          responsibilityCounter++; // Increment the counter only ONCE for the pair
        } else {
          // Assign as a single, individual task
          styleResponsibilityCell(firstResponsibilityRow + i, responsibility);
          setCheckboxes(firstResponsibilityRow + i, responsibleMemberIndex);
          i++; // Move to the next task
          
          responsibilityCounter++; // Increment the counter for this single task
        }
      }

      // --- Apply Styles to the Team Block ---
      const teamBlockNumRows = RESPONSIBILITIES.length;
      const teamBlockRange = newSheet.getRange(currentRow, 1, teamBlockNumRows, numColumns);
      
      // Team Header Row Style
      newSheet.getRange(teamHeaderRow, 1, 1, numColumns)
        .setBackground(STYLES.TEAM_HEADER.BACKGROUND)
        .setFontColor(STYLES.TEAM_HEADER.FONT_COLOR)
        .setFontSize(STYLES.TEAM_HEADER.FONT_SIZE)
        .setFontWeight('bold');
      
      // Border for the whole block
      teamBlockRange.setBorder(true, true, true, true, true, true, STYLES.TABLE_BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);

      currentRow += teamBlockNumRows
    }
  }
  newSheet.autoResizeColumns(1, numColumns);
  let maxWidth = 0;
  if (numColumns > 1) {
    for (let i = 2; i <= numColumns; i++) {
      const width = newSheet.getColumnWidth(i);
      if (width > maxWidth) {
        maxWidth = width;
      }
    }
    
    newSheet.setColumnWidths(2, maxTeamSize, maxWidth + COLUMN_PADDING);
  }
}

function cleanupOldSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = spreadsheet.getSheets();
  const scheduleSheetsByMonth = {};

  // Regex to capture Month and Year from sheet names like "June 2025 (Wednesday)"
  const sheetNameRegex = /([a-zA-Zà-üÀ-Ü]+)\s(\d{4})\s\((mercredi|vendredi)\)/;

  for (const sheet of allSheets) {
    const sheetName = sheet.getName();
    const match = sheetName.match(sheetNameRegex);
    
    if (match) {
      const monthName = match[1].toLowerCase();
      const year = match[2];
      
      if (monthName in FRENCH_MONTHS) {
        const monthIndex = FRENCH_MONTHS[monthName];
        const monthDate = new Date(year, monthIndex, 1);
        const dateKey = monthDate.toISOString();

        if (!scheduleSheetsByMonth[dateKey]) {
          scheduleSheetsByMonth[dateKey] = [];
        }
        scheduleSheetsByMonth[dateKey].push(sheet);
      }
    }
  }

  // Get all unique month keys and sort them chronologically (oldest first).
  const sortedMonthKeys = Object.keys(scheduleSheetsByMonth).sort();

  // If we have schedules for more than 2 months, delete the oldest one(s).
  while (sortedMonthKeys.length > 2) {
    const oldestMonthKey = sortedMonthKeys.shift(); // Get and remove the oldest month's key
    const sheetsToDelete = scheduleSheetsByMonth[oldestMonthKey];
    
    for (const sheet of sheetsToDelete) {
      try {
        console.log(`Deleting old sheet: ${sheet.getName()}`);
        spreadsheet.deleteSheet(sheet);
      } catch (e) {
        console.error(`Could not delete sheet: ${sheet.getName()}. Error: ${e.toString()}`);
      }
    }
  }
}

/**
 * Helper function to get all instances of a specific weekday for a given month and year.
 * @param {number} year The year.
 * @param {number} month The month (0-indexed, e.g., 0 for January).
 * @param {number} dayOfWeek The day of the week to find (Sunday=0, ... , Saturday=6).
 * @returns {Date[]} An array of Date objects.
 */
function getDaysForMonth(year, month, dayOfWeek) {
  const days = [];
  const date = new Date(year, month, 1);
  while (date.getMonth() === month) {
    if (date.getDay() === dayOfWeek) {
      days.push(new Date(date));
    }
    date.setDate(date.getDate() + 1);
  }
  return days;
}
