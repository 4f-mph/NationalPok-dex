function onEdit(e) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const s = e.source.getActiveSheet();
  const s0 = spreadsheet.getSheetByName("Main");
  const s1 = spreadsheet.getSheetByName("Options");
  const s2 = spreadsheet.getSheetByName("Unown");

  if (s.getName() === "Main") {
    let cell = e.range;
function onEdit(e) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const s = e.source.getActiveSheet();
  const s0 = spreadsheet.getSheetByName("Main");
  const s1 = spreadsheet.getSheetByName("Options");
  const s2 = spreadsheet.getSheetByName("Unown");

  if (s.getName() === "Main") {
    let cell = e.range;

    if (cell.getColumn() === 2 && cell.getRow() >= 5) {
      let cell2 = cell.offset(0, 1);
      let cell3 = cell.offset(0, 2);

      if (cell.getValue()) {
        cell.setValue(false); // uncheck the checkbox

        let outputString = collectData(s, s1);

        logEntry(cell3, outputString);

        if (cell.getRow() === 205) {
          handleUnownCheckbox(s, s1, s2, outputString, cell2);
        }
      }
    } else if (cell.getColumn() === 1 && cell.getRow() === 4) {
      s1.getRange(8, 2).setValue(cell.getValue()); // update "Current Game" in Options
    }
  } else if (s.getName() === "Options") {
    let cell = e.range;
    if (cell.getColumn() === 2 && cell.getRow() === 8) {
      s0.getRange(4, 1).setValue(cell.getValue()); // update "Current Game" in Main
    }
  }
}

function collectData(s, s1) {
  const date = new Date();
  const time_zone = "PST";
  const parsed_date = Utilities.formatDate(date, time_zone, 'MM/dd/yyyy');
  const game = s1.getRange(8, 2).getValue();
  let outputString = "Captured " + parsed_date + " in " + game;

  // Handle other conditions dynamically
  const ranges = [
    { name: "Alolan", row: 2 },
    { name: "Galarian", row: 3 },
    { name: "Hisuian", row: 4 },
    { name: "Paldean", row: 5 },
    { name: "Other", row: 6 }
    // Add more ranges as needed
  ];

  for (let range of ranges) {
    let value = s1.getRange(range.row, 2).getValue();
    if (value) {
      outputString += " (" + range.name + ")";
      break;
    }
  }

  // Handle alternate forms
  const alternateForms = [
    { startRow: 18, endRow: 31, column: 2 },
    { startRow: 33, endRow: 35, column: 2 },
    // Add more alternate forms ranges as needed
  ];

  for (let form of alternateForms) {
    for (let row = form.startRow; row <= form.endRow; row++) {
      let value = s1.getRange(row, form.column).getValue();
      if (value) {
        outputString += " (" + value + ")";
        break;
      }
    }
  }

  // Handle shiny
  if (s1.getRange(14, 2).getValue()) {
    outputString += " (Shiny☆)";
  }

  return outputString;
}

function handleUnownCheckbox(s, s1, s2, outputString, cell2) {
  const shiny = s1.getRange(14, 2).getValue();
  const letter = s1.getRange(6, 2).getValue().toString().charCodeAt(0);
  let unownCell = s2.getRange(letter - 61, 2);

  if (shiny) {
    unownCell.offset(0, 1).setValue("Caught☆");
  } else if (cell2.getValue() !== "Caught☆") {
    unownCell.offset(0, 1).setValue("Caught");
  }

  logEntry(unownCell.offset(0, 2), outputString);
}

function logEntry(cell, outputString) {
  let recorded = false;
  while (!recorded) {
    if (cell.getValue() === "") {
      if (cell.offset(0, -1).getValue() === outputString) {
        recorded = true;
      } else {
        cell.setValue(outputString);
        recorded = true;
      }
    } else {
      cell = cell.offset(0, 1);
    }
  }
}

    if (cell.getColumn() === 2 && cell.getRow() >= 5) {
      let cell2 = cell.offset(0, 1);
      let cell3 = cell.offset(0, 2);

      if (cell.getValue()) {
        cell.setValue(false); // uncheck the checkbox

        let outputString = collectData(s, s1);

        logEntry(cell3, outputString);

        if (cell.getRow() === 205) {
          handleUnownCheckbox(s, s1, s2, outputString, cell2);
        }
      }
    } else if (cell.getColumn() === 1 && cell.getRow() === 4) {
      s1.getRange(8, 2).setValue(cell.getValue()); // update "Current Game" in Options
    }
  } else if (s.getName() === "Options") {
    let cell = e.range;
    if (cell.getColumn() === 2 && cell.getRow() === 8) {
      s0.getRange(4, 1).setValue(cell.getValue()); // update "Current Game" in Main
    }
  }
}

function collectData(s, s1) {
  const date = new Date();
  const time_zone = "PST";
  const parsed_date = Utilities.formatDate(date, time_zone, 'MM/dd/yyyy');
  const game = s1.getRange(8, 2).getValue();
  let outputString = "Captured " + parsed_date + " in " + game;

  // Handle other conditions dynamically
  const ranges = [
    { name: "Alolan", row: 2 },
    { name: "Galarian", row: 3 },
    { name: "Hisuian", row: 4 },
    { name: "Paldean", row: 5 },
    { name: "Other", row: 6 }
    // Add more ranges as needed
  ];

  for (let range of ranges) {
    let value = s1.getRange(range.row, 2).getValue();
    if (value) {
      outputString += " (" + range.name + ")";
      break;
    }
  }

  // Handle alternate forms
  const alternateForms = [function onEdit(e) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const s = e.source.getActiveSheet();
  const s0 = spreadsheet.getSheetByName("Main");
  const s1 = spreadsheet.getSheetByName("Options");
  const s2 = spreadsheet.getSheetByName("Unown");

  if (s.getName() === "Main") {
    let cell = e.range;

    if (cell.getColumn() === 2 && cell.getRow() >= 5) {
      let cell2 = cell.offset(0, 1);
      let cell3 = cell.offset(0, 2);

      if (cell.getValue()) {
        cell.setValue(false); // uncheck the checkbox

        let outputString = collectData(s, s1);

        logEntry(cell3, outputString);

        if (cell.getRow() === 205) {
          handleUnownCheckbox(s, s1, s2, outputString, cell2);
        }
      }
    } else if (cell.getColumn() === 1 && cell.getRow() === 4) {
      s1.getRange(8, 2).setValue(cell.getValue()); // update "Current Game" in Options
    }
  } else if (s.getName() === "Options") {
    let cell = e.range;
    if (cell.getColumn() === 2 && cell.getRow() === 8) {
      s0.getRange(4, 1).setValue(cell.getValue()); // update "Current Game" in Main
    }
  }
}

function collectData(s, s1) {
  const date = new Date();
  const time_zone = "PST";
  const parsed_date = Utilities.formatDate(date, time_zone, 'MM/dd/yyyy');
  const game = s1.getRange(8, 2).getValue();
  let outputString = "Captured " + parsed_date + " in " + game;

  // Handle other conditions dynamically
  const ranges = [
    { name: "Alolan", row: 2 },
    { name: "Galarian", row: 3 },
    { name: "Hisuian", row: 4 },
    { name: "Paldean", row: 5 },
    { name: "Other", row: 6 }
    // Add more ranges as needed
  ];

  for (let range of ranges) {
    let value = s1.getRange(range.row, 2).getValue();
    if (value) {
      outputString += " (" + range.name + ")";
      break;
    }
  }

  // Handle alternate forms
  const alternateForms = [
    { startRow: 18, endRow: 31, column: 2 },
    { startRow: 33, endRow: 35, column: 2 },
    // Add more alternate forms ranges as needed
  ];

  for (let form of alternateForms) {
    for (let row = form.startRow; row <= form.endRow; row++) {
      let value = s1.getRange(row, form.column).getValue();
      if (value) {
        outputString += " (" + value + ")";
        break;
      }
    }
  }

  // Handle shiny
  if (s1.getRange(14, 2).getValue()) {
    outputString += " (Shiny☆)";
  }

  return outputString;
}

function handleUnownCheckbox(s, s1, s2, outputString, cell2) {
  const shiny = s1.getRange(14, 2).getValue();
  const letter = s1.getRange(6, 2).getValue().toString().charCodeAt(0);
  let unownCell = s2.getRange(letter - 61, 2);

  if (shiny) {
    unownCell.offset(0, 1).setValue("Caught☆");
  } else if (cell2.getValue() !== "Caught☆") {
    unownCell.offset(0, 1).setValue("Caught");
  }

  logEntry(unownCell.offset(0, 2), outputString);
}

function logEntry(cell, outputString) {
  let recorded = false;
  while (!recorded) {
    if (cell.getValue() === "") {
      if (cell.offset(0, -1).getValue() === outputString) {
        recorded = true;
      } else {
        cell.setValue(outputString);
        recorded = true;
      }
    } else {
      cell = cell.offset(0, 1);
    }
  }
}

    { startRow: 18, endRow: 31, column: 2 },
    { startRow: 33, endRow: 35, column: 2 },
    // Add more alternate forms ranges as needed
  ];

  for (let form of alternateForms) {
    for (let row = form.startRow; row <= form.endRow; row++) {
      let value = s1.getRange(row, form.column).getValue();
      if (value) {
        outputString += " (" + value + ")";
        break;
      }
    }
  }

  // Handle shiny
  if (s1.getRange(14, 2).getValue()) {
    outputString += " (Shiny☆)";
  }

  return outputString;
}

function handleUnownCheckbox(s, s1, s2, outputString, cell2) {
  const shiny = s1.getRange(14, 2).getValue();
  const letter = s1.getRange(6, 2).getValue().toString().charCodeAt(0);
  let unownCell = s2.getRange(letter - 61, 2);

  if (shiny) {
    unownCell.offset(0, 1).setValue("Caught☆");
  } else if (cell2.getValue() !== "Caught☆") {
    unownCell.offset(0, 1).setValue("Caught");
  }

  logEntry(unownCell.offset(0, 2), outputString);
}

function logEntry(cell, outputString) {
  let recorded = false;
  while (!recorded) {
    if (cell.getValue() === "") {
      if (cell.offset(0, -1).getValue() === outputString) {
        recorded = true;
      } else {
        cell.setValue(outputString);
        recorded = true;
      }
    } else {
      cell = cell.offset(0, 1);
    }
  }
}
