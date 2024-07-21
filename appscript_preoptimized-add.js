/*
Driver code, called when anything is edited on the spreadsheet
*/
function onEdit() {
  const s = SpreadsheetApp.getActiveSheet();
  const s0 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Main");
  const s1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Options");
  const s2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Unown");

  // If the user logged something on the Main sheet
  if (s.getName() === "Main") {
    let cell = s.getActiveCell();

    // Checks that the cell being edited is in column B and at or below row 5
    if (cell.getColumn() === 2 && cell.getRow() >= 5) {
      let cell2 = cell.offset(0, 1);
      let cell3 = cell.offset(0, 2);

      //If the pokemon checkbox is checked
      if(cell.getValue()) {
        
        // Unclick the checkbox so that another entry for this pokemon can be added later
        cell.setValue(false)

        // Collect the date, game, and pokemon details
        let outputString = collectData();

        // Record the new data
        logEntry(cell3, outputString);

        // If a Unown was added, then also add it to the unown research notes sheet
        if (cell.getRow() === 205) {
          const letter = s1.getRange(6, 2).getValue().toString().charCodeAt(0);
          let unownCell = s2.getRange(letter-61, 2);

          // Label the type of unown as "Caught" in the unown spreadsheet
          if (shiny.getValue()) {
            unownCell.offset(0, 1).setValue("Caught☆");
          }
          else if (cell2.getValue() !== "Caught☆") {
            unownCell.offset(0, 1).setValue("Caught");
          }

          // Record the new data
          unownCell = unownCell.offset(0, 2);
          
          logEntry(unownCell, outputString);
        }
      }

    }
    // If the "Current Game" field is being edited on the main sheet, update it in the options sheet
    else if (cell.getColumn() === 1 && cell.getRow() === 4) {
      s1.getRange(8, 2).setValue(cell.getValue());
    }

  }
  // If the user changed the Current Game from the options sheet
  else if (s.getName() === "Options") {
    const cell = s.getActiveCell();
    if (cell.getColumn() === 2 && cell.getRow() === 8) {
      // If the "Current Game" field is being edited on the options sheet, update it in the main sheet
      s0.getRange(4, 1).setValue(cell.getValue());
    }
  }
  // If the user logged a unown from the unown sheet
  else if (s.getName() === "Main") {
    let cell = s.getActiveCell();
    let cell2 = cell.offset(0, 1);
    let mainCell = s0.getRange(205, 3);

    // If the checkbox is checked
    if(cell.getValue()) {
      // Add new data!
      const date = new Date();
      const time_zone = "PST";
      const parsed_date = Utilities.formatDate(date, time_zone, 'MM/dd/yyyy');
      const game = s0.getRange(4, 1).getValue();
      const shiny = s1.getRange(14,2);

      // Uncheck the checkbox for next use
      cell.setValue("");

      // Record the unown in the unown research notes
      let outputString = "Captured " + parsed_date + " in " + game + " (" + cell.offset(0, -1).getValue().toString().split(" ")[1] + ")";

      // Label the type of unown as "Caught" in the unown spreadsheet
      if (shiny.getValue()) {
        outputString += " (Shiny☆)";
        cell2.setValue("Caught☆");
        //...and set the unown section on the main sheet to be shiny as well
        mainCell.setValue("Caught☆");
      }
      else if (cell2.getValue() !== "Caught☆") {
        cell2.setValue("Caught");
      }
      // Also label it as caught in the main spreadsheet
      else if (mainCell.getValue() !== "Caught☆") {
        mainCell.setValue("Caught");
      }

      cell = cell.offset(0, 2);
      logEntry(cell, outputString);

      // Record the unown in the regional pokedex
      cell = s0.getRange(205, 4);
      logEntry(cell, outputString);
    }

  }
}

/*
This function collects all necessary data from the spreadsheets to create a capture log, and parses it into a string.
Takes no inputs.

Example output:
If the date is September 5th, 2023, the game is Pokemon Moon, and the "Alolan?" box is ticked, the output should be:
"Captured 09/05/2023 in Moon (Alolan)"
*/
function collectData() {
  const s = SpreadsheetApp.getActiveSheet();
  const s1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Options");
  let cell = s.getActiveCell();
  let cell2 = cell.offset(0, 1);

  // Add new data!
  const date = new Date();
  const time_zone = "PST";
  const parsed_date = Utilities.formatDate(date, time_zone, 'MM/dd/yyyy');
  const game = s1.getRange(8, 2).getValue();
  const transfer_origin = s1.getRange(10, 2).getValue();
  const ot = s1.getRange(12, 2).getValue();

  // Important cells that we will collect info from
  const alolan = s1.getRange(2, 2);
  const galarian = s1.getRange(3, 2);
  const hisuian = s1.getRange(4, 2);
  const paldean = s1.getRange(5, 2);
  const other = s1.getRange(6, 2);
  const shiny = s1.getRange(14, 2);

  let outputString = "";

  if (transfer_origin !== "") {
    outputString = "Transferred " + parsed_date + " to " + game + " from " + transfer_origin;
    
  }
  else if (ot !== "") {
    outputString = "Traded " + parsed_date + " to " + game + " (OT: " + ot + ")";
  }
  else {
    outputString = "Captured " + parsed_date + " in " + game;
  }

  // Make the string that'll get printed into the cell
  if (alolan.getValue()) {
    outputString += " (Alolan)";
  }
  else if (galarian.getValue()) {
    outputString += " (Galarian)";
  }
  else if (hisuian.getValue()) {
    outputString += " (Hisuian)";
  }
  else if (paldean.getValue()) {
    outputString += " (Paldean)";
  }
  else if (other.getValue() !== "") {
    outputString += " (" + other.getValue() + ")";
  }
  else {
    let alt_form = ""

    // Check every Non-Regional Alternate Forms by Species box
    // Pikachu
    if (s1.getRange(18, 2).getValue()) alt_form = " (Cosplay)";
    else if (s1.getRange(19, 2).getValue()) alt_form = " (Rock Star)";
    else if (s1.getRange(20, 2).getValue()) alt_form = " (Belle)";
    else if (s1.getRange(21, 2).getValue()) alt_form = " (Pop Star)";
    else if (s1.getRange(22, 2).getValue()) alt_form = " (Ph. D)";
    else if (s1.getRange(23, 2).getValue()) alt_form = " (Libre)";
    else if (s1.getRange(24, 2).getValue()) alt_form = " (Original Cap)";
    else if (s1.getRange(25, 2).getValue()) alt_form = " (Hoenn Cap)";
    else if (s1.getRange(26, 2).getValue()) alt_form = " (Sinnoh Cap)";
    else if (s1.getRange(27, 2).getValue()) alt_form = " (Unova Cap)";
    else if (s1.getRange(28, 2).getValue()) alt_form = " (Kalos Cap)";
    else if (s1.getRange(29, 2).getValue()) alt_form = " (Alola Cap)";
    else if (s1.getRange(30, 2).getValue()) alt_form = " (Partner Cap)";
    else if (s1.getRange(31, 2).getValue()) alt_form = " (World Cap)";
    // Paldean Tauros
    else if (s1.getRange(33, 2).getValue()) alt_form = " (Combat Breed)";
    else if (s1.getRange(34, 2).getValue()) alt_form = " (Blaze Breed)";
    else if (s1.getRange(35, 2).getValue()) alt_form = " (Aqua Breed)";
    // Pichu
    else if (s1.getRange(37, 2).getValue()) alt_form = " (Spiky-eared)";
    // Castform
    else if (s1.getRange(39, 2).getValue()) alt_form = " (Normal)";
    else if (s1.getRange(40, 2).getValue()) alt_form = " (Sunny)";
    else if (s1.getRange(41, 2).getValue()) alt_form = " (Rainy)";
    else if (s1.getRange(42, 2).getValue()) alt_form = " (Snowy)";
    // Burmy & Wormadam
    else if (s1.getRange(44, 2).getValue()) alt_form = " (Plant Cloak)";
    else if (s1.getRange(45, 2).getValue()) alt_form = " (Sandy Cloak)";
    else if (s1.getRange(46, 2).getValue()) alt_form = " (Trash Cloak)";
    // Shellos & Gastrodon
    else if (s1.getRange(48, 2).getValue()) alt_form = " (West Sea)";
    else if (s1.getRange(49, 2).getValue()) alt_form = " (East Sea)";
    // Rotom
    else if (s1.getRange(51, 2).getValue()) alt_form = " (Heat)";
    else if (s1.getRange(52, 2).getValue()) alt_form = " (Wash)";
    else if (s1.getRange(53, 2).getValue()) alt_form = " (Frost)";
    else if (s1.getRange(54, 2).getValue()) alt_form = " (Fan)";
    else if (s1.getRange(55, 2).getValue()) alt_form = " (Mow)";
    // Basculin
    else if (s1.getRange(57, 2).getValue()) alt_form = " (Red-Striped)";
    else if (s1.getRange(58, 2).getValue()) alt_form = " (Blue-Striped)";
    else if (s1.getRange(59, 2).getValue()) alt_form = " (White-Striped)";
    // Deerling & Sawsbuck
    else if (s1.getRange(18, 5).getValue()) alt_form = " (Spring)";
    else if (s1.getRange(19, 5).getValue()) alt_form = " (Summer)";
    else if (s1.getRange(20, 5).getValue()) alt_form = " (Autumn)";
    else if (s1.getRange(21, 5).getValue()) alt_form = " (Winter)";
    // Vivillon
    else if (s1.getRange(23, 5).getValue()) alt_form = " (Archipelago)";
    else if (s1.getRange(24, 5).getValue()) alt_form = " (Continental)";
    else if (s1.getRange(25, 5).getValue()) alt_form = " (Elegant)";
    else if (s1.getRange(26, 5).getValue()) alt_form = " (Garden)";
    else if (s1.getRange(27, 5).getValue()) alt_form = " (High Plains)";
    else if (s1.getRange(28, 5).getValue()) alt_form = " (Icy Snow)";
    else if (s1.getRange(29, 5).getValue()) alt_form = " (Jungle)";
    else if (s1.getRange(30, 5).getValue()) alt_form = " (Marine)";
    else if (s1.getRange(31, 5).getValue()) alt_form = " (Meadow)";
    else if (s1.getRange(32, 5).getValue()) alt_form = " (Modern)";
    else if (s1.getRange(33, 5).getValue()) alt_form = " (Monsoon)";
    else if (s1.getRange(34, 5).getValue()) alt_form = " (Ocean)";
    else if (s1.getRange(35, 5).getValue()) alt_form = " (Polar)";
    else if (s1.getRange(36, 5).getValue()) alt_form = " (River)";
    else if (s1.getRange(37, 5).getValue()) alt_form = " (Sandstorm)";
    else if (s1.getRange(38, 5).getValue()) alt_form = " (Savanna)";
    else if (s1.getRange(39, 5).getValue()) alt_form = " (Sun)";
    else if (s1.getRange(40, 5).getValue()) alt_form = " (Tundra)";
    else if (s1.getRange(41, 5).getValue()) alt_form = " (Fancy)";
    else if (s1.getRange(42, 5).getValue()) alt_form = " (Poké Ball)";
    // Flabébé, Floette, & Florges
    else if (s1.getRange(44, 5).getValue()) alt_form = " (Red Flower)";
    else if (s1.getRange(45, 5).getValue()) alt_form = " (Yellow Flower)";
    else if (s1.getRange(46, 5).getValue()) alt_form = " (Orange Flower)";
    else if (s1.getRange(47, 5).getValue()) alt_form = " (Blue Flower)";
    else if (s1.getRange(48, 5).getValue()) alt_form = " (White Flower)";
    // Pumpkaboo & Gourgeist
    else if (s1.getRange(50, 5).getValue()) alt_form = " (Small)";
    else if (s1.getRange(51, 5).getValue()) alt_form = " (Medium)";
    else if (s1.getRange(52, 5).getValue()) alt_form = " (Large)";
    else if (s1.getRange(53, 5).getValue()) alt_form = " (Super)";
    // Zygarde
    else if (s1.getRange(55, 5).getValue()) alt_form = " (10% Forme)";
    else if (s1.getRange(56, 5).getValue()) alt_form = " (50% Forme)";
    else if (s1.getRange(57, 5).getValue()) alt_form = " (Complete Forme)";
    // Oricorio
    else if (s1.getRange(18, 8).getValue()) alt_form = " (Baile)";
    else if (s1.getRange(19, 8).getValue()) alt_form = " (Pom-Pom)";
    else if (s1.getRange(20, 8).getValue()) alt_form = " (Pa'u)";
    else if (s1.getRange(21, 8).getValue()) alt_form = " (Sensu)";
    // Lycanroc
    else if (s1.getRange(23, 8).getValue()) alt_form = " (Midday)";
    else if (s1.getRange(24, 8).getValue()) alt_form = " (Midnight)";
    else if (s1.getRange(25, 8).getValue()) alt_form = " (Dusk)";
    // Minior
    else if (s1.getRange(27, 8).getValue()) alt_form = " (Red Core)";
    else if (s1.getRange(28, 8).getValue()) alt_form = " (Orange Core)";
    else if (s1.getRange(29, 8).getValue()) alt_form = " (Yellow Core)";
    else if (s1.getRange(30, 8).getValue()) alt_form = " (Green Core)";
    else if (s1.getRange(31, 8).getValue()) alt_form = " (Blue Core)";
    else if (s1.getRange(32, 8).getValue()) alt_form = " (Indigo Core)";
    else if (s1.getRange(33, 8).getValue()) alt_form = " (Violet Core)";
    // Magearna
    else if (s1.getRange(35, 8).getValue()) alt_form = " (Original Color)";
    // Toxtricity
    else if (s1.getRange(37, 8).getValue()) alt_form = " (Amped)";
    else if (s1.getRange(38, 8).getValue()) alt_form = " (Low Key)";
    // Sinistea & Polteageist
    else if (s1.getRange(40, 8).getValue()) alt_form = " (Phony)";
    else if (s1.getRange(41, 8).getValue()) alt_form = " (Antique)";
    // Alcremie
    else if (s1.getRange(43, 8).getValue()) alt_form = " (Vanilla Cream)";
    else if (s1.getRange(44, 8).getValue()) alt_form = " (Ruby Cream)";
    else if (s1.getRange(45, 8).getValue()) alt_form = " (Matcha Cream)";
    else if (s1.getRange(46, 8).getValue()) alt_form = " (Mint Cream)";
    else if (s1.getRange(47, 8).getValue()) alt_form = " (Lemon Cream)";
    else if (s1.getRange(48, 8).getValue()) alt_form = " (Salted Cream)";
    else if (s1.getRange(49, 8).getValue()) alt_form = " (Ruby Swirl)";
    else if (s1.getRange(50, 8).getValue()) alt_form = " (Caramel Swirl)";
    else if (s1.getRange(51, 8).getValue()) alt_form = " (Caramel Swirl)";
    // Urshifu
    else if (s1.getRange(53, 8).getValue()) alt_form = " (Single Strike Style)";
    else if (s1.getRange(54, 8).getValue()) alt_form = " (Rapid Strike Style)";
    // Zarude
    else if (s1.getRange(56, 8).getValue()) alt_form = " (Dada)";
    // Maushold
    else if (s1.getRange(58, 8).getValue()) alt_form = " (Family of Four)";
    else if (s1.getRange(59, 8).getValue()) alt_form = " (Family of Three)";
    // Squawkabilly
    else if (s1.getRange(18, 11).getValue()) alt_form = " (Green Plumage)";
    else if (s1.getRange(19, 11).getValue()) alt_form = " (Blue Plumage)";
    else if (s1.getRange(20, 11).getValue()) alt_form = " (Yellow Plumage)";
    else if (s1.getRange(21, 11).getValue()) alt_form = " (White Plumage)";
    // Tatsugiri
    else if (s1.getRange(23, 11).getValue()) alt_form = " (Curly)";
    else if (s1.getRange(24, 11).getValue()) alt_form = " (Droopy)";
    else if (s1.getRange(25, 11).getValue()) alt_form = " (Stretchy)";
    // Dudunsparce
    else if (s1.getRange(27, 11).getValue()) alt_form = " (Two-Segment)";
    else if (s1.getRange(28, 11).getValue()) alt_form = " (Three-Segment)";
    // Gimmighoul
    else if (s1.getRange(30, 11).getValue()) alt_form = " (Chest)";
    else if (s1.getRange(31, 11).getValue()) alt_form = " (Roaming)";

    outputString += alt_form;
  }
  
  // Label the species as "Caught"
  if (shiny.getValue()) {
    outputString += " (Shiny☆)";
    cell2.setValue("Caught☆");
  }
  else if (cell2.getValue() !== "Caught☆") {
    cell2.setValue("Caught");
  }

  return outputString
}

/*
This function is meant to search the cell that it is given, followed by the cell to its right, followed to the cell to the right of that cell, etc. until an empty cell is found.
Then, outputString is put into that cell.

cell should be the first cell that will be checked to see if it is empty
outputString should be the string that will be placed into the first empty cell
*/
function logEntry(cell, outputString) {
  let recorded = false;
  while (!recorded) {
    
    if (cell.getValue() === "") {
      // If cell is empty, but the cell to the left of cell is equal to outputString, then never record this data because it's a repeat.
      if (cell.offset(0, -1).getValue() === outputString) {
        recorded = true;
      }
      // If cell is empty, and if the cell to the left of cell is not equal to outputString, record the data.
      else {
        cell.setValue(outputString);
        recorded = true;
      }
    }
    // Otherwise, cell is now the cell to the right of cell and loop
    else {
      cell = cell.offset(0, 1);
    }

  }
  return
}
