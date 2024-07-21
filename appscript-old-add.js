function update_existing_data() {
  // Should never need this code again. Just keeping it here for sentimental purposes ;🐢🐢🐢🐢
  // For the record, this code was used to update the existing capture data from the old format 
  // *I* used to use to the new one that is currently in use
  /*
  var s = SpreadsheetApp.getActiveSheet();
  for (var i = 5; i <= 1003; i++) {
    const current = s.getRange(i, 4);
    if (current.getValue() !== "") {
      const date = current.getValue().toString().split(" ")[2];
      const game = current.offset(0, 1).getValue().toString().substring(6);
      current.setValue("Captured " + date + " in " + game);
    }
  }
  */
}
