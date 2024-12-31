function generatePersonBlocks() {
  const DEBUG = true; // Set to true to enable logging
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Графік");
  const targetSheet = ss.getSheetByName("для_кожного");
  const peopleSheet = ss.getSheetByName("Список_людей");

  targetSheet.clear();

  // Load the Список_людей data into a map
  const peopleData = peopleSheet.getDataRange().getValues();
  const peopleMap = {};
  peopleData.forEach(row => {
    const name = row[1]?.trim(); // Assuming "B" is name
    const text = row[4]?.trim(); // Assuming "E" is text
    if (name && text) {
      peopleMap[name] = text;
    }
  });

  if (DEBUG) Logger.log("People Map:");
  if (DEBUG) Logger.log(JSON.stringify(peopleMap));

  const dataRange = sourceSheet.getDataRange();
  const data = dataRange.getValues();

  const daysOfWeek = ['неділя', 'понеділок', 'вівторок', 'середа', 'четвер', 'п’ятниця', 'субота'];

  const rows = data.map((row, index) => {
    const [day, time, name] = row;

    if (!day || !time || !name) {
      return null;
    }

    const formattedDay = day instanceof Date
      ? Utilities.formatDate(day, Session.getScriptTimeZone(), "dd.MM.yyyy")
      : day;

    const dayOfWeek = day instanceof Date
      ? daysOfWeek[day.getDay()]
      : '';

    return [formattedDay, dayOfWeek, time, name];
  }).filter(row => row !== null);

  const nameOrder = [];
  const nameMap = {};

  rows.forEach(([day, dayOfWeek, time, name]) => {
    if (!nameMap[name]) {
      nameMap[name] = [];
      if (!nameOrder.includes(name)) {
        nameOrder.push(name);
      }
    }
    nameMap[name].push([day, dayOfWeek, time]);
  });

  let rowIndex = 1;
  nameOrder.forEach(name => {
    const entries = nameMap[name];
    const additionalText = peopleMap[name] || '';
    const nameCell = targetSheet.getRange(rowIndex, 1);
    const additionalTextCell = targetSheet.getRange(rowIndex, 2);
    
    nameCell.setValue(name);
    nameCell.setFontWeight("bold");
    
    additionalTextCell.setValue(additionalText);
    additionalTextCell.setFontWeight("bold");
    additionalTextCell.setNumberFormat("@"); // Set format to Plain Text
    
    rowIndex++;
    entries.forEach(([day, dayOfWeek, time]) => {
      targetSheet.getRange(rowIndex, 1).setValue(day);
      targetSheet.getRange(rowIndex, 2).setValue(dayOfWeek);
      targetSheet.getRange(rowIndex, 3).setValue(time);
      rowIndex++;
    });
    rowIndex++;
  });

  // Automatically set date column format to "dd.MM.yyyy"
  targetSheet.getRange(1, 1, targetSheet.getLastRow(), 1).setNumberFormat("dd.MM.yyyy");
}
