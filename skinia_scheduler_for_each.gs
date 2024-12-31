function generatePersonBlocks() {
  const DEBUG = false; // Set to true to enable logging
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Графік");
  const targetSheet = ss.getSheetByName("для_кожного");

  targetSheet.clear();

  const dataRange = sourceSheet.getDataRange();
  const data = dataRange.getValues();

  if (DEBUG) Logger.log("Raw Dataset:");
  if (DEBUG) data.forEach((row, index) => Logger.log(`Row ${index + 1}: ${JSON.stringify(row)}`));

  const daysOfWeek = ['неділя', 'понеділок', 'вівторок', 'середа', 'четвер', 'п’ятниця', 'субота'];

  const rows = data.map((row, index) => {
    if (DEBUG) Logger.log(`Before Mapping Row ${index + 1}: ${JSON.stringify(row)}`);
    const [day, time, name] = row;

    if (!day || !time || !name) {
      if (DEBUG) Logger.log(`Skipping Row ${index + 1} due to missing data.`);
      return null;
    }

    const formattedDay = day instanceof Date
      ? Utilities.formatDate(day, Session.getScriptTimeZone(), "dd.MM.yyyy")
      : day;

    const dayOfWeek = day instanceof Date
      ? daysOfWeek[day.getDay()]
      : '';

    if (DEBUG) Logger.log(`Mapped Row ${index + 1}: Day=${formattedDay}, Time=${time}, Name=${name}, DayOfWeek=${dayOfWeek}`);
    return [formattedDay, dayOfWeek, time, name];
  }).filter(row => row !== null);

  if (DEBUG) Logger.log("Mapped Rows:");
  if (DEBUG) rows.forEach((row, index) => Logger.log(`Row ${index + 1}: ${JSON.stringify(row)}`));

  const nameOrder = [];
  const nameMap = {};

  rows.forEach(([day, dayOfWeek, time, name], index) => {
    if (DEBUG) Logger.log(`Processing Mapped Row ${index + 1}: Day=${day}, DayOfWeek=${dayOfWeek}, Time=${time}, Name=${name}`);
    if (!nameMap[name]) {
      nameMap[name] = [];
      if (!nameOrder.includes(name)) {
        nameOrder.push(name);
        if (DEBUG) Logger.log(`Added "${name}" to nameOrder at position ${nameOrder.length - 1}`);
      }
    }
    nameMap[name].push([day, dayOfWeek, time]);
  });

  if (DEBUG) Logger.log("Final nameOrder:");
  if (DEBUG) Logger.log(nameOrder);

  let rowIndex = 1;
  nameOrder.forEach((name, orderIndex) => {
    if (DEBUG) Logger.log(`Writing "${name}" (index ${orderIndex}) to sheet starting at row ${rowIndex}`);
    const entries = nameMap[name];
    targetSheet.getRange(rowIndex, 1).setValue(name);
    targetSheet.getRange(rowIndex, 1).setFontWeight("bold");
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

  if (DEBUG) Logger.log("Output written to 'для_кожного'");
}
