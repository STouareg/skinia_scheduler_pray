function createEventsForSpecificPerson() {
  const DEBUG = false; // Set to true to enable logging, false to disable
  const sheetName = "для_кожного"; 
  const personName = "Турик Сергій"; 
  const eventTitle = "Молитва за ст. пастора";
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const calendar = CalendarApp.getDefaultCalendar();

  let creatingEvents = false;
  let eventsCreated = 0;

  if (DEBUG) console.log(`Processing events for: ${personName}`);

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const firstColumn = row[0] ? row[0].toString().trim() : ""; // Person name or date
    const timeRange = row[2]; // Time range column

    if (DEBUG) console.log(`Row ${i + 1}: ${JSON.stringify(row)}`);

    // Detect the start of a person's block
    if (firstColumn === personName) {
      creatingEvents = true;
      if (DEBUG) console.log(`Started creating events for: ${personName}`);
      continue;
    }

    // Stop processing when reaching a blank row or another person's name
    if (creatingEvents && (!firstColumn || isNaN(Date.parse(firstColumn)))) {
      if (DEBUG) console.log(`Finished creating events for: ${personName}`);
      break;
    }

    // Process rows with event data
    if (creatingEvents && !isNaN(Date.parse(firstColumn)) && timeRange) {
      try {
        const startTime = new Date(firstColumn); // Parse date from column 0
        const [start, end] = timeRange.split("-");
        startTime.setHours(...start.split(":").map(Number));
        const endTime = new Date(startTime);
        endTime.setHours(...end.split(":").map(Number));

        // Create the event in Google Calendar
        const event = calendar.createEvent(eventTitle, startTime, endTime);
        event.addPopupReminder(5); // 5 min before
        event.addPopupReminder(60); // 1 hour before
        event.addPopupReminder(24 * 60); // 1 day before

        eventsCreated++;
        if (DEBUG) console.log(`Created event: ${eventTitle} on ${startTime.toLocaleString()} to ${endTime.toLocaleString()}`);
      } catch (error) {
        if (DEBUG) console.error(`Error creating event for row ${i + 1}: ${error.message}`);
      }
    }
  }

  if (DEBUG) console.log(`Total events created for ${personName}: ${eventsCreated}`);
}
