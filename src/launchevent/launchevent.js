// Initialize millisecond counter and start timer
Set(varMillisecondCounter, 0);
Set(varAutoStart, true);

// Capture the start time
Set(varStartTime, Now());
Set(varStartMillis, varMillisecondCounter);

// Collect selected parts into a collection and convert to JSON
Clear(colSelectedParts); // Clear the collection to ensure it starts empty
ForAll(
    Filter(parts, product_id = DropdownCanvas1.Selected.id), // Filter parts based on selected product ID
    Collect(colSelectedParts, {part_id: ThisRecord.id, name: ThisRecord.name, transaction_date: Today()}) // Collect part ID, name, and today's date
);
// Convert the collection to JSON format
Set(varPartIDJSON, JSON(colSelectedParts, JSONFormat.IncludeBinaryData));

// Call the stored procedure and capture the end time
With(
    {result: ms_dev_1.dboInsertSalesFromJson({json: varPartIDJSON})},
    // Capture the end time and milliseconds
    Set(varEndTime, Now());
    Set(varEndMillis, varMillisecondCounter);
    // Stop the timer
    Set(varAutoStart, false)
);

// Format times to show only time with milliseconds
Set(varStartTimeFormatted, Text(varStartTime, "hh:mm:ss") & "." & Text(Mod(varStartMillis, 1000), "000"));
Set(varEndTimeFormatted, Text(varEndTime, "hh:mm:ss") & "." & Text(Mod(varEndMillis, 1000), "000"));

Set(varDurationSeconds, DateDiff(varStartTime, varEndTime, TimeUnit.Seconds));
Set(varDurationMillis, varEndMillis - varStartMillis);
Set(varDurationFormatted, varDurationSeconds & "." & Text(Mod(varDurationMillis, 1000), "000") & " seconds");
