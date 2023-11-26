/* eslint-disable prettier/prettier */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
//@ts-nocheck

function main(workbook: ExcelScript.Workbook) {
  // Boilerplate: Getting the current sheet within the workbook
  const currentSheet: ExcelScript.Worksheet = workbook.getActiveWorksheet();

  // Get total ticket count - Used for calculating ranges 
  let ticketCount = currentSheet.getUsedRange().getRowCount();

  // Generating a names string array from the file name 
  /*
    The .getName() function returns the full file name as a string
    The .replace('.xlsx') clears out that part of the string and the .split('-') 
    function creates the names array. It's imporant to separate the filename by 
    dashes, but you can also indicate any other character inside this function. 
  */
  let names: string[] = (workbook.getName().replace('.xlsx', '')).split('-')

  /*
    Making sure we don't assign too many tickets, max 30 per tech
  */
  const maxTicketCount = 30 * names.length;

  if (ticketCount > maxTicketCount) {
    let ticketsCountToDelete = ticketCount - maxTicketCount

    currentSheet.getRange(`A${ticketCount}:J${ticketCount - ticketsCountToDelete}`).delete(ExcelScript.DeleteShiftDirection.up)
  }

  // updating the ticket count 
  ticketCount = currentSheet.getUsedRange().getRowCount();

  // -----formatting starts here; we need: ticketCount, names, and currentSheet

  const ticketRange: ExcelScript.Range = setStyling(currentSheet, names, ticketCount);

  // ----formatting ends here 

  /* Sorting per tech starts here */

  //names column gets returned as a 2D array 
  const namesColumn = currentSheet.getRange(`J2:J${ticketCount}`).getValues();

  let lastName = namesColumn[2][0]
  let startRow = 2
  let endRow = 0

  //create a new sheet 
  const tempSheet: ExcelScript.Worksheet = workbook.addWorksheet('temp');

  for (let i = 2; i < ticketCount; i++) {
    let currName = ""
    try {
      currName = namesColumn[i][0].toString();
    } catch (error) {
      break;
    }

    if (currName != lastName) {
      //if we get here we're at the start of the new section 
      endRow = i + 1

      //---- section extraction and sorting will happen here 

      // extract the range 
      let techRange = currentSheet.getRange(`A${startRow}:J${endRow}`);

      //copy the values over 

      // Paste to range A1 on temp from range A2:J29 on selectedSheet
      tempSheet.getRange("A1").copyFrom(techRange, ExcelScript.RangeCopyType.values, false, false);
      // tempSheet.getRange().setValues(techRange.getValues());

      // short by short description 
      let tempTicketCount = tempSheet.getUsedRange().getRowCount()
      let tempSortRange = tempSheet.getRange(`A1:J${tempTicketCount}`);
      tempSortRange.getSort().apply([{ key: 2, ascending: true }], false, true, ExcelScript.SortOrientation.rows);

      // copy the values back to the main sheet 
      // copy-from doesn't work too well
      // currentSheet.getRange(`A${startRow}`).copyFrom(tempSheet.getUsedRange(), ExcelScript.RangeCopyType.values, false, false);

      currentSheet.getRange(`A${startRow}:J${endRow}`).setValues(tempSheet.getUsedRange().getValues());

      // update the start row 
      startRow = i + 2

      // update the last name 
      lastName = currName

      //clear out the temp sheet
      tempSheet.getUsedRange().clear();
    }
  }

  // --- now we need to manually do it for the last tech 
  endRow = ticketCount;
  let techRange = currentSheet.getRange(`A${startRow}:J${endRow}`);

  tempSheet.getRange("A1").copyFrom(techRange, ExcelScript.RangeCopyType.values, false, false);

  let tempTicketCount = tempSheet.getUsedRange().getRowCount()
  let tempSortRange = tempSheet.getRange(`A1:J${tempTicketCount}`);
  tempSortRange.getSort().apply([{ key: 2, ascending: true }], false, true, ExcelScript.SortOrientation.rows);

  currentSheet.getRange(`A${startRow}:J${endRow}`).setValues(tempSheet.getUsedRange().getValues());

  tempSheet.delete();


  // ---validation starts here: need: currentSheet, ticketCount, ticketRange 

  setValidation(currentSheet, ticketCount, ticketRange);

  // ----- Validation ENDS Here

  // Note: Color code can't be ran more than once (weird stuff happens)

  // ---- setColorFills starts here: needs currentsheet, ticketRange, 

  const statusColors: string[] = ["#FF0000", "#FFFF00", "#92D050"]

  const assignmentStartRows: number[] = setColorFillsAndFormulas(currentSheet, ticketRange, statusColors);

  // ----- setColorFills ends here

  // ---- setDashboard starts here, needs: dashBoardSheet, assignmentStartRows

  // Renaming the current sheet and adding the dashboard 
  currentSheet.setName("Tasks");

  //TODO: No 
  currentSheet.getUsedRange().getFormat().autofitColumns();
  currentSheet.getRange("C:C").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.top);
  currentSheet.getRange("C:C").getFormat().setIndentLevel(0);
  const dashBoardSheet: ExcelScript.Worksheet = workbook.addWorksheet("DashBoard");

  // Setting and Formatting DashBoard Data
  setDashBoard(dashBoardSheet, assignmentStartRows, statusColors);




}

function setStyling(currentSheet: ExcelScript.Worksheet, names: string[], ticketCount: number): ExcelScript.Range {
  let namesList: string[][] = []

  // Formatting the names in a 2D array to place in Excel
  for (let i in names) { namesList.push([names[i]]); }

  // Set the Assigned Header
  currentSheet.getRange('J1').setValue('Assigned To');

  // Set the names
  const namesRange = currentSheet.getRange(`J2:J${names.length + 1}`);
  namesRange.setValues(namesList);

  // Gets the entire ticket range 
  const ticketRange: ExcelScript.Range = currentSheet.getRange(`J2:J${ticketCount}`)

  // Fill out the names on the rest of the ticket list and sort them for visibility
  namesRange.autoFill(ticketRange, ExcelScript.AutoFillType.fillDefault);
  ticketRange.getSort().apply([{ key: 0, ascending: true }], false, false, ExcelScript.SortOrientation.rows);

  // Set the font name and size 
  const fontOptions: ExcelScript.RangeFont = currentSheet.getRange().getFormat().getFont()
  fontOptions.setName('Calibri');
  fontOptions.setSize(11);

  // Set row height 
  currentSheet.getUsedRange().getFormat().setRowHeight(18);

  // Auto fit columns to view data
  const allWidthOptions = currentSheet.getUsedRange().getFormat();
  allWidthOptions.autofitColumns();

  /*
    Short descriptions would be too long to be viewable on this excel sheet. 
    The following set of instructions is setting the short description column 
    length at 50% of the longest size
  */
  const shortDescColumnFormat = currentSheet.getRange('C1').getFormat();
  const shortDescColumnWidth = shortDescColumnFormat.getColumnWidth();
  shortDescColumnFormat.setColumnWidth(Math.round(shortDescColumnWidth / 2));

  // Setting Header Fill color 
  const columnHeaderFormat = currentSheet.getRange('A1:J1').getFormat()
  const columnHeaderFont = columnHeaderFormat.getFont()
  columnHeaderFont.setBold(true);
  const columnHeaderFillColor = columnHeaderFormat.getFill();
  columnHeaderFillColor.setColor('#D6DCE4') // aka Blue-Gray, Text 2, Lighter 80%

  // Set Border Style
  const borderFormat = currentSheet.getUsedRange().getFormat()

  //Edges
  let edgeBottom = borderFormat.getRangeBorder(ExcelScript.BorderIndex.edgeBottom);
  let edgeRight = borderFormat.getRangeBorder(ExcelScript.BorderIndex.edgeRight);
  let edgeTop = borderFormat.getRangeBorder(ExcelScript.BorderIndex.edgeTop);
  let edgeLeft = borderFormat.getRangeBorder(ExcelScript.BorderIndex.edgeLeft);

  //Insides
  let insideHorizontal = borderFormat.getRangeBorder(ExcelScript.BorderIndex.insideHorizontal)
  let insideVertical = borderFormat.getRangeBorder(ExcelScript.BorderIndex.insideVertical)

  /* 
    You can adjust each individual border weight if needed. Here I made an array of 
    all the edges and applied the style desired with a for-loop
  */
  const edges: ExcelScript.RangeBorder[] = [edgeBottom, edgeLeft, edgeTop, edgeRight];
  const insides: ExcelScript.RangeBorder[] = [insideHorizontal, insideVertical]

  for (let i in edges) {
    edges[i].setStyle(ExcelScript.BorderLineStyle.continuous);
    edges[i].setColor('000000');
    edges[i].setWeight(ExcelScript.BorderWeight.thin);
  }

  for (let i in insides) {
    insides[i].setStyle(ExcelScript.BorderLineStyle.continuous);
    insides[i].setColor('000000');
    insides[i].setWeight(ExcelScript.BorderWeight.thin);
  }

  return ticketRange
}

function setValidation(currentSheet: ExcelScript.Worksheet, ticketCount: number, ticketRange: ExcelScript.Range) {
  // Creates a range object for the A column starting at A2
  const dropDownRange: ExcelScript.Range = currentSheet.getRange('A2:A' + ticketCount);

  // Creating the options for the drop down data validation rule
  const dropDownOptions: ExcelScript.ListDataValidation = {
    inCellDropDown: true,
    source: "Open, Closed, Escalated"
  };

  // Providing the options to the data validation rule 
  const dropDownValidation: ExcelScript.DataValidationRule = {
    list: dropDownOptions
  };

  // Getting the current data validation rules
  const rangeDataValidation: ExcelScript.DataValidation = dropDownRange.getDataValidation();

  // Setting our rule  
  /*
    Per Github: You will receive a set rule error if you try to apply a rule to a a cell that already 
    has a rule. If you are running rule tests and a test fails, you will have to open/close the workbook
    (without saving changes) and re-run the tests with your new code. Additionally, you can also 
    clear rules for that range. Here, I am pre-emptively clearing the rules for my range before I apply 
    the rule I made. 
  */
  rangeDataValidation.clear();
  rangeDataValidation.setRule(dropDownValidation);

  // Create an alert to appear when an option is not selected from the list .
  const listOptionsOnlyAlert: ExcelScript.DataValidationErrorAlert = {
    message: "Must be set to Open, Closed, or Escalated",
    showAlert: true,
    style: ExcelScript.DataValidationAlertStyle.stop,
    title: "Invalid Selection"
  };

  // Alert user for status invalid status selection
  rangeDataValidation.setErrorAlert(listOptionsOnlyAlert);
}

function setColorFillsAndFormulas(currentSheet: ExcelScript.Worksheet, ticketRange: ExcelScript.Range, statusColors: string[]): number[] {
  /*
    Here the assignments are a 2D array of users (Welcome to Office Script). 
    I'm saving the name of the last assigned user, and comparing it to the name of the currently 
    user. If it's different, I will switch to the next color pair. 
  */
  const assignments = ticketRange.getValues();
  let lastAssignment = assignments[0][0]

  currentSheet.getRange(`L2`).setValue(lastAssignment)

  currentSheet.getRange(`L3`).setValue("Open")
  currentSheet.getRange(`L3`).getFormat().getFill().setColor(statusColors[0])

  currentSheet.getRange(`L4`).setValue("Escalated")
  currentSheet.getRange(`L4`).getFormat().getFill().setColor(statusColors[1])

  currentSheet.getRange(`L5`).setValue("Closed")
  currentSheet.getRange(`L5`).getFormat().getFill().setColor(statusColors[2])

  /*
    Colors Index will help navigate the colorPairs, we switch to the next pair on the next index.
    This is also where you change how many users this script supports. If you have another color pair,
    you could add it and now this script can support 5, and so on. 
  */
  let colorsIndex = 0;
  const colorPairs: string[][] = [
    ['#A9D08E', '#E2EFDA'],
    ['#F8CBAD', '#FCE4D6'],
    ['#FFE699', '#FFF2CC'],
    ['#2F75B5', '#9BC2E6']
  ]

  let assignmentStartRows: number[] = [];

  let assignmentStartRow = 2;

  assignmentStartRows.push(assignmentStartRow);

  currentSheet.getRange('N1').setValue('Closed')
  currentSheet.getRange('O1').setValue('Escalated')
  currentSheet.getRange('P1').setValue('Open')

  /*
    I'm sure there's an opportunity for optimzation here. I use a for loop to go through each ticket assignment and get the current fill color. I check to see if I'm in another user's bucket, if I am I keep my color pairs. If I'm not, I go to the next color pair set. I alternate colors by checking if the row number is even.  
  */
  for (let i = 2; i <= assignments.length + 1; i++) {
    let assignmentsRangeFillColor = currentSheet.getRange(`A${i}:J${i}`).getFormat().getFill();

    let currAssignment = assignments[i - 2][0]

    if (lastAssignment != currAssignment) {
      colorsIndex += 1
      lastAssignment = currAssignment
      let assignmentEndRow = i - 1;
      currentSheet.getRange(`L${i}`).setValue(lastAssignment)
      currentSheet.getRange(`N${assignmentStartRow}`).setFormula(`=IF($A$${assignmentStartRow}:A$${assignmentEndRow}=N$1,1,0)`)
      currentSheet.getRange(`O${assignmentStartRow}`).setFormula(`=IF($A$${assignmentStartRow}:A$${assignmentEndRow}=O$1,1,0)`)
      currentSheet.getRange(`P${assignmentStartRow}`).setFormula(`=IF($A$${assignmentStartRow}:A$${assignmentEndRow}=P$1,1,0)`)

      currentSheet.getRange(`L${i + 1}`).setValue("Open")
      currentSheet.getRange(`L${i + 1}`).getFormat().getFill().setColor(statusColors[0])
      currentSheet.getRange(`M${assignmentStartRow + 1}`).setFormula(`=SUM($P$${assignmentStartRow}:$P$${assignmentEndRow})`);

      currentSheet.getRange(`L${i + 2}`).setValue("Escalated")
      currentSheet.getRange(`L${i + 2}`).getFormat().getFill().setColor(statusColors[1])
      currentSheet.getRange(`M${assignmentStartRow + 2}`).setFormula(`=SUM($O$${assignmentStartRow}:$O$${assignmentEndRow})`);

      currentSheet.getRange(`L${i + 3}`).setValue("Closed")
      currentSheet.getRange(`L${i + 3}`).getFormat().getFill().setColor(statusColors[2])
      currentSheet.getRange(`M${assignmentStartRow + 3}`).setFormula(`=SUM($N$${assignmentStartRow}:$N$${assignmentEndRow})`);

      assignmentStartRow = i;
      assignmentStartRows.push(assignmentStartRow);
    }

    if ((i % 2) == 0) { assignmentsRangeFillColor.setColor(colorPairs[colorsIndex][0]); }
    else { assignmentsRangeFillColor.setColor(colorPairs[colorsIndex][1]); }
  }
  let assignmentEndRow = assignments.length + 1

  // Manually Setting the last assignment here 
  currentSheet.getRange(`N${assignmentStartRow}`).setFormula(`=IF($A$${assignmentStartRow}:A$${assignmentEndRow}=N$1,1,0)`)
  currentSheet.getRange(`O${assignmentStartRow}`).setFormula(`=IF($A$${assignmentStartRow}:A$${assignmentEndRow}=O$1,1,0)`)
  currentSheet.getRange(`P${assignmentStartRow}`).setFormula(`=IF($A$${assignmentStartRow}:A$${assignmentEndRow}=P$1,1,0)`)

  currentSheet.getRange(`M${assignmentStartRow + 1}`).setFormula(`=SUM($P$${assignmentStartRow}:$P$${assignmentEndRow})`);
  currentSheet.getRange(`M${assignmentStartRow + 2}`).setFormula(`=SUM($O$${assignmentStartRow}:$O$${assignmentEndRow})`);
  currentSheet.getRange(`M${assignmentStartRow + 3}`).setFormula(`=SUM($N$${assignmentStartRow}:$N$${assignmentEndRow})`);

  return assignmentStartRows;
}

function setDashBoard(dashBoardSheet: ExcelScript.Worksheet, assignmentStartRows: number[], statusColors: string[]) {
  // Setting and Formatting DashBoard Data
  let currRow = 2
  for (let i = 0; i < assignmentStartRows.length; i++) {
    dashBoardSheet.getRange(`B${currRow}`).setFormula(`=Tasks!L${assignmentStartRows[i]}:M${assignmentStartRows[i] + 3}`)

    dashBoardSheet.getRange(`B${currRow}:C${currRow}`).getFormat().getFill().setColor("#E7E6E6");
    dashBoardSheet.getRange(`C${currRow}`).getFormat().getFont().setColor("#E7E6E6");

    dashBoardSheet.getRange(`B${currRow + 1}:C${currRow + 1}`).getFormat().getFill().setColor(statusColors[0]);
    dashBoardSheet.getRange(`B${currRow + 2}:C${currRow + 2}`).getFormat().getFill().setColor(statusColors[1]);
    dashBoardSheet.getRange(`B${currRow + 3}:C${currRow + 3}`).getFormat().getFill().setColor(statusColors[2]);

    dashBoardSheet.getRange(`B${currRow}:C${currRow + 3}`).getFormat().getFont().setSize(16);
    dashBoardSheet.getRange(`B${currRow}:C${currRow + 3}`).getFormat().autofitColumns();

    currRow += 5;
  }
}