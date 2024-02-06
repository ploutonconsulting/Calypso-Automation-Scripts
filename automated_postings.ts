
const tAccountsWorksheetName = "t_accounts";

const postingSourceTableName = "postings_source";

const debitColumns: string[] = ["DEBIT_CLP_ACCOUNT", "DEBIT_ACCOUNT_NAME"];
const creditColumns: string[] = ["CREDIT_CLP_ACCOUNT", "CREDIT_ACCOUNT_NAME"];

const referenceTable = "accounts"

const tAccountColumns: string[] = ["PostingId", "EffectiveDate", "EventType", "PostingType", "Debit", "Credit"];

const defaultNumberOfAccountColumns: number = 2;

const tAccountPrefix: string = "account_";

const firstColumn = "B";
const secondColumn = "I";
const thirdColumn = "P";
const startingRow = 1
const rowsSkipped = 5;

const sortField: ExcelScript.SortField = {key: 1, ascending: true};

class Posting {

  postingId: number;
  transferId: number;

  debitAccount: string;
  creditAccount: string;

  effective_date: Date;

  eventType: string;
  postingType: string;

  amount: number;
  currency: string;

  constructor(postingId, transferId, debitAccount, creditAccount, effectiveDate, eventType, postingType, amount, currency) {

    this.postingId = postingId;
    this.transferId = transferId;
    this.debitAccount = debitAccount;
    this.creditAccount = creditAccount;
    this.effective_date = effectiveDate;
    this.eventType = eventType;
    this.postingType = postingType;
    this.amount = amount;
    this.currency = currency;
  }

  getPostingDetails(): (string|number|boolean)[] {
    return [this.postingId, this.effective_date.toString(), this.eventType, this.amount];
  }

  getDebitAccount(): string {
    return this.debitAccount;
  }

  getCreditAccount(): string {
    return this.creditAccount;
  }

}


function main(workbook: ExcelScript.Workbook) {
  console.log("Starting script...");

  console.log("Identifying unique accounts...");
  let uniqueAccounts = parseAccounts(workbook);

  console.log("Writing accounts to the reference table...");
  populateReferenceAccounts(workbook, uniqueAccounts);

  console.log("Creating t-accounts...");
  createTAccounts(workbook, uniqueAccounts);

  console.log("Parsing postings to class...");
  let postings = parsePostings(workbook);

  console.log("Populating the t-accounts with postings...")
  populateTAccounts(workbook, postings);
}

/**
 * Parse a list of unique SAP General Ledger account numbers from the postings array.
 */
function parseAccounts(workbook: ExcelScript.Workbook): (string | number | boolean)[][] {

  // Select the postings table
  let postingSource = workbook.getTable(postingSourceTableName);
  if (!postingSource) {
    console.log('Table ' + postingSourceTableName + ' not found.');
  }

  let debitAccounts = getAccountsInRange(debitColumns, postingSource)
  let creditAccounts = getAccountsInRange(creditColumns, postingSource);

  let accounts = debitAccounts.concat(creditAccounts);

  // Remove the duplicates from the concatenated array
  let uniqueAccounts = removeDuplicates(accounts);

  // Return a list of unique accounts
  return uniqueAccounts;
}

/**
 * Retrieve all the account details from the 'posting_souce' exported data.
 */
function getAccountsInRange(columns: string[], table: ExcelScript.Table): (string | number | boolean)[][] {
  console.log("Retrieving accounts for range " + columns + ".")

  // Retrieve the account details from the columns
  let accounts = columns.map(name => {
    let column = table.getColumn(name);
    if (!column) {
      console.log("Unable to find " + name + ".");
      return null;
    }

    let columnRange = column.getRangeBetweenHeaderAndTotal();
    let allAccounts = columnRange.getValues();
    return allAccounts;
  })

  // Move all the data to the first array
  for (let index = 0; index < accounts[0].length; index++) {
    accounts[0][index][1] = accounts[1][index][0];
  }

  // Return the first array
  console.log("Identified " + accounts[0].length + " accounts.")
  return accounts[0];
}

/**
 * Remove duplicates from an array and return an array with unique values.
 */
function removeDuplicates(array: (string | number | boolean)[][]): (string | number | boolean)[][] {
  console.log("Removing duplicates from array with " + array.length + " values.")

  // Create an empty list
  let seen = {};

  let uniqueArray: (string | number | boolean)[][] = [];

  // Loop through the array and remove all duplicates
  for (let i = 0; i < array.length; i++) {
    let identifier = array[i][0].toString();

    if (!seen[identifier]) {
      uniqueArray.push(array[i]);
      seen[identifier] = true;
    }
  }

  // Convert the array to a string array and return
  console.log("Identified " + uniqueArray.length + " unique records.")
  return uniqueArray;
}

/**
 * Add the accounts into the reference table(s).
 */
function populateReferenceAccounts(workbook: ExcelScript.Workbook, accounts: (string | number | boolean)[][]) {
  console.log("Populating the reference table.")

  // Select the reference table
  let accountReference = workbook.getTable(referenceTable);
  if (!accountReference) {
    console.log('Table ' + referenceTable + ' not found.');
    return;
  } else {
    "Identified reference table."
  }

  // Delete any existing rows in the reference table
  if (accountReference.getRowCount() > 0) {
    accountReference.deleteRowsAt(0, accountReference.getRowCount());
    console.log("Deleting all rows in the account reference table.")
  } else {
    console.log("No rows in the account reference table to delete.")
  }

  // Add the accounts to the table
  console.log("Adding " + accounts.length + " records to the reference table.")
  for (let index = 0; index < accounts.length; index++) {
    accountReference.addRow(-1, accounts[index]);
  }

  console.log("Reference accounts added.")

}

/**
 * Create a new table with a specific name and location in the worksheet.
 */
function createTable(worksheet: ExcelScript.Worksheet, column: string, row: number, tableName: string) {

  // Define the range where the table will be created
  // The range should be large enough to include headers and any initial data
  // Assuming the table starts at A1, the range will end at the last column of tAccountColumns
  let endColumnLetter = String.fromCharCode(column.charCodeAt(0) + tAccountColumns.length - 1);
  let tableRange = column + row + ":" + endColumnLetter + row; // Table with only header row initially

  // Convert the range to a table and give it a name
  let table = worksheet.addTable(tableRange, true);
  table.getHeaderRowRange().setValues([tAccountColumns]);
  table.setName(tAccountPrefix + tableName.valueOf());
  table.setShowFilterButton(false);
  table.setShowTotals(true);
  table.getColumnByName("Credit").getTotalRowRange().setFormulaLocal("=SUM([Debit])-SUM([Credit])");

  // Set the amount formats
  let debitColumn: ExcelScript.TableColumn = table.getColumnByName("Debit");
  debitColumn.getRange().setNumberFormat("#,##0");
  let creditColumn: ExcelScript.TableColumn = table.getColumnByName("Credit");
  creditColumn.getRange().setNumberFormat("#,##0");
  console.log("Table " + tableName + " created.");

  // Set the date formats
  let dateColumn = table.getColumnByName(tAccountColumns[1]);
  dateColumn.getRange().setNumberFormat("YYYY-MM-DD");
}

/**
 * Delete all the tables on the worksheet.
 */
function deleteTables(worksheet: ExcelScript.Worksheet) {

  let tables = worksheet.getTables();
  if (!tables || tables.length == 0) {
    console.log("No tables to delete on worksheet " + worksheet.getName() + ".");
  } else { }
  console.log("Deleting " + tables.length + " tables in " + worksheet.getName() + ".");
  tables.forEach(table => {
    table.delete();
  })
}

/**
 * Calculate the number of columns the output should use when creating the t-accounts.
 */
function calculateNumberOfColumns(numberOfAccounts: number): number {

  if (numberOfAccounts % 2 == 0) {
    console.log("Number of columns for " + numberOfAccounts + " t-accounts calculated as 2");
    return 2;
  }

  if (numberOfAccounts % 3 == 0) {
    console.log("Number of columns for " + numberOfAccounts + " t-accounts calculated is 3");
    return 3;
  }

  console.log("Using default (" + defaultNumberOfAccountColumns + ") columns for " + numberOfAccounts + " t-accounts.");
  return defaultNumberOfAccountColumns;

}

/**
 * Create the headers and tables for the t-accounts required.
 */
function createTAccounts(workbook: ExcelScript.Workbook, accounts: (string | number | boolean)[][]) {

  // Select the correct worksheet.
  let worksheet = workbook.getWorksheet(tAccountsWorksheetName);
  if (!worksheet) {
    console.log("Unable to select " + tAccountsWorksheetName + ".");
    return;
  }

  // Clear the worksheet before starting to add new content
  deleteTables(worksheet);
  let worksheetRange = worksheet.getUsedRange();
  worksheetRange.clear(ExcelScript.ClearApplyTo.all);


  // Determine the column layout of the t-accounts
  let numberOfColumns = calculateNumberOfColumns(accounts.length);
  console.log(numberOfColumns);
  let firstColumnCount = Math.ceil(accounts.length / numberOfColumns);
  let furtherColumnCount = Math.floor(accounts.length - firstColumnCount);
  console.log("Adding " + firstColumnCount + " accounts to first column and " + furtherColumnCount + " accounts for subsequent columns.");

  // Create headers for the first column.
  console.log("Creating t-accounts for column 1.");
  var currentRow = startingRow;
  for (let index = 0; index < firstColumnCount; index++) {
    let range = worksheet.getRange(firstColumn + currentRow + ":" + firstColumn + currentRow);
    range.setValue(accounts[index][0] + " - " + accounts[index][1].toString());
    createTable(worksheet, firstColumn, currentRow + 1, accounts[index][0].toString());
    currentRow += rowsSkipped;
  }

  // Create headers for the second column.
  if (numberOfColumns >= 2) {
    console.log("Creating t-accounts for column 2.");
    var currentRow = startingRow;
    for (let index = firstColumnCount; index < firstColumnCount + furtherColumnCount; index++) {
      let range = worksheet.getRange(secondColumn + currentRow + ":" + secondColumn + currentRow);
      range.setValue(accounts[index][0] + " - " + accounts[index][1].toString());
      createTable(worksheet, secondColumn, currentRow + 1, accounts[index][0].toString());
      currentRow += rowsSkipped;
    }
  } else {
    console.log("No t-accounts column 2 required.")
  }

  // Create headers for the third column.
  if (numberOfColumns == 3) {
    console.log("Creating t-accounts for column 3.");
    var currentRow = startingRow;
    for (let index = 0; index < furtherColumnCount; index++) {
      let range = worksheet.getRange(thirdColumn + currentRow + ":" + thirdColumn + currentRow);
      range.setValue(accounts[index][0] + " - " + accounts[index][1]);
      let tableHeader = accounts[index][1].toString();
      let trimmedHeader = tableHeader.substring(27);
      currentRow += rowsSkipped;
    }
  } else {
    console.log("No t-accounts required for column 3.");
  }
}

/**
 * Parse postings from the posting source table to an array of Posting objects.
 */
function parsePostings(workbook: ExcelScript.Workbook): Posting[] {

  let rows = getAllRows(workbook, postingSourceTableName);
  console.log("Found " + rows.length + " in the posting source table.");

  // Parse the values into the class
  let postings: Posting[] = [];

  rows.forEach((row) => {
    console.log("Parsing row: " + row);
    let newPosting = new Posting(
      row[0],   // postingId 
      row[4],   // transferId
      row[8],   // debitAccount
      row[11],  // creditAccount
      row[6],   // effectiveDate
      row[1],   // postingType
      row[2],   // eventType
      row[3],   // amount
      "ZAR"     // currency
    )
    postings.push(newPosting);
    console.log("Identified " + postings.length + " postings...");
  });

  console.log("Returning " + postings.length + " postings...");
  return postings;

}

/**
 * Fetch rows from a table and return.
 */
function getAllRows(workbook: ExcelScript.Workbook, tableName: string): (string | number | boolean)[][] {

  // Select the postings table
  let postingSource = workbook.getTable(tableName);
  if (!postingSource) {
    console.log('Table ' + tableName + ' not found.');
  }

  // Fetch the values from the table
  let range = postingSource.getRangeBetweenHeaderAndTotal();
  let rows = range.getValues();

  console.log("Extracted " + rows.length + " rows from the posting source table.");
  // Return the values
  return rows;
}

function populateTAccounts(workbook: ExcelScript.Workbook, postings: Posting[]) {

  console.log("Populating " + postings.length + "postings to t-accounts.");

  postings.forEach(posting => {
    
    // Extract account numbers
    let debitAccount = posting.getDebitAccount();
    console.log("Debit account is " + debitAccount);
    let creditAccount = posting.getCreditAccount();
    console.log("Credit account is " + creditAccount);

    // Identify the debit table
    let debitTable = workbook.getTable(tAccountPrefix + debitAccount);
    if (!debitTable) {
      console.log("Unable to identify debit t-account table.");
      return;
    }

    // Identify the credit table
    let creditTable = workbook.getTable(tAccountPrefix + creditAccount);
    if (!creditTable) {
      console.log("Unable to identify the credit t-account table.")
      return;
    }

    // Create the debit posting
    let debitPosting: (string|number|boolean)[] = [
      posting.postingId,
      posting.effective_date.toString(),
      posting.postingType,
      posting.eventType,
      posting.amount,
      ""
    ]
    debitTable.addRow(-1, debitPosting);
    debitTable.getRangeBetweenHeaderAndTotal().getFormat().autofitColumns;
    debitTable.getSort().apply([sortField]);

    // Create the credit posting
    let creditPosting: (string | number | boolean)[] = [
      posting.postingId,
      posting.effective_date.toString(),
      posting.postingType,
      posting.eventType,
      "",
      posting.amount
    ]
    creditTable.addRow(-1, creditPosting);
    creditTable.getRangeBetweenHeaderAndTotal().getFormat().autofitColumns;
    creditTable.getSort().apply([sortField]);
    

    console.log("Populated t-accounts with posting " + posting.postingId.toString() + ".");
  })
}