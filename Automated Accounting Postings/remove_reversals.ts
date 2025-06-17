
const postingSourceTableName = "postings_source";


class Posting {

  postingId: string;
  transferId: string;
  linkedId: string;

  debitAccount: string;
  creditAccount: string;

  effective_date: string;

  eventType: string;
  postingType: string;

  amount: string;
  currency: string;

  constructor(postingId: string, transferId: string, linkedId: string, debitAccount: string, creditAccount: string, effectiveDate: string, eventType: string, postingType: string, amount: string, currency: string) {

    this.postingId = postingId;
    this.transferId = transferId;
    this.linkedId = linkedId;
    this.debitAccount = debitAccount;
    this.creditAccount = creditAccount;
    this.effective_date = effectiveDate;
    this.eventType = eventType;
    this.postingType = postingType;
    this.amount = amount;
    this.currency = currency;
  }

  getPostingDetails(): (string | number | boolean)[] {
    return [this.postingId, this.effective_date, this.eventType, this.amount];
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

  console.log("Parsing postings to objects...");
  let postings = parsePostings(workbook);

  console.log("Parsed postings. Identifying REVERSAL items...");
  let reversalItems = identifyReversals(postings);
  
  highlightRows(workbook, reversalItems)

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
    let newPosting = new Posting(
      row[0].toString(),   // postingId 
      row[4].toString(),   // transferId
      row[5].toString(),   // linkedId
      row[8].toString(),   // debitAccount
      row[11].toString(),  // creditAccount
      row[6].toString(),   // effectiveDate
      row[1].toString(),   // postingType
      row[2].toString(),   // eventType
      row[3].toString(),   // amount
      "ZAR"     // currency
    )
    postings.push(newPosting);
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

  // Clear any highlight formatting
  range.getFormat().getFill().clear();

  console.log("Extracted " + rows.length + " rows from the posting source table.");
  // Return the values
  return rows;
}

function identifyReversals(postings: Posting[]) {
  
  var reversals: String[] = [];
  
  for (let index = 0; index < postings.length; index++) {
    
    if (postings[index].linkedId != "0") {
      let reversalPosting = getPosting(postings, postings[index].linkedId);
      if (reversalPosting != null && 
          comparePostings(postings[index], reversalPosting)) {
        reversals.push(postings[index].postingId);
        reversals.push(reversalPosting.postingId);
      }
    }
  }

  return reversals;

}

function comparePostings(posting1: Posting, posting2: Posting): Boolean {
   if (posting1.amount == posting2.amount &&
      posting1.postingType == posting2.postingType && 
      posting1.effective_date == posting2.effective_date) {
      
      return true;
    } else {
      return false;
    }
}

function getPosting(postings: Posting[], postingId: String): Posting {
  var found = false;
  var index = 0;

  while (index < postings.length) {
    if (postings[index].postingId ==  postingId) {
      return postings[index];
    }

    index++;
  }

  console.log("No posting with id " + postingId + "identified...");
  return null;
}

function highlightRows(workbook: ExcelScript.Workbook, reversalIds: String[]) {

  const postingSourceTableName = "postings_source";
  let postingSource = workbook.getTable(postingSourceTableName);
  if (!postingSource) {
    console.log('Table not found.');
  }

  // Fetch the values from the table
  let range = postingSource.getRangeBetweenHeaderAndTotal();
  let rows = range.getValues();
  
  for (let reversalIndex = 0; reversalIndex < reversalIds.length; reversalIndex++) {
    var rowIndex = 0;
    var itemFound = false;
    while (!itemFound && rowIndex < rows.length) {
      if (rows[rowIndex][0] == reversalIds[reversalIndex]) {
        let rowRange = range.getRow(rowIndex);
        rowRange.getFormat().getFill().setColor("red");
        itemFound = true;
      } else {
        rowIndex++;
      }
    }
  }
}
