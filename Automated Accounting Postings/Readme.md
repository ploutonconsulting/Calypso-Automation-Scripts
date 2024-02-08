# Automated Accounting Postings
The purpose of this file and script is to use postings extracted from Calypso and parse those postings to t-accounts that can be reviewed with the end users.

## Source Data
To function, extract postings data from Calypso and past the data into the table located in the `postings_source` tab.
- BO_POSTING_ID
- POSTING_TYPE
- BO_POSTING_TYPE
- AMOUNT
- TRANSFER_ID
- LINKED_ID
- EFFECTIVE_DATE
- CURRENCY_CODE
- DEBIT_CLP_ACCOUNT
- DEBIT_ACCOUNT_NAME
- DEBIT_GL_ACCOUNT
- CREDIT_CLP_ACCOUNT
- CREDIT_ACCOUNT_NAME
- CREDIT_GL_ACCOUNT
- SENT_DATE
- SENT_STATUS

Although not all values are required for the t-accounts, additional values are kept for future enhancements.

## Parse the information to the t-accounts
Click on the 'Parse Account Postings' to parse the postings in the source table to the tables in the `t-accounts` tab.
The script will execute and process the following:
1. Extract all postings in the posting source table.
2. Parse the unique account numbers from the table and populate this to the `reference_data`.
3. Parse the t-account debit and credit entries and populate the tables in the `t_accounts` tab.

## Important
1. Do not amend any of the table columns, headers or names in the file.
2. Do not amend any of the tab names, etc.
3. You may add additional tabs for more information.
4. It is safe to reference the source and other tables using Excel formulas.

For any questions, comments or bugs please contact me on pierre.oosthuizen@ploutonconsulting.com

## Fixes and updates
- **1.1 (20240208):** Fixed issue where the tables in column 3 is not updating with Excel tables, just headers.
  
