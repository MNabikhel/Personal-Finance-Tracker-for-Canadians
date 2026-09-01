Attribute VB_Name = "modConst"
Option Explicit

'== Canadian Finance Tracker ==================================================
' Shared constants. Sheet and range names live here so that renaming anything
' in the workbook only needs one edit.
'=============================================================================

Public Const APP_NAME As String = "Canadian Finance Tracker"
Public Const APP_VERSION As String = "1.0.0"

' Worksheet names
Public Const SH_DASHBOARD As String = "Dashboard"
Public Const SH_TXN As String = "Transactions"
Public Const SH_ACCOUNTS As String = "Accounts"
Public Const SH_CATEGORIES As String = "Categories"
Public Const SH_RULES As String = "Rules"
Public Const SH_BUDGET As String = "Budget"
Public Const SH_REPORTS As String = "Reports"
Public Const SH_HOUSEHOLD As String = "Household"
Public Const SH_TAX As String = "Tax Summary"
Public Const SH_REGISTERED As String = "Registered Plans"
Public Const SH_FORMATS As String = "Bank Formats"
Public Const SH_SETTINGS As String = "Settings"
Public Const SH_LOG As String = "Import Log"
Public Const SH_HELP As String = "Help"
Public Const SH_ENGINE As String = "Engine"

' Table names
Public Const TBL_TXN As String = "tblTxn"
Public Const TBL_ACCOUNTS As String = "tblAccounts"
Public Const TBL_CATEGORIES As String = "tblCategories"
Public Const TBL_RULES As String = "tblRules"
Public Const TBL_FORMATS As String = "tblFormats"
Public Const TBL_LOG As String = "tblLog"

' Transaction column headers
Public Const COL_ID As String = "Txn ID"
Public Const COL_DATE As String = "Date"
Public Const COL_MONTH As String = "Month"
Public Const COL_ACCOUNT As String = "Account"
Public Const COL_PAIDBY As String = "Paid By"
Public Const COL_OWNER As String = "Owner"
Public Const COL_DESC As String = "Description"
Public Const COL_MERCHANT As String = "Merchant"
Public Const COL_AMOUNT As String = "Amount"
Public Const COL_CATEGORY As String = "Category"
Public Const COL_GROUP As String = "Group"
Public Const COL_TYPE As String = "Type"
Public Const COL_ESSENTIAL As String = "Essential"
Public Const COL_TAXTAG As String = "Tax Tag"
Public Const COL_SPLIT As String = "Split A %"
Public Const COL_SHARE_A As String = "Share A"
Public Const COL_SHARE_B As String = "Share B"
Public Const COL_REIMBURSE As String = "Reimbursable"
Public Const COL_NOTES As String = "Notes"
Public Const COL_SOURCE As String = "Source File"
Public Const COL_BATCH As String = "Batch"
Public Const COL_KEY As String = "Match Key"
Public Const COL_TAGGEDBY As String = "Tagged By"
Public Const COL_VIEW As String = "View Amount"

' Named ranges on the Settings sheet
Public Const NR_MODE As String = "HouseholdMode"
Public Const NR_PERSON_A As String = "PersonA"
Public Const NR_PERSON_B As String = "PersonB"
Public Const NR_SPLIT As String = "DefaultSplitA"
Public Const NR_PROVINCE As String = "Province"
Public Const NR_CONFIGURED As String = "Configured"
Public Const NR_TRANSFER_DAYS As String = "TransferWindowDays"
Public Const NR_DUPES As String = "SkipDuplicates"

' Settings values
Public Const MODE_SINGLE As String = "Single"
Public Const MODE_COUPLE As String = "Couple"
Public Const OWNER_JOINT As String = "Joint"

' Category names the macros rely on
Public Const CAT_UNCATEGORIZED As String = "Uncategorized"
Public Const CAT_TRANSFER As String = "Internal Transfer"
Public Const CAT_CARD_PAYMENT As String = "Credit Card Payment"

' Amount conventions used by the bank format profiles
Public Const MODE_SIGNED As String = "Signed"
Public Const MODE_SIGNED_FLIP As String = "Signed (flip)"
Public Const MODE_DEBIT_CREDIT As String = "Debit/Credit"

Public Const TAG_RULE As String = "Rule"
Public Const TAG_MANUAL As String = "Manual"
Public Const TAG_IMPORT As String = "Import"
Public Const TAG_TRANSFER As String = "Transfer match"
