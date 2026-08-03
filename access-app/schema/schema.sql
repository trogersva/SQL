-- ============================================================
-- New schema for the budget/fiscal Access app.
-- This is the human-readable reference copy. The actual tables
-- are created by running modSchemaBuilder.BuildSchema in Access
-- (see ../README.md) -- that VBA module executes the equivalent
-- DDL via DAO. Keep the two in sync if you change one.
-- ============================================================

-- ---------- Reference / lookup tables ----------

CREATE TABLE CostCenters (
    CostCenter   TEXT(10)  NOT NULL,
    CCName       TEXT(100),
    CONSTRAINT PK_CostCenters PRIMARY KEY (CostCenter)
);

CREATE TABLE BudgetObjectCodes (
    BOC              TEXT(10) NOT NULL,
    OC               TEXT(4),
    BOCDescription   TEXT(100),
    CONSTRAINT PK_BudgetObjectCodes PRIMARY KEY (BOC)
);

-- The old TABLE_Vendors was NOT a vendor registry -- it's a BOC -> default
-- vendor-number lookup (56 rows, BOC unique, only 7 distinct vendor numbers,
-- 41 of them the literal "MISCE"). Named and keyed accordingly.
CREATE TABLE BOCDefaultVendors (
    BOC            TEXT(10) NOT NULL,
    VendorNumber   TEXT(20),
    CONSTRAINT PK_BOCDefaultVendors PRIMARY KEY (BOC),
    CONSTRAINT FK_BOCDefaultVendors_BOC FOREIGN KEY (BOC) REFERENCES BudgetObjectCodes (BOC)
);

-- NOTE: StatusCode values inherited from the old database are truncated at
-- 33 characters ("Ordered (No Fiscal Action Require"). This is a lookup for
-- display, deliberately NOT an FK target -- a real IFCAP export's full
-- status text would not match these truncated keys.
CREATE TABLE StatusDescriptions (
    StatusCode    TEXT(40) NOT NULL,
    Description   MEMO,
    CONSTRAINT PK_StatusDescriptions PRIMARY KEY (StatusCode)
);

CREATE TABLE CostTransferTypes (
    TransferTypeCode  TEXT(20) NOT NULL,  -- was [ExcelTemplate] in the old schema
    TransInitials     TEXT(4),
    TransType         TEXT(4),
    ExpRevGLBG        TEXT(4),
    LogBookCode       TEXT(10),           -- 'EB', 'EW', 'EBGIP', 'EWGIP'
    CONSTRAINT PK_CostTransferTypes PRIMARY KEY (TransferTypeCode)
);

CREATE TABLE Employees (
    EmployeeID    AUTOINCREMENT,
    EmployeeName  TEXT(100) NOT NULL,
    Title         TEXT(100),
    Email         TEXT(100),
    Initials      TEXT(10),
    IsActive      YESNO,
    CONSTRAINT PK_Employees PRIMARY KEY (EmployeeID)
);

CREATE TABLE FundControlPoints (
    FCPNo                 LONG      NOT NULL,
    FY                    TEXT(4)   NOT NULL,
    ALLW                  TEXT(4),
    ACC                   TEXT(9),
    FCPName               TEXT(100),
    BudgetID              TEXT(22),
    Station               TEXT(3),
    Fund                  TEXT(10),
    OrganizationalChart   TEXT(100),
    ReportingCategory     TEXT(100),
    Active                YESNO,
    BudgetCall            YESNO,
    FCPType               TEXT(30),
    Notes                 MEMO,
    CONSTRAINT PK_FundControlPoints PRIMARY KEY (FCPNo, FY)
);

CREATE TABLE FCPAccess (
    FCPAccessID     AUTOINCREMENT,
    EmployeeID      LONG NOT NULL,
    FCPNo           LONG NOT NULL,
    FY              TEXT(4) NOT NULL,
    AccessLevel     TEXT(30),
    IsSpecialAccess YESNO,
    Notes           TEXT(255),
    CONSTRAINT PK_FCPAccess PRIMARY KEY (FCPAccessID),
    CONSTRAINT FK_FCPAccess_Employee FOREIGN KEY (EmployeeID) REFERENCES Employees (EmployeeID),
    CONSTRAINT FK_FCPAccess_FCP FOREIGN KEY (FCPNo, FY) REFERENCES FundControlPoints (FCPNo, FY)
);

-- ---------- Import automation infrastructure (new) ----------

CREATE TABLE ImportFileTypes (
    FileType         TEXT(20) NOT NULL,   -- e.g. 'F820', 'F826', 'AACS', 'SALT', 'ALLW', 'IFCAP', 'UnpaidPC'
    DisplayName      TEXT(100),
    Delimiter        TEXT(5),             -- e.g. '^' for FileMan caret-delimited exports, ',' for CSV
    TargetTable      TEXT(60),            -- 'FiscalSnapshots' or 'PendingOrders'
    StaleAfterDays   INTEGER,             -- drives the "this data hasn't been refreshed" alert
    CONSTRAINT PK_ImportFileTypes PRIMARY KEY (FileType)
);

CREATE TABLE ImportFieldMap (
    FileType      TEXT(20) NOT NULL,
    ColumnIndex   INTEGER  NOT NULL,      -- 1-based position in the source file
    TargetField   TEXT(60) NOT NULL,      -- column name in the target table
    DataType      TEXT(20),               -- 'TEXT', 'CURRENCY', 'DATE', 'LONG'
    CONSTRAINT PK_ImportFieldMap PRIMARY KEY (FileType, ColumnIndex),
    CONSTRAINT FK_ImportFieldMap_Type FOREIGN KEY (FileType) REFERENCES ImportFileTypes (FileType)
);

CREATE TABLE ImportLog (
    ImportID       AUTOINCREMENT,
    FileType       TEXT(20),
    FilePath       TEXT(255),
    RowsImported   LONG,
    RowsRejected   LONG,
    ImportedAt     DATETIME,
    ImportedBy     TEXT(100),
    CONSTRAINT PK_ImportLog PRIMARY KEY (ImportID),
    CONSTRAINT FK_ImportLog_Type FOREIGN KEY (FileType) REFERENCES ImportFileTypes (FileType)
);

CREATE TABLE ImportRejects (
    RejectID      AUTOINCREMENT,
    ImportID      LONG,
    LineNumber    LONG,
    RawLine       MEMO,
    ErrorMessage  TEXT(255),
    CONSTRAINT PK_ImportRejects PRIMARY KEY (RejectID),
    CONSTRAINT FK_ImportRejects_Import FOREIGN KEY (ImportID) REFERENCES ImportLog (ImportID)
);

-- ---------- Fiscal data (populated by modImportReports.bas) ----------
-- These two tables replace the original guess at a generic "delimited file"
-- FiscalSnapshots shape. Real F820/F20D/F826 downloads are fixed-width
-- printed reports (see docs/DESIGN.md "Real file formats"), not delimited
-- data, and they split naturally into two kinds of rows: point-in-time
-- balances (FiscalSnapshots) and individual ledger transactions
-- (FiscalTransactions).

CREATE TABLE FiscalSnapshots (
    SnapshotID     AUTOINCREMENT,
    FileType       TEXT(20)  NOT NULL,   -- 'F820', 'F20D', 'F826'
    Station        TEXT(3),
    BFYS           TEXT(10),             -- budget FY; multi-year funds print as "25 26"
    AO             TEXT(4),
    FundCode       TEXT(10),
    ClassCode      TEXT(20),             -- F820/F20D "Account Classification Code" / F826 "Orgn/Act"
    ClassCodeName  TEXT(100),
    Program        TEXT(10),             -- F826 only
    BalanceType    TEXT(10),             -- F820/F20D only: 'Beginning' / 'Ending'
    AmountType     TEXT(30)  NOT NULL,   -- 'BudgetCeiling'/'Obligations'/'UnobligatedBalance' (F820/F20D)
                                          -- or 'Budget'/'Obligations'/'Available' (F826)
    Amount         CURRENCY,
    RunDate        DATETIME,             -- report's own "Run Date", not when you happened to import it
    AsOfDate       DATETIME,             -- F820/F20D only
    ImportID       LONG,
    CONSTRAINT PK_FiscalSnapshots PRIMARY KEY (SnapshotID),
    CONSTRAINT FK_FiscalSnapshots_Type FOREIGN KEY (FileType) REFERENCES ImportFileTypes (FileType),
    CONSTRAINT FK_FiscalSnapshots_Import FOREIGN KEY (ImportID) REFERENCES ImportLog (ImportID)
);

CREATE TABLE FiscalTransactions (
    TransactionID       AUTOINCREMENT,
    FileType            TEXT(20) NOT NULL,  -- 'F820' or 'F20D'
    Station             TEXT(3),
    BFYS                TEXT(10),
    AO                  TEXT(4),
    FundCode            TEXT(10),
    ClassCode           TEXT(20),
    ClassCodeName       TEXT(100),
    DocID               TEXT(20),
    TransDate           DATETIME,
    Vendor               TEXT(60),
    BOC                 TEXT(10),
    SubBOC               TEXT(10),
    CostCenter           TEXT(10),
    CeilingAdjAmount     CURRENCY,
    ObligationAdjAmount  CURRENCY,
    RunDate              DATETIME,
    ImportID             LONG,
    CONSTRAINT PK_FiscalTransactions PRIMARY KEY (TransactionID),
    CONSTRAINT FK_FiscalTransactions_Import FOREIGN KEY (ImportID) REFERENCES ImportLog (ImportID)
);

-- ---------- Purchase card / order tracking ----------

CREATE TABLE PendingOrders (
    TransactionNumber   TEXT(30) NOT NULL,
    RecordType          TEXT(10) NOT NULL,  -- 'Pending','PFY','Returned'
    FCPNo               LONG,
    FY                  TEXT(4),
    PO                  TEXT(30),
    VendorNumber        TEXT(20),
    BOC                 TEXT(10),
    CostCenter          TEXT(10),
    ApprovalDate        DATETIME,
    Amount              CURRENCY,
    StatusCode          TEXT(40),
    LastUpdated         DATETIME,
    CONSTRAINT PK_PendingOrders PRIMARY KEY (TransactionNumber, RecordType),
    CONSTRAINT FK_PendingOrders_BOC FOREIGN KEY (BOC) REFERENCES BudgetObjectCodes (BOC),
    CONSTRAINT FK_PendingOrders_CC FOREIGN KEY (CostCenter) REFERENCES CostCenters (CostCenter)
    -- No FK on VendorNumber: there is no authoritative vendor registry to
    -- point at (see BOCDefaultVendors above). No FK on StatusCode either,
    -- since the inherited status values are truncated.
);

CREATE TABLE PendingAdjustments (
    AdjustmentID       AUTOINCREMENT,
    TransactionNumber  TEXT(30) NOT NULL,
    OriginalAmount     CURRENCY,
    AdjustmentAmount   CURRENCY,
    AdjustedTotal      CURRENCY,
    ReversalDate       DATETIME,
    Notes              TEXT(255),
    CONSTRAINT PK_PendingAdjustments PRIMARY KEY (AdjustmentID)
);

-- ---------- Journal voucher / cost transfer logbook ----------

CREATE TABLE JournalVouchers (
    JVID              AUTOINCREMENT,
    DocID             TEXT(30),
    TransferTypeCode  TEXT(20),
    LogBookCode       TEXT(10) NOT NULL,   -- 'EB','EW','EBGIP','EWGIP'
    TransDate         DATETIME,
    ShortDescription  TEXT(100),
    Purpose           MEMO,
    Amount            CURRENCY,
    EnteredBy         LONG,
    ApprovedBy        LONG,
    ApprovedDate      DATETIME,
    CONSTRAINT PK_JournalVouchers PRIMARY KEY (JVID),
    CONSTRAINT FK_JV_TransferType FOREIGN KEY (TransferTypeCode) REFERENCES CostTransferTypes (TransferTypeCode),
    CONSTRAINT FK_JV_EnteredBy FOREIGN KEY (EnteredBy) REFERENCES Employees (EmployeeID),
    CONSTRAINT FK_JV_ApprovedBy FOREIGN KEY (ApprovedBy) REFERENCES Employees (EmployeeID)
);

-- ---------- Budget ledger ----------

CREATE TABLE BudgetTransactions (
    TransID           AUTOINCREMENT,
    TransDate         DATETIME,
    FCPNo             LONG,
    FY                TEXT(4),
    TDA               LONG,
    CostCenter        TEXT(10),
    BOC               TEXT(10),
    Description       TEXT(150),
    Qtr1Amt           CURRENCY,
    Qtr2Amt           CURRENCY,
    Qtr3Amt           CURRENCY,
    Qtr4Amt           CURRENCY,
    IsRecurring       YESNO,
    IsPaperOnly       YESNO,
    IsFMSOnly         YESNO,
    Notes             MEMO,
    CONSTRAINT PK_BudgetTransactions PRIMARY KEY (TransID),
    CONSTRAINT FK_BudgetTrans_CC FOREIGN KEY (CostCenter) REFERENCES CostCenters (CostCenter),
    CONSTRAINT FK_BudgetTrans_BOC FOREIGN KEY (BOC) REFERENCES BudgetObjectCodes (BOC)
);

CREATE TABLE BudgetAdjustments (
    AdjustmentID   AUTOINCREMENT,
    AdjDate        DATETIME,
    FCPNo          LONG,
    FY             TEXT(4),
    Qtr            TEXT(4),
    CostCenter     TEXT(10),
    BOC            TEXT(10),
    AdjType        TEXT(30),
    Amount         CURRENCY,
    Description    TEXT(150),
    CONSTRAINT PK_BudgetAdjustments PRIMARY KEY (AdjustmentID),
    CONSTRAINT FK_BudgetAdj_CC FOREIGN KEY (CostCenter) REFERENCES CostCenters (CostCenter),
    CONSTRAINT FK_BudgetAdj_BOC FOREIGN KEY (BOC) REFERENCES BudgetObjectCodes (BOC)
);

-- ---------- Tasks / reminders ----------

CREATE TABLE Tasks (
    TaskID            AUTOINCREMENT,
    TaskName          TEXT(100) NOT NULL,
    Description       MEMO,
    IsRecurring       YESNO,
    RecurrenceDays    INTEGER,
    DueDate           DATETIME,
    CompletedDate     DATETIME,
    CompletedBy       LONG,
    CONSTRAINT PK_Tasks PRIMARY KEY (TaskID),
    CONSTRAINT FK_Tasks_CompletedBy FOREIGN KEY (CompletedBy) REFERENCES Employees (EmployeeID)
);

-- ---------- Attachments (metadata only -- no embedded blobs) ----------

CREATE TABLE Attachments (
    AttachmentID   AUTOINCREMENT,
    Description    TEXT(255),
    FilePath       TEXT(255) NOT NULL,   -- path on disk / SharePoint, not embedded bytes
    CONSTRAINT PK_Attachments PRIMARY KEY (AttachmentID)
);

-- ---------- App log ----------

CREATE TABLE AppLog (
    LogID       AUTOINCREMENT,
    LoggedAt    DATETIME,
    Source      TEXT(60),
    Message     MEMO,
    CONSTRAINT PK_AppLog PRIMARY KEY (LogID)
);
