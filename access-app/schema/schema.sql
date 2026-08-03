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

CREATE TABLE Vendors (
    VendorNumber   TEXT(20) NOT NULL,
    VendorName     TEXT(150),
    DefaultBOC     TEXT(10),
    Notes          TEXT(255),
    CONSTRAINT PK_Vendors PRIMARY KEY (VendorNumber),
    CONSTRAINT FK_Vendors_BOC FOREIGN KEY (DefaultBOC) REFERENCES BudgetObjectCodes (BOC)
);

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

-- ---------- Fiscal data (populated by the import engine) ----------

CREATE TABLE FiscalSnapshots (
    SnapshotID     AUTOINCREMENT,
    FileType       TEXT(20)  NOT NULL,
    FCPNo          LONG,
    FY             TEXT(4),
    Fund           TEXT(10),
    DocID          TEXT(30),
    VendorNumber   TEXT(20),
    BOC            TEXT(10),
    CostCenter     TEXT(10),
    AmountType     TEXT(30)  NOT NULL,    -- 'Ceiling','Obligated','Budget','Unobligated','Uncommitted','Pending','Undistributed'
    Amount         CURRENCY,
    SnapshotDate   DATETIME,
    ImportID       LONG,
    CONSTRAINT PK_FiscalSnapshots PRIMARY KEY (SnapshotID),
    CONSTRAINT FK_FiscalSnapshots_Type FOREIGN KEY (FileType) REFERENCES ImportFileTypes (FileType),
    CONSTRAINT FK_FiscalSnapshots_Import FOREIGN KEY (ImportID) REFERENCES ImportLog (ImportID)
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
    CONSTRAINT FK_PendingOrders_Vendor FOREIGN KEY (VendorNumber) REFERENCES Vendors (VendorNumber),
    CONSTRAINT FK_PendingOrders_BOC FOREIGN KEY (BOC) REFERENCES BudgetObjectCodes (BOC),
    CONSTRAINT FK_PendingOrders_CC FOREIGN KEY (CostCenter) REFERENCES CostCenters (CostCenter),
    CONSTRAINT FK_PendingOrders_Status FOREIGN KEY (StatusCode) REFERENCES StatusDescriptions (StatusCode)
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
