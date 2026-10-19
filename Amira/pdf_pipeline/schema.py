"""The results.csv metadata schema (column order matters) and attribute groupings."""

RESULT_COLUMNS = [
    "File Name",
    "File Size",
    "File Format",
    "Title",
    "Dataset Type",
    "Sub Dataset Type",
    "Data Status",
    "Description",
    "Created By",
    "Created For",
    "Remarks",
    "Top Depth",
    "Top Depth UOM",
    "Base Depth",
    "Base Depth UOM",
    "Spud Date",
    "Completion Date",
    "Transmittal Date",
    "Unique Wellbore Identifier",
    "Well Name",
    "Model",
    "FIELD_NAME",
    "CREATE_DATE",
    "DOCUMENT_TYPE",
    "PAGE_COUNT",
    "LOGGING_COMPANY",
    "SCALE_RATIO",
    "MEDIA_TYPE",
    "WELL_LOG_CLASS_ID",
    "LOG_TITLE",
    "REPORT_LOG_RUN",
    "TRIP_DATE",
    "ORIGINAL_FILE_NAME",
    "DECRYPT_KEY",
    "DECRYPTION_TYPE",
    "DIGITAL_SIZE_UOM",
    "BA_LONG_NAME_1",
    "BA_TYPE_1",
    "DATA_STORE_NAME",
    "DATA_STORE_TYPE",
    "BAR_CODE",
    "LOCATION_ID",
    "SOURCE",
    "ROW_QUALITY",
    "CHECKED_BY_BA_ID",
    "EXTRACTION STATUS",
]

EVIDENCE_COLUMNS = [
    "File Name",
    "Attribute",
    "Value",
    "Source Page(s)",
    "Source",
    "Extraction Method",
    "Validation Status",
    "Context",
    "Confidence",
]

# Never filled by the pipeline or by configuration: secrets must not be exported.
NEVER_FILLED = {"DECRYPT_KEY"}

# Set by the pipeline itself (filesystem / PDF structure); configuration may not override them.
SYSTEM_COLUMNS = {
    "File Name", "File Size", "File Format", "PAGE_COUNT", "ORIGINAL_FILE_NAME",
    "DECRYPTION_TYPE", "DIGITAL_SIZE_UOM", "EXTRACTION STATUS",
}

# Columns that config rules (constant_columns / classification_rules) may fill.
CONFIGURABLE_COLUMNS = [c for c in RESULT_COLUMNS if c not in NEVER_FILLED | SYSTEM_COLUMNS]

# Attributes expected to hold one value per document; more than one is flagged for review.
SINGLE_VALUED = {
    "Title", "CREATE_DATE", "Spud Date", "Completion Date", "Transmittal Date", "TRIP_DATE",
    "Top Depth", "Base Depth", "Unique Wellbore Identifier", "Model", "LOG_TITLE",
}

# Validation status codes written to evidence.csv
LABELLED = "LABELLED"          # value next to an explicit label in the document, e.g. "Spud Date: ..."
PATTERN = "PATTERN"            # matched a name/number pattern; review recommended
INFERRED = "INFERRED"          # heuristic (e.g. title block, first date on cover); review recommended
FILESYSTEM = "FILESYSTEM"      # read from the file system
PDF_STRUCTURE = "PDF_STRUCTURE"  # read from the PDF file structure (page count, encryption type)
PDF_METADATA = "PDF_METADATA"  # embedded PDF metadata; reference only, never copied into document fields
CONFIG = "CONFIG"              # set by a rule in dictionaries.json
CONFIRMED = "CONFIRMED"        # document text and PDF metadata agree
MULTIPLE = "MULTIPLE_VALUES"   # appended when a single-valued attribute had several distinct values
CANDIDATE = "CANDIDATE"        # seen but NOT exported (uncertain); kept for manual review
EXCLUDED = "EXCLUDED_SECTION"  # bibliography/reference text ignored for document-level metadata

# Evidence rows for columns that ended up empty
NOT_FOUND = "NOT_FOUND"                        # extractable, but no reliable evidence in this PDF
REFERENCE_DATA = "REFERENCE_DATA_REQUIRED"     # needs classification_rules / your reference lists
MANUAL = "MANUAL"                              # set in your data system or by a reviewer
NOT_ENCRYPTED = "NOT_ENCRYPTED"
NEVER_EXPORTED = "NEVER_EXPORTED"              # DECRYPT_KEY

REFERENCE_DATA_COLUMNS = {"Dataset Type", "Sub Dataset Type", "DOCUMENT_TYPE", "WELL_LOG_CLASS_ID"}
MANUAL_COLUMNS = {
    "Data Status", "Description", "Remarks", "MEDIA_TYPE", "BA_LONG_NAME_1", "BA_TYPE_1",
    "DATA_STORE_NAME", "DATA_STORE_TYPE", "BAR_CODE", "LOCATION_ID", "SOURCE", "ROW_QUALITY",
    "CHECKED_BY_BA_ID",
}

_CONFIDENCE = {
    FILESYSTEM: "HIGH", PDF_STRUCTURE: "HIGH", CONFIRMED: "HIGH", LABELLED: "HIGH", CONFIG: "HIGH",
    PATTERN: "MEDIUM", INFERRED: "LOW",
}


def confidence(validation: str) -> str:
    """HIGH / MEDIUM / LOW for exported values; '' for evidence that is not an exported value."""
    if MULTIPLE in validation:
        return "LOW"
    return _CONFIDENCE.get(validation.split(";")[0].strip(), "")
