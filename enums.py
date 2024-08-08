from enum import Enum


class FileExtension(Enum):
    XLS = ".xls"
    XLSX = ".xlsx"
    ZIP = ".zip"
    INI = ".ini"


class ValidationType(Enum):
    """Enumeration for different types of validations."""
    INI = "INI"
    BOM = "BOM"
    CONTENT = "CONTENT"
    ATTRIBUTES = "ATTRIBUTES"
    AM = "AM"
    SCHEDULE = "SCHEDULE"


class ElementType(Enum):
    RADAR = "Radar"
    CAMERA = "Camera"
    BRACKET = "Bracket"
    COVER = "Cover"
    RESISTOR = "Resistor"
    OTHER = "Other"


class LetterIdentifier(Enum):
    X = "X"
    R = "R"
    N = "N"
    SC = "SC"


class LifeCycleStatus(Enum):
    WORKING = "Working"
    PROPOSED = "Proposed"
    REJECTED = "Rejected"
    APPROVED = "Approved"
    RELEASED = "Released"


class DocumentType(Enum):
    ASSEMBLY_DRAWING = "Assembly Drawing"
    INSTALLATION_DRAWING = "Installation Drawing"
    SERVICE_DATA = "Service Data"


class BomScheduleKey(Enum):
    PART_NUMBER = "PART NUMBER"
    RADAR_PART_NUMBER = "RADAR PART NUMBER"
    SW_PART_NUMBER = "SW PART NUMBER"
    SW_VERSION = "SW VERSION"
    CONFIG_INI_DATA_SET = "CONFIG INI DATA SET"
    PRODUCT_ID_LABEL_PART_NUMBER = "PRODUCT ID LABEL PART NUMBER"
    PART_NUMBER_LABEL = "PART NUMBER LABEL"
    COVER = "COVER"
    JUMPER_HARNESS = "JUMPER HARNESS"
    BOOT_SOFTWARE_PART_NUMBER = "BOOT SOFTWARE PART NUMBER"
    MAIN_SOFTWARE_PART_NUMBER = "MAIN SOFTWARE PART NUMBER"
    SERVICE_LABEL = "SERVICE LABEL"
    REMAN_CAMERA_PART_NUMBER = "REMAN CAMERA PART NUMBER"
    REMAN_LABEL = "REMAN LABEL"
    LABEL = "LABEL"
    BRACKET = "BRACKET"
    CAMERA_PART_NUMBER = "CAMERA PART NUMBER"


class DropdownType(Enum):
    """Enumeration for different dropdown selections."""
    FLC_20 = "FLC-20"
    FLR_21 = "FLR-21"
    FLC_25 = "FLC-25"
    FLR_25 = "FLR-25"
    FLR_25_BRACKET = "FLR-25 Bracket"
    FLC_25_BRACKET = "FLC-25 Bracket"
    FLC_25_COVER = "FLC-25 Cover"


class TypeOption(Enum):
    COMPARE_INI_FILE = "Compare INI file"
    BOM_STRUCTURE = "BOM Structure"
    DOCUMENT_VALIDATION = "Document Validation"
    AFTERMARKET_PARTS = "AfterMarket Parts"
    ATTRIBUTES_VALIDATION = "Attributes Validation"


class ScheduleOption(Enum):
    Z068947 = "Z068947-SCHEDULE, US, American English, 008"
    Z069022 = "Z069022-SCHEDULE, US, American English, 022"
    Z071718 = "Z071718-SCHEDULE, US, American English, 008"
    Z073299 = "Z073299-SCHEDULE, US, American English, 009"
    Z080473 = "Z080473-SCHEDULE, US, American English, 023"
    Z214761 = "Z214761-SCHEDULE, US, American English, 011"
    Z251731 = "Z251731-SCHEDULE, US, American English, 002"
    Z251849 = "Z251849-SCHEDULE, US, American English, 002"
    Z258301 = "Z258301-SCHEDULE, US, American English, 002"
    Z259492 = "Z259492-SCHEDULE, US, American English, 000"
    Z259704 = "Z259704-SCHEDULE, US, American English, 000"
    Z267482 = "Z267482-SCHEDULE, US, American English, 003"
    Z268950 = "Z268950-SCHEDULE, US, American English, 006"
    Z296801 = "Z296801-SCHEDULE, US, American English, 000"
