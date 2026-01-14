# non-asset object types to skip in assets-and-layouts
$specialObjectTypes = @{
    "kbs"                           = "articles"
    "documents"                     = "articles"
    "companies"                     = "companies"
    "ipnetworks"                    = "ipam"
}

# junk props to ignore globally
$IgnoreFields = @(
    "mod_date",
    "IgnoreGlobalBlock",
    "Changes",
    "Deleted",
    "FirstName",
    "LastName",
    "Filesubmitdate",
    "Filesubmittime",
    "UserID",
    "SubmitBy",
    "SubmitDate",
    "Lock",
    "LockClients",
    "CustomerID",
    "ChangeCount",
    "ChangesCount",
    "IgnoreGlobalAllow",
    "ForeignContactID",
    "IgnoreGlobalBlock",
    "ForeignConfigObjectID",
    "ForeignConfigObjectTypeID",
    "is_file_indexed",
    "is_object_indexed_open_ai",
    "is_file_indexed_open_ai",
    "ClientID",
    "FileDeleted",
    "LastATActivityDate",
    "DeletedByID",
    "DeletedByName",
    "DeletedDate",
    "Deleted",
    "ModifiedDate",
    "deviceDeviceID",
    "objectGUID",
    "ForeignType",
    "InputTime",
    "FileSubmitDate",
    "CollectorID",
    "SiteID",
    "UserId",
    "ContactID",
    "DeletedByID",
    "DeletedByName",
    "ForeignConfigObjectID",
    "ForeignConfigObjectTypeID"
)

# password tables
$PassTypes = @("Passwords-FieldPasswords","Passwords-Devices","Passwords-LocalPasswords","Passwords-Accounts")

# tables to skip entirely in assets-and-layouts
$SkipTables= @(
)

# junk props to ignore for specific layouts
$LayoutSpecificIgnoreFields = @{
    "Sites" =@("ClientID","SiteId")
    "Devices"=@("Lock","LoclClients")
}

# naming for specific object types
$NameFields = @{
    "Contacts"="FullName"
    "ConfigItems"="ciName"
}

# which properties in company to use for which fields in hudu
$companyPropMap = @{
    nickname            = "Nickname"
    address_line_1      = "AddressLine1"
    address_line_2      = "AddressLine2"
    city                = "City"
    state               = "State"
    zip                 = "Zip"
    notes               = "Notes"
    country_name        = "CountryName"
    company_type        = "CompanyType"
    phone_number        = "PhoneNumber"
    fax_number          = "FaxNumber"
    website             = "NewWebsite"
}

# which article properties are associated with which parts (if handled in assets-and-layouts *advanced)
$articlesPropMap = @{
    DocumentId          = "Footer"
    KBID                = "Footer"
    documenttype        = "Folder"
    category            = "Folder"
    subcategory         = "SubFolder"
    description         = "Header"
    link                = "Header"
}

# depth at which to scan values to determine type in layout
$discernment = 4096

# typical system-wide identifiers
# AccountID
# AccountManagerID
# AccountTypeID
# addressID
# AddressID
# AgreementID
# AgreementSubTypeID
# AgreementTypeID
# CabinetFacilityID
# CabinetID
# CabinetSiteID
# ciDescription
# ClientID
# ClientTwoID
# ClientTypeID
# CollectorID
# ConfigObjectID
# ConfigObjectTypeID
# ContactID
# ContactTypeID
# CustomerID
# DeletedByID
# deviceDeviceID
# DeviceID
# DeviceTypeID
# DocumentID
# DocumentTypeID
# DSubcategoryID
# FacilityID
# ForeignClientID
# ForeignConfigObjectID
# ForeignConfigObjectTypeID
# ForeignContactID
# ForeignDeviceID
# ForeignDeviceTypeID
# ForeignSiteID
# gwID
# IPID
# KBCategoryID
# KBID
# KBSubCategoryID
# LogoID
# middleinitial
# MiddleInitial
# objectGUID
# ParentClientID
# RaidController
# ScoreboardID
# SecondaryUserID
# SiteID
# SiteTwoID
# SubnetID
# userid
# UserID
# VLANID

# typical object types
# Contacts                       
# Companies                      
# KBs                            
# Agreements                     
# Passwords-FieldPasswords       
# Sites                          
# ConfigItems                    
# Passwords-LocalPasswords       
# Accounts                       
# Devices                        
# Passwords-Devices              
# Documents                      
# IPNetworks                     
# Passwords-Accounts        