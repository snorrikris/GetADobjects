using System;
using System.Data;

public class TableColDef
{
    public string ColName;
    public Type datatype;
    public string ADpropName;
    public string OPtype;

    public TableColDef(string ColName, Type datatype, string ADpropName, string OPtype)
    {
        this.ColName = ColName;
        this.datatype = datatype;
        this.ADpropName = ADpropName;
        this.OPtype = OPtype;
    }
}

public class ADcolsTable
{
    public TableColDef[] collist;
    public bool IsUser, IsContact, IsComputer, IsGroup, IsWellKnownSIDs;

    public ADcolsTable(string ADfilter)
    {
        IsUser = IsContact = IsComputer = IsGroup = IsWellKnownSIDs = false;
        // Use ADfilter parameter to determine the type of AD objects wanted.
        if (ADfilter.Contains("user"))
            MakeUserColList();
        else if (ADfilter.Contains("contact"))
            MakeContactColList();
        else if (ADfilter.Contains("computer"))
            MakeComputerColList();
        else if (ADfilter.Contains("group"))
            MakeGroupColList();
        else if (ADfilter.Contains("objectSID=S-1-5-4"))
            MakeWellKnownSIDsList();
        else
            collist = new TableColDef[0];
    }

    private void MakeUserColList()
    {
        IsUser = true;
        collist = new TableColDef[]  
        {
            // COPY CODE from "AD Column map.xlsx" if list changes.
            // Copy/Paste all cells from "ColListDef" column in Users sheet.
            new TableColDef("GivenName", typeof(String), "givenname","Adprop"),
            new TableColDef("Initials", typeof(String), "initials","Adprop"),
            new TableColDef("Surname", typeof(String), "sn","Adprop"),
            new TableColDef("DisplayName", typeof(String), "displayname","Adprop"),
            new TableColDef("Description", typeof(String), "description","Adprop"),
            new TableColDef("Office", typeof(String), "physicalDeliveryOfficeName","Adprop"),
            new TableColDef("OfficePhone", typeof(String), "telephonenumber","Adprop"),
            new TableColDef("EmailAddress", typeof(String), "mail","Adprop"),
            new TableColDef("HomePage", typeof(String), "wwwhomepage","Adprop"),
            new TableColDef("StreetAddress", typeof(String), "streetaddress","Adprop"),
            new TableColDef("POBox", typeof(String), "postofficebox","Adprop"),
            new TableColDef("City", typeof(String), "l","Adprop"),
            new TableColDef("State", typeof(String), "st","Adprop"),
            new TableColDef("PostalCode", typeof(String), "postalcode","Adprop"),
            new TableColDef("Country", typeof(String), "co","Adprop"),
            new TableColDef("HomePhone", typeof(String), "homephone","Adprop"),
            new TableColDef("Pager", typeof(String), "pager","Adprop"),
            new TableColDef("MobilePhone", typeof(String), "mobile","Adprop"),
            new TableColDef("Fax", typeof(String), "facsimileTelephoneNumber","Adprop"),
            new TableColDef("Title", typeof(String), "title","Adprop"),
            new TableColDef("Department", typeof(String), "department","Adprop"),
            new TableColDef("Company", typeof(String), "company","Adprop"),
            new TableColDef("Manager", typeof(String), "manager","Adprop"),
            new TableColDef("EmployeeID", typeof(String), "employeeid","Adprop"),
            new TableColDef("EmployeeNumber", typeof(String), "employeenumber","Adprop"),
            new TableColDef("Division", typeof(String), "division","Adprop"),
            new TableColDef("Enabled", typeof(Boolean), "ACCOUNTDISABLE","UAC"),
            new TableColDef("LockedOut", typeof(Boolean), "LOCKOUT","UAC"),
            new TableColDef("MNSLogonAccount", typeof(Boolean), "MNS_LOGON_ACCOUNT","UAC"),
            new TableColDef("CannotChangePassword", typeof(Boolean), "PASSWD_CANT_CHANGE","UAC"),
            new TableColDef("PasswordExpired", typeof(Boolean), "PASSWORD_EXPIRED","UAC"),
            new TableColDef("PasswordNeverExpires", typeof(Boolean), "DONT_EXPIRE_PASSWD","UAC"),
            new TableColDef("PasswordNotRequired", typeof(Boolean), "PASSWD_NOTREQD","UAC"),
            new TableColDef("SmartcardLogonRequired", typeof(Boolean), "SMARTCARD_REQUIRED","UAC"),
            new TableColDef("DoesNotRequirePreAuth", typeof(Boolean), "DONT_REQ_PREAUTH","UAC"),
            new TableColDef("AllowReversiblePasswordEncryption", typeof(Boolean), "ENCRYPTED_TEXT_PWD_ALLOWED","UAC"),
            new TableColDef("AccountNotDelegated", typeof(Boolean), "NOT_DELEGATED","UAC"),
            new TableColDef("TrustedForDelegation", typeof(Boolean), "TRUSTED_FOR_DELEGATION","UAC"),
            new TableColDef("TrustedToAuthForDelegation", typeof(Boolean), "TRUSTED_TO_AUTH_FOR_DELEGATION","UAC"),
            new TableColDef("UseDESKeyOnly", typeof(Boolean), "USE_DES_KEY_ONLY","UAC"),
            new TableColDef("UserMustChangePasswordAtNextLogon", typeof(Boolean), "does_not_exist_in_AD","nop"),
            new TableColDef("HomedirRequired", typeof(Boolean), "HOMEDIR_REQUIRED","UAC"),
            new TableColDef("LastBadPasswordAttempt", typeof(DateTime), "badpasswordtime","filetime"),
            new TableColDef("BadLogonCount", typeof(Int32), "badpwdcount","Adprop"),
            new TableColDef("LastLogonDate", typeof(DateTime), "lastlogon","filetime"),
            new TableColDef("LogonCount", typeof(Int32), "logoncount","Adprop"),
            new TableColDef("PasswordLastSet", typeof(DateTime), "pwdlastset","filetime"),
            new TableColDef("PasswordExpiryTime", typeof(DateTime), "msDS-UserPasswordExpiryTimeComputed","filetime"),
            new TableColDef("AccountLockoutTime", typeof(DateTime), "lockouttime","filetime"),
            new TableColDef("AccountExpirationDate", typeof(DateTime), "accountexpires","filetime"),
            new TableColDef("LogonWorkstations", typeof(String), "userworkstations","Adprop"),
            new TableColDef("HomeDirectory", typeof(String), "homedirectory","Adprop"),
            new TableColDef("HomeDrive", typeof(String), "homedrive","Adprop"),
            new TableColDef("ProfilePath", typeof(String), "profilepath","Adprop"),
            new TableColDef("ScriptPath", typeof(String), "scriptpath","Adprop"),
            new TableColDef("userAccountControl", typeof(Int32), "useraccountcontrol","Adprop"),
            new TableColDef("PrimaryGroupID", typeof(Int32), "primarygroupid","Adprop"),
            new TableColDef("Name", typeof(String), "name","Adprop"),
            new TableColDef("CN", typeof(String), "cn","Adprop"),
            new TableColDef("UserPrincipalName", typeof(String), "userprincipalname","Adprop"),
            new TableColDef("SamAccountName", typeof(String), "samaccountname","Adprop"),
            new TableColDef("DistinguishedName", typeof(String), "distinguishedname","Adprop"),
            new TableColDef("Created", typeof(DateTime), "whencreated","Adprop"),
            new TableColDef("Modified", typeof(DateTime), "whenchanged","Adprop"),
            new TableColDef("ObjectCategory", typeof(String), "objectcategory","Adprop"),
            new TableColDef("ObjectClass", typeof(String), "SchemaClassName","ObjClass"),
            new TableColDef("SID", typeof(String), "objectsid","SID"),
            new TableColDef("ObjectGUID", typeof(Guid), "Guid","ObjGuid"),
            new TableColDef("ManagerGUID", typeof(Guid), "does_not_exist_in_AD","nop")
        };
    }

    private void MakeContactColList()
    {
        IsContact = true;
        collist = new TableColDef[]
        {
            // COPY CODE from "AD Column map.xlsx" if list changes.
            // Copy/Paste all cells from "ColListDef" column in Contacts sheet.
            new TableColDef("GivenName", typeof(String), "givenname","Adprop"),
            new TableColDef("Initials", typeof(String), "initials","Adprop"),
            new TableColDef("Surname", typeof(String), "sn","Adprop"),
            new TableColDef("DisplayName", typeof(String), "displayname","Adprop"),
            new TableColDef("Description", typeof(String), "description","Adprop"),
            new TableColDef("Office", typeof(String), "physicalDeliveryOfficeName","Adprop"),
            new TableColDef("OfficePhone", typeof(String), "telephonenumber","Adprop"),
            new TableColDef("EmailAddress", typeof(String), "mail","Adprop"),
            new TableColDef("HomePage", typeof(String), "wwwhomepage","Adprop"),
            new TableColDef("StreetAddress", typeof(String), "streetaddress","Adprop"),
            new TableColDef("POBox", typeof(String), "postofficebox","Adprop"),
            new TableColDef("City", typeof(String), "l","Adprop"),
            new TableColDef("State", typeof(String), "st","Adprop"),
            new TableColDef("PostalCode", typeof(String), "postalcode","Adprop"),
            new TableColDef("Country", typeof(String), "co","Adprop"),
            new TableColDef("HomePhone", typeof(String), "homephone","Adprop"),
            new TableColDef("Pager", typeof(String), "pager","Adprop"),
            new TableColDef("MobilePhone", typeof(String), "mobile","Adprop"),
            new TableColDef("Fax", typeof(String), "facsimileTelephoneNumber","Adprop"),
            new TableColDef("Title", typeof(String), "title","Adprop"),
            new TableColDef("Department", typeof(String), "department","Adprop"),
            new TableColDef("Company", typeof(String), "company","Adprop"),
            new TableColDef("Manager", typeof(String), "manager","Adprop"),
            new TableColDef("EmployeeID", typeof(String), "employeeid","Adprop"),
            new TableColDef("EmployeeNumber", typeof(String), "employeenumber","Adprop"),
            new TableColDef("Division", typeof(String), "division","Adprop"),
            new TableColDef("DistinguishedName", typeof(String), "distinguishedname","Adprop"),
            new TableColDef("Name", typeof(String), "name","Adprop"),
            new TableColDef("CN", typeof(String), "cn","Adprop"),
            new TableColDef("Created", typeof(DateTime), "whencreated","Adprop"),
            new TableColDef("Modified", typeof(DateTime), "whenchanged","Adprop"),
            new TableColDef("ObjectCategory", typeof(String), "objectcategory","Adprop"),
            new TableColDef("ObjectClass", typeof(String), "SchemaClassName","ObjClass"),
            new TableColDef("ObjectGUID", typeof(Guid), "Guid","ObjGuid")
        };
    }

    private void MakeComputerColList()
    {
        IsComputer = true;
        collist = new TableColDef[]
        {
            // COPY CODE from "AD Column map.xlsx" if list changes.
            // Copy/Paste all cells from "ColListDef" column in Computers sheet.
            new TableColDef("Name", typeof(String), "name","Adprop"),
            new TableColDef("DNSHostName", typeof(String), "dnshostname","Adprop"),
            new TableColDef("Description", typeof(String), "description","Adprop"),
            new TableColDef("Location", typeof(String), "location","Adprop"),
            new TableColDef("OperatingSystem", typeof(String), "operatingsystem","Adprop"),
            new TableColDef("OperatingSystemVersion", typeof(String), "operatingsystemversion","Adprop"),
            new TableColDef("OperatingSystemServicePack", typeof(String), "operatingsystemservicepack","Adprop"),
            new TableColDef("ManagedBy", typeof(String), "managedby","Adprop"),
            new TableColDef("Enabled", typeof(Boolean), "ACCOUNTDISABLE","UAC"),
            new TableColDef("LockedOut", typeof(Boolean), "LOCKOUT","UAC"),
            new TableColDef("MNSLogonAccount", typeof(Boolean), "MNS_LOGON_ACCOUNT","UAC"),
            new TableColDef("CannotChangePassword", typeof(Boolean), "PASSWD_CANT_CHANGE","UAC"),
            new TableColDef("PasswordExpired", typeof(Boolean), "PASSWORD_EXPIRED","UAC"),
            new TableColDef("PasswordNeverExpires", typeof(Boolean), "DONT_EXPIRE_PASSWD","UAC"),
            new TableColDef("PasswordNotRequired", typeof(Boolean), "PASSWD_NOTREQD","UAC"),
            new TableColDef("SmartcardLogonRequired", typeof(Boolean), "SMARTCARD_REQUIRED","UAC"),
            new TableColDef("DoesNotRequirePreAuth", typeof(Boolean), "DONT_REQ_PREAUTH","UAC"),
            new TableColDef("AllowReversiblePasswordEncryption", typeof(Boolean), "ENCRYPTED_TEXT_PWD_ALLOWED","UAC"),
            new TableColDef("AccountNotDelegated", typeof(Boolean), "NOT_DELEGATED","UAC"),
            new TableColDef("TrustedForDelegation", typeof(Boolean), "TRUSTED_FOR_DELEGATION","UAC"),
            new TableColDef("TrustedToAuthForDelegation", typeof(Boolean), "TRUSTED_TO_AUTH_FOR_DELEGATION","UAC"),
            new TableColDef("UseDESKeyOnly", typeof(Boolean), "USE_DES_KEY_ONLY","UAC"),
            new TableColDef("LastBadPasswordAttempt", typeof(DateTime), "badpasswordtime","filetime"),
            new TableColDef("BadLogonCount", typeof(Int32), "badpwdcount","Adprop"),
            new TableColDef("LastLogonDate", typeof(DateTime), "lastlogon","filetime"),
            new TableColDef("logonCount", typeof(Int32), "logoncount","Adprop"),
            new TableColDef("PasswordLastSet", typeof(DateTime), "pwdlastset","filetime"),
            new TableColDef("AccountLockoutTime", typeof(DateTime), "lockouttime","filetime"),
            new TableColDef("AccountExpirationDate", typeof(DateTime), "accountexpires","filetime"),
            new TableColDef("Created", typeof(DateTime), "whenCreated","Adprop"),
            new TableColDef("Modified", typeof(DateTime), "whenChanged","Adprop"),
            new TableColDef("CN", typeof(String), "cn","Adprop"),
            new TableColDef("DisplayName", typeof(String), "displayname","Adprop"),
            new TableColDef("DistinguishedName", typeof(String), "distinguishedname","Adprop"),
            new TableColDef("PrimaryGroupID", typeof(Int32), "primarygroupid","Adprop"),
            new TableColDef("SamAccountName", typeof(String), "samaccountname","Adprop"),
            new TableColDef("userAccountControl", typeof(Int32), "useraccountcontrol","Adprop"),
            new TableColDef("ObjectCategory", typeof(String), "objectcategory","Adprop"),
            new TableColDef("ObjectClass", typeof(String), "SchemaClassName","ObjClass"),
            new TableColDef("SID", typeof(String), "objectsid","SID"),
            new TableColDef("ObjectGUID", typeof(Guid), "Guid","ObjGuid")
        };
    }

    private void MakeGroupColList()
    {
        IsGroup = true;
        collist = new TableColDef[]
        {
            // COPY CODE from "AD Column map.xlsx" if list changes.
            // Copy/Paste all cells from "ColListDef" column in Groups sheet.
            new TableColDef("Name", typeof(String), "name","Adprop"),
            new TableColDef("GroupCategory", typeof(String), "grouptype","GrpCat"),
            new TableColDef("GroupScope", typeof(String), "grouptype","GrpScope"),
            new TableColDef("Description", typeof(String), "description","Adprop"),
            new TableColDef("EmailAddress", typeof(String), "mail","Adprop"),
            new TableColDef("ManagedBy", typeof(String), "managedby","Adprop"),
            new TableColDef("DistinguishedName", typeof(String), "distinguishedname","Adprop"),
            new TableColDef("DisplayName", typeof(String), "displayname","Adprop"),
            new TableColDef("Created", typeof(DateTime), "whenCreated","Adprop"),
            new TableColDef("Modified", typeof(DateTime), "whenChanged","Adprop"),
            new TableColDef("SamAccountName", typeof(String), "samaccountname","Adprop"),
            new TableColDef("ObjectCategory", typeof(String), "objectcategory","Adprop"),
            new TableColDef("ObjectClass", typeof(String), "SchemaClassName","ObjClass"),
            new TableColDef("SID", typeof(String), "objectsid","SID"),
            new TableColDef("ObjectGUID", typeof(Guid), "Guid","ObjGuid")
        };
    }

    private void MakeWellKnownSIDsList()
    {
        IsWellKnownSIDs = true;
        collist = new TableColDef[] 
        {
            // COPY CODE from "AD Column map.xlsx" if list changes.
            // Copy/Paste all cells from "ColListDef" column in WellKnownSIDs sheet.
            new TableColDef("Name", typeof(String), "name", "Adprop"),
            new TableColDef("Description", typeof(String), "description", "PostOp"),
            new TableColDef("DistinguishedName", typeof(String), "distinguishedname", "Adprop"),
            new TableColDef("DisplayName", typeof(String), "displayname", "PostOp"),
            new TableColDef("ObjectCategory", typeof(String), "objectcategory", "Adprop"),
            new TableColDef("ObjectClass", typeof(String), "SchemaClassName", "ObjClass"),
            new TableColDef("SID", typeof(String), "objectsid", "SID"),
            new TableColDef("ObjectGUID", typeof(Guid), "Guid", "ObjGuid")
        };
    }

    public DataTable CreateTable()
    {
        DataTable tbl = new DataTable();
        foreach (TableColDef col in collist)
        {
            tbl.Columns.Add(col.ColName, col.datatype);
        }
        return tbl;
    }
}

