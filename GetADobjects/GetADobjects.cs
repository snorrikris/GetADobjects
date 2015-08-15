using System;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using Microsoft.SqlServer.Server;
using System.IO;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Security.Principal;
using System.Xml;
using System.Collections.Generic;

/*
 * Must run this on the database
CREATE ASSEMBLY  
    [System.DirectoryServices.AccountManagement] from 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.DirectoryServices.AccountManagement.dll'
    with permission_set = UNSAFE --Fails if not 64 on 64 bit machines 
GO
*/

public partial class StoredProcedures
{
    [Microsoft.SqlServer.Server.SqlProcedure]
    public static void clr_GetADobjects(SqlString ADpath, SqlString ADfilter, out SqlXml MemberList)
    {
        // Filter syntax: https://msdn.microsoft.com/en-us/library/aa746475(v=vs.85).aspx
        // AD attributes: https://msdn.microsoft.com/en-us/library/ms675089(v=vs.85).aspx

        MemberList = new SqlXml();

        System.IO.StreamWriter file = CreateLogFile();

        SearchResultCollection results = null;
        try
        {
            XmlDocument doc = new XmlDocument();
            XmlDeclaration xmlDeclaration = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
            XmlElement root = doc.DocumentElement;
            doc.InsertBefore(xmlDeclaration, root);
            XmlElement body = doc.CreateElement(string.Empty, "body", string.Empty);
            doc.AppendChild(body);

            ADcolsTable TblData = new ADcolsTable((string)ADfilter);
            DataTable tbl = TblData.CreateTable();
            DataRow row;

            DirectoryEntry entry = new DirectoryEntry((string)ADpath);
            DirectorySearcher searcher = new DirectorySearcher(entry);
            searcher.Filter = (string)ADfilter;
            searcher.PageSize = 500;

            results = searcher.FindAll();
            foreach (SearchResult searchResult in results)
            {
                DirectoryEntry item = searchResult.GetDirectoryEntry();
                row = tbl.NewRow();

                UACflags Item_UAC_flags = null;
                PropertyValueCollection ADGroupType = null;

                for (int i = 0; i < TblData.collist.Length; i++)
                {
                    TableColDefEx coldef = TblData.collist[i];
                    switch(coldef.OPtype)
                    {
                        case "Adprop":
                            PropertyValueCollection prop = GetADproperty(item, coldef.ADpropName);
                            if (prop != null)
                                row[i] = prop.Value;
                            break;

                        case "UAC":
                            if (Item_UAC_flags == null)
                                Item_UAC_flags = new UACflags(Get_userAccountControl(item));
                            row[i] = Item_UAC_flags.GetFlag(coldef.ADpropName);
                            break;

                        case "ObjClass":
                            row[i] = item.SchemaClassName;
                            break;

                        case "ObjGuid":
                            row[i] = item.Guid;
                            break;

                        case "filetime":
                            Int64 time = GetFileTime(searchResult, coldef.ADpropName);
                            if(time > 0 && time != 0x7fffffffffffffff)
                            {
                                row[i] = DateTime.FromFileTimeUtc(time);
                            }
                            break;

                        case "SID":
                            row[i] = GetSID(item, coldef.ADpropName);
                            break;

                        case "GrpCat":
                            if (ADGroupType == null)
                                ADGroupType = GetADproperty(item, "grouptype");
                            row[i] = GetGroupCategory(ADGroupType);
                            break;

                        case "GrpScope":
                            if (ADGroupType == null)
                                ADGroupType = GetADproperty(item, "grouptype");
                            row[i] = GetGroupScope(ADGroupType);
                            break;
                    }
                }
                tbl.Rows.Add(row);

                // Save group members into the Xml document.
                if (TblData.IsGroup && item.Properties.Contains("member"))
                {
                    PropertyValueCollection coll = GetADproperty(item, "member");
                    string parent = (string)row["distinguishedname"];
                    SaveGroupMembersToXml(doc, body, parent, coll);
                }
            }
            DataSetUtilities.SendDataTable(tbl);

            using (XmlNodeReader xnr = new XmlNodeReader(doc))
            {
                MemberList = new SqlXml(xnr);
            }
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            System.Runtime.InteropServices.COMException exception = new System.Runtime.InteropServices.COMException();
            file.WriteLine("COMException: " + exception);
        }
        catch (InvalidOperationException)
        {
            InvalidOperationException InvOpEx = new InvalidOperationException();
            file.WriteLine("InvalidOperationException: " + InvOpEx.Message);
        }
        catch (NotSupportedException)
        {
            NotSupportedException NotSuppEx = new NotSupportedException();
            file.WriteLine("NotSupportedException: " + NotSuppEx.Message);
        }
        catch (Exception ex)
        {
            file.WriteLine("Exception: " + ex.Message);
        }
        finally
        {
            if (null != results)
            {
                results.Dispose();  // To prevent memory leaks, always call 
                results = null;     // SearchResultCollection.Dispose() manually.
            }
        }
        file.Close();
    }   // endof: clr_GetADobjects

    public static System.IO.StreamWriter CreateLogFile()
    {
        System.IO.StreamWriter file = null;
        try
        {
            string folder = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);

            // Combine the base folder with your specific folder....
            string specificFolder = Path.Combine(folder, "GetADobjects");

            // Check if folder exists and if not, create it
            if (!Directory.Exists(specificFolder))
                Directory.CreateDirectory(specificFolder);

            string filename = Path.Combine(specificFolder, "Log.txt");

            file = new System.IO.StreamWriter(filename);
        }
        catch(Exception ex)
        {
            string Msg = "Exception in CreateLogFile: " + ex.Message;
        }
        return file;
    }

    private static PropertyValueCollection GetADproperty(DirectoryEntry item, string ADpropName)
    {
        PropertyValueCollection prop = null;
        if (item.Properties.Contains(ADpropName))
        {
            try
            {
                prop = item.Properties[ADpropName];
            }
            catch (Exception ex)
            {
                //file.WriteLine("Exception on AD property (" + ADpropName + "). Error: " + ex.Message);
            }
        }
        return prop;
    }

    private static Int32 Get_userAccountControl(DirectoryEntry item)
    {
        Int32 uac = 0;
        if (item.Properties.Contains("userAccountControl"))
        {
            try
            {
                uac = (Int32)item.Properties["userAccountControl"][0];
            }
            catch (Exception ex)
            {
                //file.WriteLine("Exception on AD property (" + ADpropName + "). Error: " + ex.Message);
            }
        }
        return uac;
    }

    private static Int64 GetFileTime(SearchResult sr, string ADpropName)
    {
        Int64 time = 0;
        if (sr.Properties.Contains(ADpropName))
        {
            try
            {
                time = (Int64)sr.Properties[ADpropName][0];
            }
            catch (Exception ex)
            {
                //file.WriteLine("Exception on AD property (" + ADpropName + "). Error: " + ex.Message);
            }
        }
        return time;
    }

    private static string GetSID(DirectoryEntry item, string ADpropName)
    {
        string SID = "";
        if (item.Properties.Contains(ADpropName))
        {
            try
            {
                SecurityIdentifier sid = new SecurityIdentifier(
                    item.Properties[ADpropName][0] as byte[], 0);
                SID = sid.ToString();
            }
            catch (Exception ex)
            {
                //file.WriteLine("Exception on AD property (" + ADpropName + "). Error: " + ex.Message);
            }
        }
        return SID;
    }

    private static string GetGroupCategory(PropertyValueCollection ADgrptype)
    {
        // Source: https://msdn.microsoft.com/en-us/library/ms675935(v=vs.85).aspx
        Int32 grouptype = 0;
        if (ADgrptype != null)
            grouptype = (Int32)ADgrptype.Value;
        string GroupCategory = "Distribution";
        if ((grouptype & 0x80000000) > 0)
            GroupCategory = "Security";
        return GroupCategory;
    }

    private static string GetGroupScope(PropertyValueCollection ADgrptype)
    {
        // Source: https://msdn.microsoft.com/en-us/library/ms675935(v=vs.85).aspx
        Int32 grouptype = 0;
        if (ADgrptype != null)
            grouptype = (Int32)ADgrptype.Value;
        string GroupScope = "";
        if ((grouptype & 0x1) == 0x1)
            GroupScope = "BuiltIn";
        else if ((grouptype & 0x2) == 0x2)
            GroupScope = "Global";
        else if ((grouptype & 0x4) == 0x4)
            GroupScope = "DomainLocal";
        else if ((grouptype & 0x8) == 0x8)
            GroupScope = "Universal";
        return GroupScope;
    }

    public static void SaveGroupMembersToXml(XmlDocument doc, XmlElement body, string ParentDS, 
        PropertyValueCollection GroupMembers)
    {
        try
        {
            XmlElement group = doc.CreateElement(string.Empty, "Group", string.Empty);
            group.SetAttribute("GrpDS", ParentDS);
            body.AppendChild(group);

            foreach (Object obj in GroupMembers)
            {
                string GrpMember = (string)obj;
                XmlElement member = doc.CreateElement(string.Empty, "Member", string.Empty);
                member.SetAttribute("MemberDS", GrpMember);
                group.AppendChild(member);
            }
        }
        catch(Exception ex)
        {

        }
    }
}   // endof: StoredProcedures partial class

public class UACflags
{
    // Source: https://support.microsoft.com/en-us/kb/305144
    // and: http://www.selfadsi.org/ads-attributes/user-userAccountControl.htm
    public Int32 ADobj_flags;
    public Dictionary<string, Int32> flagsLookup;
    public UACflags(Int32 UAC_flags)
    {
        this.ADobj_flags = UAC_flags;
        this.flagsLookup = new Dictionary<string, Int32>();
        this.flagsLookup.Add("SCRIPT", 0x0001);
        this.flagsLookup.Add("ACCOUNTDISABLE", 0x0002);
        this.flagsLookup.Add("HOMEDIR_REQUIRED", 0x0008);
        this.flagsLookup.Add("LOCKOUT", 0x0010);
        this.flagsLookup.Add("PASSWD_NOTREQD", 0x0020);
        this.flagsLookup.Add("PASSWD_CANT_CHANGE", 0x0040);
        this.flagsLookup.Add("ENCRYPTED_TEXT_PWD_ALLOWED", 0x0080);
        this.flagsLookup.Add("TEMP_DUPLICATE_ACCOUNT", 0x0100);
        this.flagsLookup.Add("NORMAL_ACCOUNT", 0x0200);
        this.flagsLookup.Add("INTERDOMAIN_TRUST_ACCOUNT", 0x0800);
        this.flagsLookup.Add("WORKSTATION_TRUST_ACCOUNT", 0x1000);
        this.flagsLookup.Add("SERVER_TRUST_ACCOUNT", 0x2000);
        this.flagsLookup.Add("DONT_EXPIRE_PASSWORD", 0x10000);
        this.flagsLookup.Add("MNS_LOGON_ACCOUNT", 0x20000);
        this.flagsLookup.Add("SMARTCARD_REQUIRED", 0x40000);
        this.flagsLookup.Add("TRUSTED_FOR_DELEGATION", 0x80000);
        this.flagsLookup.Add("NOT_DELEGATED", 0x100000);
        this.flagsLookup.Add("USE_DES_KEY_ONLY", 0x200000);
        this.flagsLookup.Add("DONT_REQ_PREAUTH", 0x400000);
        this.flagsLookup.Add("PASSWORD_EXPIRED", 0x800000);
        this.flagsLookup.Add("TRUSTED_TO_AUTH_FOR_DELEGATION", 0x1000000);
        this.flagsLookup.Add("PARTIAL_SECRETS_ACCOUNT", 0x04000000);
    }

    public Boolean GetFlag(string UAC_flag)
    {
        Boolean ret_flag = false;
        if(flagsLookup.ContainsKey(UAC_flag))
        {
            Int32 mask = flagsLookup[UAC_flag];
            ret_flag = ((ADobj_flags & mask) != 0) ? true : false;
            if (UAC_flag == "ACCOUNTDISABLE")
                ret_flag = !ret_flag;
        }
        return ret_flag;
    }
}

public class TableColDefEx
{
    public string ColName;
    public Type datatype;
    public string ADpropName;
    public string OPtype;

    public TableColDefEx(string ColName, Type datatype, string ADpropName, string OPtype)
    {
        this.ColName = ColName;
        this.datatype = datatype;
        this.ADpropName = ADpropName;
        this.OPtype = OPtype;
    }
}

public class ADcolsTable
{
    public TableColDefEx[] collist;
    public bool IsUser, IsContact, IsComputer, IsGroup;

    public ADcolsTable(string ADfilter)
    {
        IsUser = IsContact = IsComputer = IsGroup = false;
        // Use ADfilter parameter to determine the type of AD objects wanted.
        if (ADfilter.Contains("user"))
            MakeUserColList();
        else if (ADfilter.Contains("contact"))
            MakeContactColList();
        else if (ADfilter.Contains("computer"))
            MakeComputerColList();
        else if (ADfilter.Contains("group"))
            MakeGroupColList();
        else
            collist = new TableColDefEx[0];
    }

    private void MakeUserColList()
    {
        IsUser = true;
        collist = new TableColDefEx[66];  // <-- SET number of elements to number of cells copied below.!

        // COPY CODE from "AD_DW_Users table map to .net V4.xlsx".
        // Copy/Paste all cells from "ColListDef" column.
        collist[0] = new TableColDefEx("GivenName", typeof(String), "givenname", "Adprop");
        collist[1] = new TableColDefEx("Initials", typeof(String), "initials", "Adprop");
        collist[2] = new TableColDefEx("Surname", typeof(String), "sn", "Adprop");
        collist[3] = new TableColDefEx("DisplayName", typeof(String), "displayname", "Adprop");
        collist[4] = new TableColDefEx("Description", typeof(String), "description", "Adprop");
        collist[5] = new TableColDefEx("Office", typeof(String), "physicalDeliveryOfficeName", "Adprop");
        collist[6] = new TableColDefEx("OfficePhone", typeof(String), "telephonenumber", "Adprop");
        collist[7] = new TableColDefEx("EmailAddress", typeof(String), "mail", "Adprop");
        collist[8] = new TableColDefEx("HomePage", typeof(String), "wwwhomepage", "Adprop");
        collist[9] = new TableColDefEx("StreetAddress", typeof(String), "streetaddress", "Adprop");
        collist[10] = new TableColDefEx("POBox", typeof(String), "postofficebox", "Adprop");
        collist[11] = new TableColDefEx("City", typeof(String), "l", "Adprop");
        collist[12] = new TableColDefEx("State", typeof(String), "st", "Adprop");
        collist[13] = new TableColDefEx("PostalCode", typeof(String), "postalcode", "Adprop");
        collist[14] = new TableColDefEx("Country", typeof(String), "co", "Adprop");
        collist[15] = new TableColDefEx("HomePhone", typeof(String), "homephone", "Adprop");
        collist[16] = new TableColDefEx("Pager", typeof(String), "pager", "Adprop");
        collist[17] = new TableColDefEx("MobilePhone", typeof(String), "mobile", "Adprop");
        collist[18] = new TableColDefEx("Fax", typeof(String), "facsimileTelephoneNumber", "Adprop");
        collist[19] = new TableColDefEx("Title", typeof(String), "title", "Adprop");
        collist[20] = new TableColDefEx("Department", typeof(String), "department", "Adprop");
        collist[21] = new TableColDefEx("Company", typeof(String), "company", "Adprop");
        collist[22] = new TableColDefEx("Manager", typeof(String), "manager", "Adprop");
        collist[23] = new TableColDefEx("EmployeeID", typeof(String), "employeeid", "Adprop");
        collist[24] = new TableColDefEx("EmployeeNumber", typeof(String), "employeenumber", "Adprop");
        collist[25] = new TableColDefEx("Division", typeof(String), "division", "Adprop");
        collist[26] = new TableColDefEx("Enabled", typeof(Boolean), "ACCOUNTDISABLE", "UAC");
        collist[27] = new TableColDefEx("LockedOut", typeof(Boolean), "LOCKOUT", "UAC");
        collist[28] = new TableColDefEx("MNSLogonAccount", typeof(Boolean), "MNS_LOGON_ACCOUNT", "UAC");
        collist[29] = new TableColDefEx("CannotChangePassword", typeof(Boolean), "PASSWD_CANT_CHANGE", "UAC");
        collist[30] = new TableColDefEx("PasswordExpired", typeof(Boolean), "PASSWORD_EXPIRED", "UAC");
        collist[31] = new TableColDefEx("PasswordNeverExpires", typeof(Boolean), "DONT_EXPIRE_PASSWD", "UAC");
        collist[32] = new TableColDefEx("PasswordNotRequired", typeof(Boolean), "PASSWD_NOTREQD", "UAC");
        collist[33] = new TableColDefEx("SmartcardLogonRequired", typeof(Boolean), "SMARTCARD_REQUIRED", "UAC");
        collist[34] = new TableColDefEx("DoesNotRequirePreAuth", typeof(Boolean), "DONT_REQ_PREAUTH", "UAC");
        collist[35] = new TableColDefEx("AllowReversiblePasswordEncryption", typeof(Boolean), "ENCRYPTED_TEXT_PWD_ALLOWED", "UAC");
        collist[36] = new TableColDefEx("AccountNotDelegated", typeof(Boolean), "NOT_DELEGATED", "UAC");
        collist[37] = new TableColDefEx("TrustedForDelegation", typeof(Boolean), "TRUSTED_FOR_DELEGATION", "UAC");
        collist[38] = new TableColDefEx("TrustedToAuthForDelegation", typeof(Boolean), "TRUSTED_TO_AUTH_FOR_DELEGATION", "UAC");
        collist[39] = new TableColDefEx("UseDESKeyOnly", typeof(Boolean), "USE_DES_KEY_ONLY", "UAC");
        collist[40] = new TableColDefEx("HomedirRequired", typeof(Boolean), "HOMEDIR_REQUIRED", "UAC");
        collist[41] = new TableColDefEx("LastBadPasswordAttempt", typeof(DateTime), "badpasswordtime", "filetime");
        collist[42] = new TableColDefEx("BadLogonCount", typeof(Int32), "badpwdcount", "Adprop");
        collist[43] = new TableColDefEx("LastLogonDate", typeof(DateTime), "lastlogon", "filetime");
        collist[44] = new TableColDefEx("LogonCount", typeof(Int32), "logoncount", "Adprop");
        collist[45] = new TableColDefEx("PasswordLastSet", typeof(DateTime), "pwdlastset", "filetime");
        collist[46] = new TableColDefEx("AccountLockoutTime", typeof(DateTime), "lockouttime", "filetime");
        collist[47] = new TableColDefEx("AccountExpirationDate", typeof(DateTime), "accountexpires", "filetime");
        collist[48] = new TableColDefEx("LogonWorkstations", typeof(String), "userworkstations", "Adprop");
        collist[49] = new TableColDefEx("HomeDirectory", typeof(String), "homedirectory", "Adprop");
        collist[50] = new TableColDefEx("HomeDrive", typeof(String), "homedrive", "Adprop");
        collist[51] = new TableColDefEx("ProfilePath", typeof(String), "profilepath", "Adprop");
        collist[52] = new TableColDefEx("ScriptPath", typeof(String), "scriptpath", "Adprop");
        collist[53] = new TableColDefEx("userAccountControl", typeof(Int32), "useraccountcontrol", "Adprop");
        collist[54] = new TableColDefEx("PrimaryGroupID", typeof(Int32), "primarygroupid", "Adprop");
        collist[55] = new TableColDefEx("Name", typeof(String), "name", "Adprop");
        collist[56] = new TableColDefEx("CN", typeof(String), "cn", "Adprop");
        collist[57] = new TableColDefEx("UserPrincipalName", typeof(String), "userprincipalname", "Adprop");
        collist[58] = new TableColDefEx("SamAccountName", typeof(String), "samaccountname", "Adprop");
        collist[59] = new TableColDefEx("DistinguishedName", typeof(String), "distinguishedname", "Adprop");
        collist[60] = new TableColDefEx("Created", typeof(DateTime), "whencreated", "Adprop");
        collist[61] = new TableColDefEx("Modified", typeof(DateTime), "whenchanged", "Adprop");
        collist[62] = new TableColDefEx("ObjectCategory", typeof(String), "objectcategory", "Adprop");
        collist[63] = new TableColDefEx("ObjectClass", typeof(String), "SchemaClassName", "ObjClass");
        collist[64] = new TableColDefEx("SID", typeof(String), "objectsid", "SID");
        collist[65] = new TableColDefEx("ObjectGUID", typeof(Guid), "Guid", "ObjGuid");
    }

    private void MakeContactColList()
    {
        IsContact = true;
        collist = new TableColDefEx[34];  // <-- SET number of elements to number of cells copied below.!

        // COPY CODE from "AD_DW_Users table map to .net V4.xlsx".
        // Copy/Paste all cells from "ColListDef" column.
        collist[0] = new TableColDefEx("GivenName", typeof(String), "givenname", "Adprop");
        collist[1] = new TableColDefEx("Initials", typeof(String), "initials", "Adprop");
        collist[2] = new TableColDefEx("Surname", typeof(String), "sn", "Adprop");
        collist[3] = new TableColDefEx("DisplayName", typeof(String), "displayname", "Adprop");
        collist[4] = new TableColDefEx("Description", typeof(String), "description", "Adprop");
        collist[5] = new TableColDefEx("Office", typeof(String), "physicalDeliveryOfficeName", "Adprop");
        collist[6] = new TableColDefEx("OfficePhone", typeof(String), "telephonenumber", "Adprop");
        collist[7] = new TableColDefEx("EmailAddress", typeof(String), "mail", "Adprop");
        collist[8] = new TableColDefEx("HomePage", typeof(String), "wwwhomepage", "Adprop");
        collist[9] = new TableColDefEx("StreetAddress", typeof(String), "streetaddress", "Adprop");
        collist[10] = new TableColDefEx("POBox", typeof(String), "postofficebox", "Adprop");
        collist[11] = new TableColDefEx("City", typeof(String), "l", "Adprop");
        collist[12] = new TableColDefEx("State", typeof(String), "st", "Adprop");
        collist[13] = new TableColDefEx("PostalCode", typeof(String), "postalcode", "Adprop");
        collist[14] = new TableColDefEx("Country", typeof(String), "co", "Adprop");
        collist[15] = new TableColDefEx("HomePhone", typeof(String), "homephone", "Adprop");
        collist[16] = new TableColDefEx("Pager", typeof(String), "pager", "Adprop");
        collist[17] = new TableColDefEx("MobilePhone", typeof(String), "mobile", "Adprop");
        collist[18] = new TableColDefEx("Fax", typeof(String), "facsimileTelephoneNumber", "Adprop");
        collist[19] = new TableColDefEx("Title", typeof(String), "title", "Adprop");
        collist[20] = new TableColDefEx("Department", typeof(String), "department", "Adprop");
        collist[21] = new TableColDefEx("Company", typeof(String), "company", "Adprop");
        collist[22] = new TableColDefEx("Manager", typeof(String), "manager", "Adprop");
        collist[23] = new TableColDefEx("EmployeeID", typeof(String), "employeeid", "Adprop");
        collist[24] = new TableColDefEx("EmployeeNumber", typeof(String), "employeenumber", "Adprop");
        collist[25] = new TableColDefEx("Division", typeof(String), "division", "Adprop");
        collist[26] = new TableColDefEx("DistinguishedName", typeof(String), "distinguishedname", "Adprop");
        collist[27] = new TableColDefEx("Name", typeof(String), "name", "Adprop");
        collist[28] = new TableColDefEx("CN", typeof(String), "cn", "Adprop");
        collist[29] = new TableColDefEx("Created", typeof(DateTime), "whencreated", "Adprop");
        collist[30] = new TableColDefEx("Modified", typeof(DateTime), "whenchanged", "Adprop");
        collist[31] = new TableColDefEx("ObjectCategory", typeof(String), "objectcategory", "Adprop");
        collist[32] = new TableColDefEx("ObjectClass", typeof(String), "SchemaClassName", "ObjClass");
        collist[33] = new TableColDefEx("ObjectGUID", typeof(Guid), "Guid", "ObjGuid");
    }

    private void MakeComputerColList()
    {
        IsComputer = true;
        collist = new TableColDefEx[41];  // <-- SET number of elements to number of cells copied below.!

        // COPY CODE from "AD_DW_Users table map to .net V4.xlsx".
        // Copy/Paste all cells from "ColListDef" column.
        collist[0] = new TableColDefEx("Name", typeof(String), "name", "Adprop");
        collist[1] = new TableColDefEx("DNSHostName", typeof(String), "dnshostname", "Adprop");
        collist[2] = new TableColDefEx("Description", typeof(String), "description", "Adprop");
        collist[3] = new TableColDefEx("Location", typeof(String), "location", "Adprop");
        collist[4] = new TableColDefEx("OperatingSystem", typeof(String), "operatingsystem", "Adprop");
        collist[5] = new TableColDefEx("OperatingSystemVersion", typeof(String), "operatingsystemversion", "Adprop");
        collist[6] = new TableColDefEx("OperatingSystemServicePack", typeof(String), "operatingsystemservicepack", "Adprop");
        collist[7] = new TableColDefEx("ManagedBy", typeof(String), "managedby", "Adprop");
        collist[8] = new TableColDefEx("Enabled", typeof(Boolean), "ACCOUNTDISABLE", "UAC");
        collist[9] = new TableColDefEx("LockedOut", typeof(Boolean), "LOCKOUT", "UAC");
        collist[10] = new TableColDefEx("MNSLogonAccount", typeof(Boolean), "MNS_LOGON_ACCOUNT", "UAC");
        collist[11] = new TableColDefEx("CannotChangePassword", typeof(Boolean), "PASSWD_CANT_CHANGE", "UAC");
        collist[12] = new TableColDefEx("PasswordExpired", typeof(Boolean), "PASSWORD_EXPIRED", "UAC");
        collist[13] = new TableColDefEx("PasswordNeverExpires", typeof(Boolean), "DONT_EXPIRE_PASSWD", "UAC");
        collist[14] = new TableColDefEx("PasswordNotRequired", typeof(Boolean), "PASSWD_NOTREQD", "UAC");
        collist[15] = new TableColDefEx("SmartcardLogonRequired", typeof(Boolean), "SMARTCARD_REQUIRED", "UAC");
        collist[16] = new TableColDefEx("DoesNotRequirePreAuth", typeof(Boolean), "DONT_REQ_PREAUTH", "UAC");
        collist[17] = new TableColDefEx("AllowReversiblePasswordEncryption", typeof(Boolean), "ENCRYPTED_TEXT_PWD_ALLOWED", "UAC");
        collist[18] = new TableColDefEx("AccountNotDelegated", typeof(Boolean), "NOT_DELEGATED", "UAC");
        collist[19] = new TableColDefEx("TrustedForDelegation", typeof(Boolean), "TRUSTED_FOR_DELEGATION", "UAC");
        collist[20] = new TableColDefEx("TrustedToAuthForDelegation", typeof(Boolean), "TRUSTED_TO_AUTH_FOR_DELEGATION", "UAC");
        collist[21] = new TableColDefEx("UseDESKeyOnly", typeof(Boolean), "USE_DES_KEY_ONLY", "UAC");
        collist[22] = new TableColDefEx("LastBadPasswordAttempt", typeof(DateTime), "badpasswordtime", "filetime");
        collist[23] = new TableColDefEx("BadLogonCount", typeof(Int32), "badpwdcount", "Adprop");
        collist[24] = new TableColDefEx("LastLogonDate", typeof(DateTime), "lastlogon", "filetime");
        collist[25] = new TableColDefEx("logonCount", typeof(Int32), "logoncount", "Adprop");
        collist[26] = new TableColDefEx("PasswordLastSet", typeof(DateTime), "pwdlastset", "filetime");
        collist[27] = new TableColDefEx("AccountLockoutTime", typeof(DateTime), "lockouttime", "filetime");
        collist[28] = new TableColDefEx("AccountExpirationDate", typeof(DateTime), "accountexpires", "filetime");
        collist[29] = new TableColDefEx("Created", typeof(DateTime), "whenCreated", "Adprop");
        collist[30] = new TableColDefEx("Modified", typeof(DateTime), "whenChanged", "Adprop");
        collist[31] = new TableColDefEx("CN", typeof(String), "cn", "Adprop");
        collist[32] = new TableColDefEx("DisplayName", typeof(String), "displayname", "Adprop");
        collist[33] = new TableColDefEx("DistinguishedName", typeof(String), "distinguishedname", "Adprop");
        collist[34] = new TableColDefEx("PrimaryGroupID", typeof(Int32), "primarygroupid", "Adprop");
        collist[35] = new TableColDefEx("SamAccountName", typeof(String), "samaccountname", "Adprop");
        collist[36] = new TableColDefEx("userAccountControl", typeof(Int32), "useraccountcontrol", "Adprop");
        collist[37] = new TableColDefEx("ObjectCategory", typeof(String), "objectcategory", "Adprop");
        collist[38] = new TableColDefEx("ObjectClass", typeof(String), "SchemaClassName", "ObjClass");
        collist[39] = new TableColDefEx("SID", typeof(String), "objectsid", "SID");
        collist[40] = new TableColDefEx("ObjectGUID", typeof(Guid), "Guid", "ObjGuid");
    }

    private void MakeGroupColList()
    {
        IsGroup = true;
        collist = new TableColDefEx[15];  // <-- SET number of elements to number of cells copied below.!

        // COPY CODE from "AD_DW_Users table map to .net V4.xlsx".
        // Copy/Paste all cells from "ColListDef" column.
        collist[0] = new TableColDefEx("Name", typeof(String), "name", "Adprop");
        collist[1] = new TableColDefEx("GroupCategory", typeof(String), "grouptype", "GrpCat");
        collist[2] = new TableColDefEx("GroupScope", typeof(String), "grouptype", "GrpScope");
        collist[3] = new TableColDefEx("Description", typeof(String), "description", "Adprop");
        collist[4] = new TableColDefEx("EmailAddress", typeof(String), "mail", "Adprop");
        collist[5] = new TableColDefEx("ManagedBy", typeof(String), "managedby", "Adprop");
        collist[6] = new TableColDefEx("DistinguishedName", typeof(String), "distinguishedname", "Adprop");
        collist[7] = new TableColDefEx("DisplayName", typeof(String), "displayname", "Adprop");
        collist[8] = new TableColDefEx("Created", typeof(DateTime), "whenCreated", "Adprop");
        collist[9] = new TableColDefEx("Modified", typeof(DateTime), "whenChanged", "Adprop");
        collist[10] = new TableColDefEx("SamAccountName", typeof(String), "samaccountname", "Adprop");
        collist[11] = new TableColDefEx("ObjectCategory", typeof(String), "objectcategory", "Adprop");
        collist[12] = new TableColDefEx("ObjectClass", typeof(String), "SchemaClassName", "ObjClass");
        collist[13] = new TableColDefEx("SID", typeof(String), "objectsid", "SID");
        collist[14] = new TableColDefEx("ObjectGUID", typeof(Guid), "Guid", "ObjGuid");
    }

    public DataTable CreateTable()
    {
        DataTable tbl = new DataTable();
        foreach (TableColDefEx col in collist)
        {
            tbl.Columns.Add(col.ColName, col.datatype);
        }
        return tbl;
    }
}

// Source: https://msdn.microsoft.com/en-us/library/ff878201.aspx
public static class DataSetUtilities
{

    public static void SendDataSet(DataSet ds)
    {
        if (ds == null)
        {
            throw new ArgumentException("SendDataSet requires a non-null data set.");
        }
        else
        {
            foreach (DataTable dt in ds.Tables)
            {
                SendDataTable(dt);
            }
        }
    }


    public static void SendDataTable(DataTable dt)
    {
        bool[] coerceToString;  // Do we need to coerce this column to string?
        SqlMetaData[] metaData = ExtractDataTableColumnMetaData(dt, out coerceToString);

        SqlDataRecord record = new SqlDataRecord(metaData);
        SqlPipe pipe = SqlContext.Pipe;
        pipe.SendResultsStart(record);
        try
        {
            foreach (DataRow row in dt.Rows)
            {
                for (int index = 0; index < record.FieldCount; index++)
                {
                    object value = row[index];
                    if (null != value && coerceToString[index])
                        value = value.ToString();
                    record.SetValue(index, value);
                }

                pipe.SendResultsRow(record);
            }
        }
        finally
        {
            pipe.SendResultsEnd();
        }
    }

    private static SqlMetaData[] ExtractDataTableColumnMetaData(DataTable dt, out bool[] coerceToString)
    {
        SqlMetaData[] metaDataResult = new SqlMetaData[dt.Columns.Count];
        coerceToString = new bool[dt.Columns.Count];
        for (int index = 0; index < dt.Columns.Count; index++)
        {
            DataColumn column = dt.Columns[index];
            metaDataResult[index] = SqlMetaDataFromColumn(column, out coerceToString[index]);
        }

        return metaDataResult;
    }

    private static Exception InvalidDataTypeCode(TypeCode code)
    {
        return new ArgumentException("Invalid type: " + code);
    }

    private static Exception UnknownDataType(Type clrType)
    {
        return new ArgumentException("Unknown type: " + clrType);
    }

    private static SqlMetaData SqlMetaDataFromColumn(DataColumn column, out bool coerceToString)
    {
        coerceToString = false;
        SqlMetaData sql_md = null;
        Type clrType = column.DataType;
        string name = column.ColumnName;
        switch (Type.GetTypeCode(clrType))
        {
            case TypeCode.Boolean: sql_md = new SqlMetaData(name, SqlDbType.Bit); break;
            case TypeCode.Byte: sql_md = new SqlMetaData(name, SqlDbType.TinyInt); break;
            case TypeCode.Char: sql_md = new SqlMetaData(name, SqlDbType.NVarChar, 1); break;
            case TypeCode.DateTime: sql_md = new SqlMetaData(name, SqlDbType.DateTime); break;
            case TypeCode.DBNull: throw InvalidDataTypeCode(TypeCode.DBNull);
            case TypeCode.Decimal: sql_md = new SqlMetaData(name, SqlDbType.Decimal, 18, 0); break;
            case TypeCode.Double: sql_md = new SqlMetaData(name, SqlDbType.Float); break;
            case TypeCode.Empty: throw InvalidDataTypeCode(TypeCode.Empty);
            case TypeCode.Int16: sql_md = new SqlMetaData(name, SqlDbType.SmallInt); break;
            case TypeCode.Int32: sql_md = new SqlMetaData(name, SqlDbType.Int); break;
            case TypeCode.Int64: sql_md = new SqlMetaData(name, SqlDbType.BigInt); break;
            case TypeCode.SByte: throw InvalidDataTypeCode(TypeCode.SByte);
            case TypeCode.Single: sql_md = new SqlMetaData(name, SqlDbType.Real); break;
            case TypeCode.String:
                sql_md = new SqlMetaData(name, SqlDbType.NVarChar, column.MaxLength);
                break;
            case TypeCode.UInt16: throw InvalidDataTypeCode(TypeCode.UInt16);
            case TypeCode.UInt32: throw InvalidDataTypeCode(TypeCode.UInt32);
            case TypeCode.UInt64: throw InvalidDataTypeCode(TypeCode.UInt64);
            case TypeCode.Object:
                sql_md = SqlMetaDataFromObjectColumn(name, column, clrType);
                if (sql_md == null)
                {
                    // Unknown type, try to treat it as string;
                    sql_md = new SqlMetaData(name, SqlDbType.NVarChar, column.MaxLength);
                    coerceToString = true;
                }
                break;

            default: throw UnknownDataType(clrType);
        }

        return sql_md;
    }

    private static SqlMetaData SqlMetaDataFromObjectColumn(string name, DataColumn column, Type clrType)
    {
        SqlMetaData sql_md = null;
        if (clrType == typeof(System.Byte[]) || clrType == typeof(SqlBinary) || clrType == typeof(SqlBytes) ||
    clrType == typeof(System.Char[]) || clrType == typeof(SqlString) || clrType == typeof(SqlChars))
            sql_md = new SqlMetaData(name, SqlDbType.VarBinary, column.MaxLength);
        else if (clrType == typeof(System.Guid))
            sql_md = new SqlMetaData(name, SqlDbType.UniqueIdentifier);
        else if (clrType == typeof(System.Object))
            sql_md = new SqlMetaData(name, SqlDbType.Variant);
        else if (clrType == typeof(SqlBoolean))
            sql_md = new SqlMetaData(name, SqlDbType.Bit);
        else if (clrType == typeof(SqlByte))
            sql_md = new SqlMetaData(name, SqlDbType.TinyInt);
        else if (clrType == typeof(SqlDateTime))
            sql_md = new SqlMetaData(name, SqlDbType.DateTime);
        else if (clrType == typeof(SqlDouble))
            sql_md = new SqlMetaData(name, SqlDbType.Float);
        else if (clrType == typeof(SqlGuid))
            sql_md = new SqlMetaData(name, SqlDbType.UniqueIdentifier);
        else if (clrType == typeof(SqlInt16))
            sql_md = new SqlMetaData(name, SqlDbType.SmallInt);
        else if (clrType == typeof(SqlInt32))
            sql_md = new SqlMetaData(name, SqlDbType.Int);
        else if (clrType == typeof(SqlInt64))
            sql_md = new SqlMetaData(name, SqlDbType.BigInt);
        else if (clrType == typeof(SqlMoney))
            sql_md = new SqlMetaData(name, SqlDbType.Money);
        else if (clrType == typeof(SqlDecimal))
            sql_md = new SqlMetaData(name, SqlDbType.Decimal, SqlDecimal.MaxPrecision, 0);
        else if (clrType == typeof(SqlSingle))
            sql_md = new SqlMetaData(name, SqlDbType.Real);
        else if (clrType == typeof(SqlXml))
            sql_md = new SqlMetaData(name, SqlDbType.Xml);
        else
            sql_md = null;

        return sql_md;
    }

}   // endof: class DataSetUtilities
