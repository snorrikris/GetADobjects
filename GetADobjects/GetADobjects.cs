using System;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using Microsoft.SqlServer.Server;
using System.IO;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;

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
    public static void clr_GetADusers ()
    {
        string folder = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);

        // Combine the base folder with your specific folder....
        string specificFolder = Path.Combine(folder, "CLR_Test3");

        // Check if folder exists and if not, create it
        if (!Directory.Exists(specificFolder))
            Directory.CreateDirectory(specificFolder);

        string filename = Path.Combine(specificFolder, "Log.txt");

        System.IO.StreamWriter file =
            new System.IO.StreamWriter(filename);

        try
        {
            UserTable UserTblData = new UserTable();
            DataTable tbl = UserTblData.CreateTable();

            PrincipalContext oPrincipalContext = new PrincipalContext(ContextType.Domain, "veca.is", "DC=veca,DC=is", ContextOptions.Negotiate);

            UserPrincipal up = new UserPrincipal(oPrincipalContext);

            up.SamAccountName = "*";

            PrincipalSearcher ps = new PrincipalSearcher();
            ps.QueryFilter = up;

            PrincipalSearchResult<Principal> results = ps.FindAll();

            if (results != null)
            {
                DataRow row;

                foreach (UserPrincipal user in results)
                {
                    row = tbl.NewRow();

                    // COPY CODE from "AD_DW_Users table map to .net V4.xlsx".
                    // Filter on Method=true, IsUACflag=no
                    // Copy/Paste all filtered cells from "Methods" column.
                    if (user.AccountExpirationDate != null) row[0] = user.AccountExpirationDate;
                    if (user.AccountLockoutTime != null) row[1] = user.AccountLockoutTime;
                    row[3] = user.AllowReversiblePasswordEncryption;
                    row[4] = user.BadLogonCount;
                    row[5] = user.UserCannotChangePassword;
                    row[12] = user.Description;
                    row[13] = user.DisplayName;
                    row[14] = user.DistinguishedName;
                    row[17] = user.EmailAddress;
                    row[18] = user.EmployeeId;
                    row[20] = user.Enabled;
                    row[22] = user.GivenName;
                    row[23] = user.HomeDirectory;
                    row[25] = user.HomeDrive;
                    if (user.LastBadPasswordAttempt != null) row[29] = user.LastBadPasswordAttempt;
                    if (user.LastLogon != null) row[30] = user.LastLogon;
                    row[31] = user.IsAccountLockedOut();
                    PrincipalValueCollection<string> wslist = user.PermittedWorkstations;
                    string wscsv = "";
                    bool IsFirst = true;
                    foreach (string ws in wslist)
                    {
                        if (!IsFirst)
                            wscsv += ", ";
                        wscsv += ws;
                        IsFirst = false;
                    }
                    row[32] = wscsv;
                    row[37] = user.Name;
                    row[39] = user.StructuralObjectClass;
                    row[40] = user.Guid;
                    if (user.LastPasswordSet != null) row[45] = user.LastPasswordSet;
                    row[46] = user.PasswordNeverExpires;
                    row[47] = user.PasswordNotRequired;
                    row[52] = user.SamAccountName;
                    row[53] = user.ScriptPath;
                    row[54] = user.Sid;
                    row[55] = user.SmartcardLogonRequired;
                    row[58] = user.Surname;
                    row[63] = user.UserPrincipalName;


                    bool HasPropList = false;
                    DirectoryEntry directoryEntry = user.GetUnderlyingObject() as DirectoryEntry;
                    for (int i = 0; i < UserTblData.collist.Length; i++)
                    {
                        TableColDef coldef = UserTblData.collist[i];
                        if (coldef.IsMethod)
                            continue;
                        if (!HasPropList)
                        {
                            System.DirectoryServices.PropertyCollection props = directoryEntry.Properties;
                            foreach (string propertyName in props.PropertyNames)
                            {
                                file.WriteLine("Property name: " + propertyName);
                                foreach (object value in directoryEntry.Properties[propertyName])
                                {
                                    file.WriteLine("\t{0} \t({1})", value.ToString(), value.GetType());
                                }
                            }
                            HasPropList = true;
                        }
                        if (directoryEntry.Properties.Contains(coldef.ADpropName))
                        {
                            try
                            {
                                row[i] = directoryEntry.Properties[coldef.ADpropName].Value;
                            }
                            catch (Exception ex)
                            {
                                file.WriteLine("Exception on AD property (" + coldef.ADpropName + "). Error: " + ex.Message);
                            }
                        }
                        else
                        {
                            file.WriteLine("Missing property (" + coldef.ADpropName + ") on user " + user.SamAccountName);
                        }
                    }

                    // Get "userAccountControl" from current row.
                    Int32 uac = (Int32)row["userAccountControl"];

                    //const Int32 SCRIPT = 0x0001;
                    //const Int32 ACCOUNTDISABLE = 0x0002;
                    const Int32 HOMEDIR_REQUIRED = 0x0008;
                    //const Int32 LOCKOUT = 0x0010;
                    //const Int32 PASSWD_NOTREQD = 0x0020;
                    //const Int32 PASSWD_CANT_CHANGE = 0x0040;
                    //const Int32 ENCRYPTED_TEXT_PWD_ALLOWED = 0x0080;
                    //const Int32 TEMP_DUPLICATE_ACCOUNT = 0x0100;
                    //const Int32 NORMAL_ACCOUNT = 0x0200;
                    //const Int32 INTERDOMAIN_TRUST_ACCOUNT = 0x0800;
                    //const Int32 WORKSTATION_TRUST_ACCOUNT = 0x1000;
                    //const Int32 SERVER_TRUST_ACCOUNT = 0x2000;
                    //const Int32 DONT_EXPIRE_PASSWORD = 0x10000;
                    const Int32 MNS_LOGON_ACCOUNT = 0x20000;
                    //const Int32 SMARTCARD_REQUIRED = 0x40000;
                    const Int32 TRUSTED_FOR_DELEGATION = 0x80000;
                    const Int32 NOT_DELEGATED = 0x100000;
                    const Int32 USE_DES_KEY_ONLY = 0x200000;
                    const Int32 DONT_REQ_PREAUTH = 0x400000;
                    const Int32 PASSWORD_EXPIRED = 0x800000;
                    const Int32 TRUSTED_TO_AUTH_FOR_DELEGATION = 0x1000000;
                    //const Int32 PARTIAL_SECRETS_ACCOUNT = 0x04000000;

                    // COPY CODE from "AD_DW_Users table map to .net V4.xlsx".
                    // Filter on IsUACflag=yes
                    // Copy/Paste all filtered cells from "Methods" column.
                    row[2] = ((uac & NOT_DELEGATED) != 0) ? true : false;
                    row[16] = ((uac & DONT_REQ_PREAUTH) != 0) ? true : false;
                    row[24] = ((uac & HOMEDIR_REQUIRED) != 0) ? true : false;
                    row[34] = ((uac & MNS_LOGON_ACCOUNT) != 0) ? true : false;
                    row[44] = ((uac & PASSWORD_EXPIRED) != 0) ? true : false;
                    row[60] = ((uac & TRUSTED_FOR_DELEGATION) != 0) ? true : false;
                    row[61] = ((uac & TRUSTED_TO_AUTH_FOR_DELEGATION) != 0) ? true : false;
                    row[62] = ((uac & USE_DES_KEY_ONLY) != 0) ? true : false;

                    tbl.Rows.Add(row);
                }
            }
            DataSetUtilities.SendDataTable(tbl);
        }
        catch (Exception ex)
        {
            file.WriteLine("Exception: " + ex.Message);
        }

        file.Close();
    }   // endof: clr_GetADusers

    [Microsoft.SqlServer.Server.SqlProcedure]
    public static void clr_GetADcontacts()
    {
        string folder = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);

        // Combine the base folder with your specific folder....
        string specificFolder = Path.Combine(folder, "GetADobjects");

        // Check if folder exists and if not, create it
        if (!Directory.Exists(specificFolder))
            Directory.CreateDirectory(specificFolder);

        string filename = Path.Combine(specificFolder, "Log.txt");

        System.IO.StreamWriter file =
            new System.IO.StreamWriter(filename);

        try
        {
            ContactsTable ContactsTblData = new ContactsTable();
            DataTable tbl = ContactsTblData.CreateTable();

            PrincipalContext oPrincipalContext = new PrincipalContext(ContextType.Domain, "veca.is", "DC=veca,DC=is", ContextOptions.Negotiate);

            ContactPrincipal up = new ContactPrincipal(oPrincipalContext);

            up.Surname = "*";

            PrincipalSearcher ps = new PrincipalSearcher();
            ps.QueryFilter = up;

            PrincipalSearchResult<Principal> results = ps.FindAll();

            if (results != null)
            {
                DataRow row;

                foreach (ContactPrincipal contact in results)
                {
                    row = tbl.NewRow();

                    row[23] = contact.StructuralObjectClass;
                    row[24] = contact.Guid;

                    bool HasPropList = false;
                    DirectoryEntry directoryEntry = contact.GetUnderlyingObject() as DirectoryEntry;
                    for (int i = 0; i < ContactsTblData.collist.Length; i++)
                    {
                        TableColDef coldef = ContactsTblData.collist[i];
                        if (coldef.IsMethod)
                            continue;
                        if (!HasPropList)
                        {
                            System.DirectoryServices.PropertyCollection props = directoryEntry.Properties;
                            foreach (string propertyName in props.PropertyNames)
                            {
                                file.WriteLine("Property name: " + propertyName);
                                foreach (object value in directoryEntry.Properties[propertyName])
                                {
                                    file.WriteLine("\t{0} \t({1})", value.ToString(), value.GetType());
                                }
                            }
                            HasPropList = true;
                        }
                        if (directoryEntry.Properties.Contains(coldef.ADpropName))
                        {
                            try
                            {
                                row[i] = directoryEntry.Properties[coldef.ADpropName].Value;
                            }
                            catch (Exception ex)
                            {
                                file.WriteLine("Exception on AD property (" + coldef.ADpropName + "). Error: " + ex.Message);
                            }
                        }
                        else
                        {
                            file.WriteLine("Missing property (" + coldef.ADpropName + ") on contact " + contact.Name);
                        }
                    }

                    tbl.Rows.Add(row);
                }
            }
            DataSetUtilities.SendDataTable(tbl);
        }
        catch (Exception ex)
        {
            file.WriteLine("Exception: " + ex.Message);
        }

        file.Close();
    }   // endof: clr_GetADcontacts
}

public class TableColDef
{
    public string ColName;
    public Type datatype;
    public string ADpropName;
    public bool IsMethod;

    public TableColDef(string ColName, Type datatype, string ADpropName, bool IsMethod)
    {
        this.ColName = ColName;
        this.datatype = datatype;
        this.ADpropName = ADpropName;
        this.IsMethod = IsMethod;
    }
}

public class UserTable
{
    public TableColDef[] collist;

    public UserTable()
    {
        collist = new TableColDef[65];  // <-- SET number of elements to number of cells copied below.!

        // COPY CODE from "AD_DW_Users table map to .net V4.xlsx".
        // Copy/Paste all cells from "ColListDef" column.
        collist[0] = new TableColDef("AccountExpirationDate", typeof(DateTime), "AccountExpirationDate", true);
        collist[1] = new TableColDef("AccountLockoutTime", typeof(DateTime), "AccountLockoutTime", true);
        collist[2] = new TableColDef("AccountNotDelegated", typeof(Boolean), "NOT_DELEGATED", true);
        collist[3] = new TableColDef("AllowReversiblePasswordEncryption", typeof(Boolean), "AllowReversiblePasswordEncryption", true);
        collist[4] = new TableColDef("BadLogonCount", typeof(Int32), "BadLogonCount", true);
        collist[5] = new TableColDef("CannotChangePassword", typeof(Boolean), "UserCannotChangePassword", true);
        collist[6] = new TableColDef("City", typeof(String), "l", false);
        collist[7] = new TableColDef("CN", typeof(String), "CN", false);
        collist[8] = new TableColDef("Company", typeof(String), "Company", false);
        collist[9] = new TableColDef("Country", typeof(String), "co", false);
        collist[10] = new TableColDef("Created", typeof(DateTime), "whenCreated", false);
        collist[11] = new TableColDef("Department", typeof(String), "Department", false);
        collist[12] = new TableColDef("Description", typeof(String), "Description", true);
        collist[13] = new TableColDef("DisplayName", typeof(String), "DisplayName", true);
        collist[14] = new TableColDef("DistinguishedName", typeof(String), "DistinguishedName", true);
        collist[15] = new TableColDef("Division", typeof(String), "Division", false);
        collist[16] = new TableColDef("DoesNotRequirePreAuth", typeof(Boolean), "DONT_REQ_PREAUTH", true);
        collist[17] = new TableColDef("EmailAddress", typeof(String), "EmailAddress", true);
        collist[18] = new TableColDef("EmployeeID", typeof(String), "EmployeeId", true);
        collist[19] = new TableColDef("EmployeeNumber", typeof(String), "EmployeeNumber", false);
        collist[20] = new TableColDef("Enabled", typeof(Boolean), "Enabled", true);
        collist[21] = new TableColDef("Fax", typeof(String), "facsimileTelephoneNumber", false);
        collist[22] = new TableColDef("GivenName", typeof(String), "GivenName", true);
        collist[23] = new TableColDef("HomeDirectory", typeof(String), "HomeDirectory", true);
        collist[24] = new TableColDef("HomedirRequired", typeof(Boolean), "HOMEDIR_REQUIRED", true);
        collist[25] = new TableColDef("HomeDrive", typeof(String), "HomeDrive", true);
        collist[26] = new TableColDef("HomePage", typeof(String), "wWWHomePage", false);
        collist[27] = new TableColDef("HomePhone", typeof(String), "HomePhone", false);
        collist[28] = new TableColDef("Initials", typeof(String), "Initials", false);
        collist[29] = new TableColDef("LastBadPasswordAttempt", typeof(DateTime), "LastBadPasswordAttempt", true);
        collist[30] = new TableColDef("LastLogonDate", typeof(DateTime), "LastLogon", true);
        collist[31] = new TableColDef("LockedOut", typeof(Boolean), "IsAccountLockedOut", true);
        collist[32] = new TableColDef("LogonWorkstations", typeof(String), "PermittedWorkstations", true);
        collist[33] = new TableColDef("Manager", typeof(String), "Manager", false);
        collist[34] = new TableColDef("MNSLogonAccount", typeof(Boolean), "MNS_LOGON_ACCOUNT", true);
        collist[35] = new TableColDef("MobilePhone", typeof(String), "mobile", false);
        collist[36] = new TableColDef("Modified", typeof(DateTime), "whenChanged", false);
        collist[37] = new TableColDef("Name", typeof(String), "Name", true);
        collist[38] = new TableColDef("ObjectCategory", typeof(String), "ObjectCategory", false);
        collist[39] = new TableColDef("ObjectClass", typeof(String), "StructuralObjectClass", true);
        collist[40] = new TableColDef("ObjectGUID", typeof(Guid), "Guid", true);
        collist[41] = new TableColDef("Office", typeof(String), "physicalDeliveryOfficeName", false);
        collist[42] = new TableColDef("OfficePhone", typeof(String), "telephoneNumber", false);
        collist[43] = new TableColDef("Pager", typeof(String), "Pager", false);
        collist[44] = new TableColDef("PasswordExpired", typeof(Boolean), "PASSWORD_EXPIRED", true);
        collist[45] = new TableColDef("PasswordLastSet", typeof(DateTime), "LastPasswordSet", true);
        collist[46] = new TableColDef("PasswordNeverExpires", typeof(Boolean), "PasswordNeverExpires", true);
        collist[47] = new TableColDef("PasswordNotRequired", typeof(Boolean), "PasswordNotRequired", true);
        collist[48] = new TableColDef("POBox", typeof(String), "postOfficeBox", false);
        collist[49] = new TableColDef("PostalCode", typeof(String), "PostalCode", false);
        collist[50] = new TableColDef("PrimaryGroupID", typeof(Int32), "PrimaryGroupID", false);
        collist[51] = new TableColDef("ProfilePath", typeof(String), "ProfilePath", false);
        collist[52] = new TableColDef("SamAccountName", typeof(String), "SamAccountName", true);
        collist[53] = new TableColDef("ScriptPath", typeof(String), "ScriptPath", true);
        collist[54] = new TableColDef("SID", typeof(String), "Sid", true);
        collist[55] = new TableColDef("SmartcardLogonRequired", typeof(Boolean), "SmartcardLogonRequired", true);
        collist[56] = new TableColDef("State", typeof(String), "st", false);
        collist[57] = new TableColDef("StreetAddress", typeof(String), "StreetAddress", false);
        collist[58] = new TableColDef("Surname", typeof(String), "Surname", true);
        collist[59] = new TableColDef("Title", typeof(String), "Title", false);
        collist[60] = new TableColDef("TrustedForDelegation", typeof(Boolean), "TRUSTED_FOR_DELEGATION", true);
        collist[61] = new TableColDef("TrustedToAuthForDelegation", typeof(Boolean), "TRUSTED_TO_AUTH_FOR_DELEGATION", true);
        collist[62] = new TableColDef("UseDESKeyOnly", typeof(Boolean), "USE_DES_KEY_ONLY", true);
        collist[63] = new TableColDef("UserPrincipalName", typeof(String), "UserPrincipalName", true);
        collist[64] = new TableColDef("userAccountControl", typeof(Int32), "userAccountControl", false);
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

public class ContactsTable
{
    public TableColDef[] collist;

    public ContactsTable()
    {
        collist = new TableColDef[34];  // <-- SET number of elements to number of cells copied below.!

        // COPY CODE from "AD_DW_Users table map to .net V4.xlsx".
        // Copy/Paste all cells from "ColListDef" column in "Contacts" sheet.
        collist[0] = new TableColDef("City", typeof(String), "l", false);
        collist[1] = new TableColDef("CN", typeof(String), "CN", false);
        collist[2] = new TableColDef("Company", typeof(String), "Company", false);
        collist[3] = new TableColDef("Country", typeof(String), "co", false);
        collist[4] = new TableColDef("Created", typeof(DateTime), "whenCreated", false);
        collist[5] = new TableColDef("Department", typeof(String), "Department", false);
        collist[6] = new TableColDef("Description", typeof(String), "description", false);
        collist[7] = new TableColDef("DisplayName", typeof(String), "displayName", false);
        collist[8] = new TableColDef("DistinguishedName", typeof(String), "distinguishedName", false);
        collist[9] = new TableColDef("Division", typeof(String), "Division", false);
        collist[10] = new TableColDef("EmailAddress", typeof(String), "mail", false);
        collist[11] = new TableColDef("EmployeeID", typeof(String), "employeeID", false);
        collist[12] = new TableColDef("EmployeeNumber", typeof(String), "EmployeeNumber", false);
        collist[13] = new TableColDef("Fax", typeof(String), "facsimileTelephoneNumber", false);
        collist[14] = new TableColDef("GivenName", typeof(String), "givenName", false);
        collist[15] = new TableColDef("HomePage", typeof(String), "wWWHomePage", false);
        collist[16] = new TableColDef("HomePhone", typeof(String), "HomePhone", false);
        collist[17] = new TableColDef("Initials", typeof(String), "Initials", false);
        collist[18] = new TableColDef("Manager", typeof(String), "Manager", false);
        collist[19] = new TableColDef("MobilePhone", typeof(String), "mobile", false);
        collist[20] = new TableColDef("Modified", typeof(DateTime), "whenChanged", false);
        collist[21] = new TableColDef("Name", typeof(String), "name", false);
        collist[22] = new TableColDef("ObjectCategory", typeof(String), "ObjectCategory", false);
        collist[23] = new TableColDef("ObjectClass", typeof(String), "StructuralObjectClass", true);
        collist[24] = new TableColDef("ObjectGUID", typeof(Guid), "Guid", true);
        collist[25] = new TableColDef("Office", typeof(String), "physicalDeliveryOfficeName", false);
        collist[26] = new TableColDef("OfficePhone", typeof(String), "telephoneNumber", false);
        collist[27] = new TableColDef("Pager", typeof(String), "Pager", false);
        collist[28] = new TableColDef("POBox", typeof(String), "postOfficeBox", false);
        collist[29] = new TableColDef("PostalCode", typeof(String), "PostalCode", false);
        collist[30] = new TableColDef("State", typeof(String), "st", false);
        collist[31] = new TableColDef("StreetAddress", typeof(String), "StreetAddress", false);
        collist[32] = new TableColDef("Surname", typeof(String), "sn", false);
        collist[33] = new TableColDef("Title", typeof(String), "Title", false);
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
}   // endof: ContactsTable class

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

// Source: http://stackoverflow.com/questions/14158995/adding-contact-to-distribution-list-the-principal-object-must-have-a-valid-sid
[DirectoryObjectClass("contact")]
[DirectoryRdnPrefix("CN")]
public class ContactPrincipal : AuthenticablePrincipal
{
    public ContactPrincipal(PrincipalContext context)
        : base(context)
    {
    }

    //public static ContactPrincipal FindByIdentity(PrincipalContext context, string identityValue)
    //{
    //    return (ContactPrincipal)Principal.FindByIdentityWithType(context, typeof(ContactPrincipal), identityValue);
    //}

    //public static ContactPrincipal FindByIdentity(PrincipalContext context, IdentityType identityType,
    //                                              string identityValue)
    //{
    //    return
    //       (ContactPrincipal)
    //       Principal.FindByIdentityWithType(context, typeof(ContactPrincipal), identityType, identityValue);
    //}

    //[DirectoryProperty("mail")]
    //public string EmailAddress
    //{
    //    get
    //    {
    //        if (ExtensionGet("mail").Length == 1)
    //        {
    //            return ExtensionGet("mail")[0].ToString();
    //        }
    //        else
    //        {
    //            return null;
    //        }
    //    }
    //    set { ExtensionSet("mail", value); }
    //}

    //[DirectoryProperty("givenName")]
    //public string GivenName
    //{
    //    get
    //    {
    //        if (ExtensionGet("givenName").Length == 1)
    //        {
    //            return ExtensionGet("givenName")[0].ToString();
    //        }
    //        else
    //        {
    //            return null;
    //        }
    //    }
    //    set { ExtensionSet("givenName", value); }
    //}

    //[DirectoryProperty("middleName")]
    //public string MiddleName
    //{
    //    get
    //    {
    //        if (ExtensionGet("middleName").Length == 1)
    //        {
    //            return ExtensionGet("middleName")[0].ToString();
    //        }
    //        else
    //        {
    //            return null;
    //        }
    //    }
    //    set { ExtensionSet("middleName", value); }
    //}

    [DirectoryProperty("sn")]
    public string Surname
    {
        get
        {
            if (ExtensionGet("sn").Length == 1)
            {
                return ExtensionGet("sn")[0].ToString();
            }
            else
            {
                return null;
            }
        }
        set { ExtensionSet("sn", value); }
    }

    //[DirectoryProperty("mobile")]
    //public string MobileTelephoneNumber
    //{
    //    get
    //    {
    //        if (ExtensionGet("mobile").Length == 1)
    //        {
    //            return ExtensionGet("mobile")[0].ToString();
    //        }
    //        else
    //        {
    //            return null;
    //        }
    //    }
    //    set { ExtensionSet("mobile", value); }
    //}

    //[DirectoryProperty("telephoneNumber")]
    //public string VoiceTelephoneNumber
    //{
    //    get
    //    {
    //        if (ExtensionGet("telephoneNumber").Length == 1)
    //        {
    //            return ExtensionGet("telephoneNumber")[0].ToString();
    //        }
    //        else
    //        {
    //            return null;
    //        }
    //    }
    //    set { ExtensionSet("telephoneNumber", value); }
    //}
}   // endof: ContactPrincipal class

