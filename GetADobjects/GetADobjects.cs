using Microsoft.SqlServer.Server;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlTypes;
using System.DirectoryServices;
using System.IO;
using System.Security.Principal;
using System.Xml;

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
        Int32 itemcount = 0;
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
                itemcount++;
                DirectoryEntry item = searchResult.GetDirectoryEntry();
                row = tbl.NewRow();

                UACflags Item_UAC_flags = null;
                Int64 UserPasswordExpiryTimeComputed = 0;
                PropertyValueCollection ADGroupType = null;

                for (int i = 0; i < TblData.collist.Length; i++)
                {
                    TableColDef coldef = TblData.collist[i];
                    switch(coldef.OPtype)
                    {
                        case "Adprop":
                            if (coldef.ADpropName == "useraccountcontrol" && Item_UAC_flags != null)
                            {
                                row[i] = Item_UAC_flags.ADobj_flags;
                                break;
                            }
                            PropertyValueCollection prop = GetADproperty(item, coldef.ADpropName);
                            if (prop != null)
                                row[i] = prop.Value;
                            break;

                        case "UAC":
                            if (Item_UAC_flags == null)
                            {   // Get UAC flags only once per AD object.
                                Item_UAC_flags = new UACflags(Get_userAccountControl(item, out UserPasswordExpiryTimeComputed));
                            }
                            row[i] = Item_UAC_flags.GetFlag(coldef.ADpropName);
                            break;

                        case "ObjClass":
                            row[i] = item.SchemaClassName;
                            break;

                        case "ObjGuid":
                            row[i] = item.Guid;
                            break;

                        case "filetime":
                            Int64 time = 0;
                            if (coldef.ADpropName == "msDS-UserPasswordExpiryTimeComputed")
                                time = UserPasswordExpiryTimeComputed;
                            else
                                time = GetFileTime(searchResult, coldef.ADpropName);
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
            SqlContext.Pipe.Send("COMException in clr_GetADobjects. ItemCounter = " + itemcount.ToString());
            throw;
        }
        catch (InvalidOperationException)
        {
            SqlContext.Pipe.Send("InvalidOperationException in clr_GetADobjects. ItemCounter = " + itemcount.ToString());
            throw;
        }
        catch (NotSupportedException)
        {
            SqlContext.Pipe.Send("NotSupportedException in clr_GetADobjects. ItemCounter = " + itemcount.ToString());
            throw;
        }
        catch (Exception)
        {
            SqlContext.Pipe.Send("Exception in clr_GetADobjects. ItemCounter = " + itemcount.ToString());
            throw;
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

    [Microsoft.SqlServer.Server.SqlProcedure]
    public static void clr_GetADusersPhotos(SqlString ADpath, SqlString ADfilter)
    {
        System.IO.StreamWriter file = CreateLogFile();

        SearchResultCollection results = null;
        Int32 itemcount = 0;
        try
        {
            DataTable tbl = new DataTable();
            tbl.Columns.Add("ObjectGUID", typeof(Guid));
            tbl.Columns.Add("Width", typeof(int));
            tbl.Columns.Add("Height", typeof(int));
            tbl.Columns.Add("Photo", typeof(byte[]));
            DataRow row;

            DirectoryEntry entry = new DirectoryEntry((string)ADpath);
            DirectorySearcher searcher = new DirectorySearcher(entry);
            searcher.Filter = (string)ADfilter;
            searcher.PageSize = 500;

            results = searcher.FindAll();
            foreach (SearchResult searchResult in results)
            {
                itemcount++;
                DirectoryEntry item = searchResult.GetDirectoryEntry();

                PropertyValueCollection prop = GetADproperty(item, "thumbnailphoto");
                if (prop == null)
                    continue;

                // Get image size
                ImgSize imgsize = new ImgSize(0, 0);
                try
                {
                    imgsize = ImageHeader.GetDimensions((byte[])prop[0]);
                }
                catch(Exception ex)
                {
                    SqlContext.Pipe.Send("Warning: Get image size failed for user (" + GetDistinguishedName(item) + ")"
                        + " Exception: " + ex.Message);
                }

                row = tbl.NewRow();
                row[0] = item.Guid;
                if (!imgsize.IsEmpty()) // Image size will be NULL unless size has been read from the image header.
                {
                    row[1] = imgsize.Width;
                    row[2] = imgsize.Height;
                }
                row[3] = prop[0];
                tbl.Rows.Add(row);
            }
            DataSetUtilities.SendDataTable(tbl);
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            SqlContext.Pipe.Send("COMException in clr_GetADusersPhotos. ItemCounter = " + itemcount.ToString());
            throw;
        }
        catch (InvalidOperationException)
        {
            SqlContext.Pipe.Send("InvalidOperationException in clr_GetADusersPhotos. ItemCounter = " + itemcount.ToString());
            throw;
        }
        catch (NotSupportedException)
        {
            SqlContext.Pipe.Send("NotSupportedException in clr_GetADusersPhotos. ItemCounter = " + itemcount.ToString());
            throw;
        }
        catch (Exception)
        {
            SqlContext.Pipe.Send("Exception in clr_GetADusersPhotos. ItemCounter = " + itemcount.ToString());
            throw;
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
    }   // endof: clr_GetADusersPhotos

    [Microsoft.SqlServer.Server.SqlProcedure]
    public static void clr_GetADpropertiesToFile(SqlString ADpath, SqlString ADfilter)
    {
        System.IO.StreamWriter file = CreateLogFile();

        SearchResultCollection results = null;

        try
        {
            DirectoryEntry entry = new DirectoryEntry((string)ADpath);

            DirectorySearcher mySearcher = new DirectorySearcher(entry);
            mySearcher.PropertiesToLoad.Add("msDS-User-Account-Control-Computed");

            
            mySearcher.PropertiesToLoad.Add("ms-DS-UserAccountAutoLocked"); // *
            //mySearcher.PropertiesToLoad.Add("msDS-User-Account-Control-Computed");  // *
            mySearcher.PropertiesToLoad.Add("msDS-UserAccountDisabled");
            mySearcher.PropertiesToLoad.Add("msDS-UserDontExpirePassword");
            mySearcher.PropertiesToLoad.Add("ms-DS-UserEncryptedTextPasswordAllowed");
            mySearcher.PropertiesToLoad.Add("msDS-UserPasswordExpired");    // *
            mySearcher.PropertiesToLoad.Add("ms-DS-UserPasswordNotRequired");
            mySearcher.PropertiesToLoad.Add("userAccountControl");
            
            mySearcher.Filter = (string)ADfilter;

            mySearcher.PageSize = 500;

            results = mySearcher.FindAll();

            foreach (SearchResult searchResult in results)
            {
                DirectoryEntry item = searchResult.GetDirectoryEntry();
                Int64 exptm;
                Int32 uac = Get_userAccountControl(item, out exptm);
                file.WriteLine("uac: 0x" + uac.ToString("X"));

                // Iterate through each property name in each SearchResult.
                foreach (string propertyKey in searchResult.Properties.PropertyNames)
                {
                    // Retrieve the value assigned to that property name 
                    // in the ResultPropertyValueCollection.
                    ResultPropertyValueCollection valueCollection =
                        searchResult.Properties[propertyKey];

                    if (propertyKey == "objectsid")
                    {
                        file.WriteLine("{0}: {1}", propertyKey, GetSID(item, "objectsid"));
                    }
                    else if (propertyKey == "objectguid")
                    {
                        file.WriteLine("{0}: {1}", propertyKey, item.Guid.ToString());
                    }
                    else
                    {
                        // Iterate through values for each property name in each 
                        // SearchResult.
                        foreach (Object propertyValue in valueCollection)
                        {
                            // Handle results. Be aware that the following 
                            // WriteLine only returns readable results for 
                            // properties that are strings.
                            file.WriteLine("{0}: {1}", propertyKey, propertyValue.ToString());
                        }
                    }
                }
                file.WriteLine("---------------------------------------------------------");
            }
        }
        catch
        {
            throw;
        }
        finally
        {
            // To prevent memory leaks, always call 
            // SearchResultCollection.Dispose() manually.
            if (null != results)
            {
                results.Dispose();
                results = null;
            }
        }
        file.Close();
    }   // endof: clr_GetADpropertiesToFile

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
                SqlContext.Pipe.Send("Warning: GetADproperty (" + ADpropName + ") failed for user (" + GetDistinguishedName(item) + ")"
                        + " Exception: " + ex.Message);
            }
        }
        return prop;
    }

    /// <summary>
    /// Get User Account Control flags.
    /// </summary>
    /// <param name="item"></param>
    /// <returns></returns>
    /// <remarks>
    /// References:
    /// https://msdn.microsoft.com/en-us/library/cc223145.aspx
    /// https://msdn.microsoft.com/en-us/library/cc223393.aspx
    /// https://msdn.microsoft.com/en-us/library/ms677840(v=vs.85).aspx
    /// https://technet.microsoft.com/en-us/library/ee198831.aspx
    /// http://stackoverflow.com/questions/25213146/constructed-attributes-in-active-directory-global-catalog-get-password-expiry-f
    /// </remarks>
    private static Int32 Get_userAccountControl(DirectoryEntry item, out Int64 PwdExpComputed)
    {
        Int32 uac = 0;
        PwdExpComputed = 0;
        SearchResult res = null;
        try
        {
            // Need to query AD for every user to get up to date msDS-User-Account-Control-Computed.
            DirectorySearcher srch = new DirectorySearcher(item, "(objectClass=*)",
                new string[] { "userAccountControl", "msDS-User-Account-Control-Computed", "msDS-UserPasswordExpiryTimeComputed" },
                SearchScope.Base);

            if ((res = srch.FindOne()) == null)
                return uac;

            Int32 AC1 = 0, AC2 = 0;
            if(res.Properties.Contains("userAccountControl"))
                AC1 = Convert.ToInt32(res.Properties["userAccountControl"][0]);
            if(res.Properties.Contains("msDS-User-Account-Control-Computed"))
                AC2 = Convert.ToInt32(res.Properties["msDS-User-Account-Control-Computed"][0]);
            uac = AC1 | AC2;
            if (IsUserCannotChangePassword(item))
                uac |= 0x40;

            PwdExpComputed = GetFileTime(res, "msDS-UserPasswordExpiryTimeComputed");
        }
        catch (Exception ex)
        {
            SqlContext.Pipe.Send("Warning: Get_userAccountControl failed for user (" + GetDistinguishedName(item) + ")"
                    + " Exception: " + ex.Message);
        }
        return uac;
    }

    /// <summary>
    /// Get "User cannot change password" property from AD user object.
    /// </summary>
    /// <param name="user"></param>
    /// <returns>true if user cannot change password.</returns>
    /// <remarks>
    /// This function gets the security descripto of the user from AD in SDDL form.
    /// and parses the ACE strings for "Change password".
    /// References:
    /// http://stackoverflow.com/questions/7724110/convert-sddl-to-readable-text-in-net
    /// https://msdn.microsoft.com/en-us/library/windows/desktop/aa379567(v=vs.85).aspx
    /// https://msdn.microsoft.com/en-us/library/windows/desktop/aa379570(v=vs.85).aspx
    /// https://msdn.microsoft.com/en-us/library/windows/desktop/aa374928(v=vs.85).aspx
    /// https://msdn.microsoft.com/en-us/library/windows/desktop/aa379637(v=vs.85).aspx
    /// http://www.netid.washington.edu/documentation/domains/sddl.aspx
    /// http://blogs.technet.com/b/askds/archive/2008/04/18/the-security-descriptor-definition-language-of-love-part-1.aspx
    /// https://technet.microsoft.com/en-us/library/ee198831.aspx
    /// https://msdn.microsoft.com/en-us/library/aa746398.aspx
    /// More (on the problem):
    /// http://stackoverflow.com/questions/29812872/find-users-who-cannot-change-their-password
    /// http://devblog.rayonnant.net/2011/04/ad-net-toggle-users-cant-change.html
    /// https://msdn.microsoft.com/en-us/library/system.directoryservices.accountmanagement.authenticableprincipal.usercannotchangepassword(v=vs.100).aspx
    /// </remarks>
    public static bool IsUserCannotChangePassword(DirectoryEntry user)
    {
        bool cantChange = false;
        try
        {
            ActiveDirectorySecurity userSecurity = user.ObjectSecurity;
            string userSDDL = userSecurity.GetSecurityDescriptorSddlForm(System.Security.AccessControl.AccessControlSections.Access);
            int ix_chgpwdGUID = -1, startIdx = 0;

            // Find ACE strings for "Change password" permission on "this" user.
            while ((ix_chgpwdGUID = userSDDL.IndexOf(
                "AB721A53-1E2F-11D0-9819-00AA0040529B", // <-- GUID of "change password" permission.
                startIdx, StringComparison.CurrentCultureIgnoreCase)) > 0)
            {
                // ACE string format (ace_type;ace_flags;rights;object_guid;inherit_object_guid;account_sid;(resource_attribute))
                // e.g. (OA;;CR;ab721a53-1e2f-11d0-9819-00aa0040529b;;WD)
                int ix_ace_type = -1, ix_ace_flags = -1, ix_rights = -1, ix_object_guid = -1,
                    ix_inherit_object_guid = -1, ix_account_sid = -1, ixStart = -1;
                string ace_type = "", ace_flags = "", rights = "", object_guid = "", inherit_object_guid = "", account_sid = "";

                // Find starting '(' of the ACE string.
                if (ix_chgpwdGUID > 0)
                    ixStart = userSDDL.LastIndexOf('(', ix_chgpwdGUID);  // Find previous '('
                // Find all expected ';' in the ACE string.
                if (ixStart > 0)
                    ix_ace_type = userSDDL.IndexOf(';', ixStart + 1);
                if (ix_ace_type > 0)
                    ix_ace_flags = userSDDL.IndexOf(';', ix_ace_type + 1);
                if (ix_ace_flags > 0)
                    ix_rights = userSDDL.IndexOf(';', ix_ace_flags + 1);
                if (ix_rights > 0)
                    ix_object_guid = userSDDL.IndexOf(';', ix_rights + 1);
                if (ix_object_guid > 0)
                    ix_inherit_object_guid = userSDDL.IndexOf(';', ix_object_guid + 1);
                // Find the closing ')'.
                if (ix_inherit_object_guid > 0)
                    ix_account_sid = userSDDL.IndexOf(')', ix_inherit_object_guid + 1);
                if (ix_account_sid > 0)  // All expected tokens found?
                {
                    // Get string values from ACE string.
                    ace_type = userSDDL.Substring(ixStart + 1, (ix_ace_type - ixStart) - 1);
                    ace_flags = userSDDL.Substring(ix_ace_type + 1, (ix_ace_flags - ix_ace_type) - 1);
                    rights = userSDDL.Substring(ix_ace_flags + 1, (ix_rights - ix_ace_flags) - 1);
                    object_guid = userSDDL.Substring(ix_rights + 1, (ix_object_guid - ix_rights) - 1);
                    inherit_object_guid = userSDDL.Substring(ix_object_guid + 1, (ix_inherit_object_guid - ix_object_guid) - 1);
                    account_sid = userSDDL.Substring(ix_inherit_object_guid + 1, (ix_account_sid - ix_inherit_object_guid) - 1);
                }
                // If Change password permission denied (OD) for Everyone (WD)
                //   OR Change password denied (OD) for 'SELF' (PS)
                // Then "User cannot change password" is set on "this" user.
                if ((ace_type == "OD" && account_sid == "WD")
                    || (ace_type == "OD" && account_sid == "PS"))
                    cantChange = true;

                startIdx = ix_account_sid + 1;   // Continue searching after current ACE string.
            }
        }
        catch (Exception ex)
        {
        }
        return cantChange;
    }

    //private static Int64 GetFileTime(DirectoryEntry item, string ADpropName)
    //{
    //    Int64 time = 0;
    //    if (item.Properties.Contains(ADpropName))
    //    {
    //        try
    //        {
    //            time = (Int64)item.Properties[ADpropName][0];
    //        }
    //        catch (Exception ex)
    //        {
    //            SqlContext.Pipe.Send("Warning: GetFileTime (" + ADpropName + ") failed for object (" + GetDistinguishedName(item) + ")"
    //                    + " Exception: " + ex.Message);
    //        }
    //    }
    //    return time;
    //}

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
                SqlContext.Pipe.Send("Warning: GetFileTime (" + ADpropName + ") failed for object (" + GetDistinguishedName(sr) + ")"
                        + " Exception: " + ex.Message);
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
                SqlContext.Pipe.Send("Warning: GetSID (" + ADpropName + ") failed for object (" + GetDistinguishedName(item) + ")"
                        + " Exception: " + ex.Message);
            }
        }
        return SID;
    }

    // Used when reporting error
    private static string GetDistinguishedName(DirectoryEntry item)
    {
        string ds = "";
        if (item.Properties.Contains("distinguishedname"))
        {
            try
            {
                ds = (string)item.Properties["distinguishedname"][0];
            }
            catch
            {   // ignore exception.
                ds = "[GetDistinguishedName failed]";
            }
        }
        return ds;
    }

    // Used when reporting error
    private static string GetDistinguishedName(SearchResult sr)
    {
        string ds = "";
        if (sr.Properties.Contains("distinguishedname"))
        {
            try
            {
                ds = (string)sr.Properties["distinguishedname"][0];
            }
            catch
            {   // ignore exception.
                ds = "[GetDistinguishedName failed]";
            }
        }
        return ds;
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
            SqlContext.Pipe.Send("Warning: SaveGroupMembersToXml (" + ParentDS + ") failed. Exception: " + ex.Message);
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
        this.flagsLookup.Add("DONT_EXPIRE_PASSWD", 0x10000);
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
        if (flagsLookup.ContainsKey(UAC_flag))
        {
            Int32 mask = flagsLookup[UAC_flag];
            ret_flag = ((ADobj_flags & mask) != 0) ? true : false;
            if (UAC_flag == "ACCOUNTDISABLE")
                ret_flag = !ret_flag;
        }
        else
        {
            SqlContext.Pipe.Send("Warning: UAC flag not found in list of flags (" + UAC_flag + ").");
        }
        return ret_flag;
    }
}

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
            new TableColDef("ObjectGUID", typeof(Guid), "Guid","ObjGuid")
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

public struct ImgSize
{
    public Int32 Width;
    public Int32 Height;
    public ImgSize(Int32 width, Int32 height)
    {
        this.Width = width;
        this.Height = height;
    }
    public bool IsEmpty()
    {
        if (Width == 0 && Height == 0)
            return true;
        return false;
    }
};

// Source: http://www.codeproject.com/Articles/35978/Reading-Image-Headers-to-Get-Width-and-Height
/// <summary>
/// Taken from http://stackoverflow.com/questions/111345/getting-image-dimensions-without-reading-the-entire-file/111349
/// Minor improvements including supporting unsigned 16-bit integers when decoding Jfif and added logic
/// to load the image using new Bitmap if reading the headers fails
/// </summary>
public static class ImageHeader
{
    private static Dictionary<byte[], Func<BinaryReader, ImgSize>> imageFormatDecoders = new Dictionary<byte[], Func<BinaryReader, ImgSize>>()
    { 
        { new byte[] { 0xff, 0xd8 }, DecodeJfif }, 
        { new byte[] { 0x42, 0x4D }, DecodeBitmap }, 
        { new byte[] { 0x47, 0x49, 0x46, 0x38, 0x37, 0x61 }, DecodeGif }, 
        { new byte[] { 0x47, 0x49, 0x46, 0x38, 0x39, 0x61 }, DecodeGif }, 
        { new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }, DecodePng },
    };

    public static ImgSize GetDimensions(byte[] imgdata)
    {
        MemoryStream memstream = null;
        ImgSize imgsize = new ImgSize(0, 0);
        try
        {
            memstream = new MemoryStream(imgdata);

            using (BinaryReader binaryReader = new BinaryReader(memstream))
            {
                imgsize = GetDimensions(binaryReader);
            }
        }
        finally
        {
            if (memstream != null)
                memstream.Close();
        }
        return imgsize;
    }

    /// <summary>        
    /// Gets the dimensions of an image.        
    /// </summary>        
    /// <param name="path">The path of the image to get the dimensions of.</param>        
    /// <returns>The dimensions of the specified image.</returns>        
    /// <exception cref="ArgumentException">The image was of an unrecognised format.</exception>            
    public static ImgSize GetDimensions(BinaryReader binaryReader)
    {
        int maxMagicBytesLength = 8; // imageFormatDecoders.Keys.OrderByDescending(x => x.Length).First().Length;
        byte[] magicBytes = new byte[maxMagicBytesLength];
        for (int i = 0; i < maxMagicBytesLength; i += 1)
        {
            magicBytes[i] = binaryReader.ReadByte();
            foreach (var kvPair in imageFormatDecoders)
            {
                if (StartsWith(magicBytes, kvPair.Key))
                {
                    return kvPair.Value(binaryReader);
                }
            }
        }
        return new ImgSize(0, 0);
    }

    private static bool StartsWith(byte[] thisBytes, byte[] thatBytes)
    {
        for (int i = 0; i < thatBytes.Length; i += 1)
        {
            if (thisBytes[i] != thatBytes[i])
            {
                return false;
            }
        }
        return true;
    }

    private static short ReadLittleEndianInt16(BinaryReader binaryReader)
    {
        byte[] bytes = new byte[sizeof(short)];

        for (int i = 0; i < sizeof(short); i += 1)
        {
            bytes[sizeof(short) - 1 - i] = binaryReader.ReadByte();
        }
        return BitConverter.ToInt16(bytes, 0);
    }

    private static ushort ReadLittleEndianUInt16(BinaryReader binaryReader)
    {
        byte[] bytes = new byte[sizeof(ushort)];

        for (int i = 0; i < sizeof(ushort); i += 1)
        {
            bytes[sizeof(ushort) - 1 - i] = binaryReader.ReadByte();
        }
        return BitConverter.ToUInt16(bytes, 0);
    }

    private static int ReadLittleEndianInt32(BinaryReader binaryReader)
    {
        byte[] bytes = new byte[sizeof(int)];
        for (int i = 0; i < sizeof(int); i += 1)
        {
            bytes[sizeof(int) - 1 - i] = binaryReader.ReadByte();
        }
        return BitConverter.ToInt32(bytes, 0);
    }

    private static ImgSize DecodeBitmap(BinaryReader binaryReader)
    {
        binaryReader.ReadBytes(16);
        int width = binaryReader.ReadInt32();
        int height = binaryReader.ReadInt32();
        return new ImgSize(width, height);
    }

    private static ImgSize DecodeGif(BinaryReader binaryReader)
    {
        int width = binaryReader.ReadInt16();
        int height = binaryReader.ReadInt16();
        return new ImgSize(width, height);
    }

    private static ImgSize DecodePng(BinaryReader binaryReader)
    {
        binaryReader.ReadBytes(8);
        int width = ReadLittleEndianInt32(binaryReader);
        int height = ReadLittleEndianInt32(binaryReader);
        return new ImgSize(width, height);
    }

    private static ImgSize DecodeJfif(BinaryReader binaryReader)
    {
        while (binaryReader.ReadByte() == 0xff)
        {
            byte marker = binaryReader.ReadByte();
            short chunkLength = ReadLittleEndianInt16(binaryReader);
            if (marker == 0xc0)
            {
                binaryReader.ReadByte();
                int height = ReadLittleEndianInt16(binaryReader);
                int width = ReadLittleEndianInt16(binaryReader);
                return new ImgSize(width, height);
            }

            if (chunkLength < 0)
            {
                ushort uchunkLength = (ushort)chunkLength;
                binaryReader.ReadBytes(uchunkLength - 2);
            }
            else
            {
                binaryReader.ReadBytes(chunkLength - 2);
            }
        }
        return new ImgSize(0, 0);
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
