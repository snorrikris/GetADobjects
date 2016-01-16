using Microsoft.SqlServer.Server;
using System;
using System.DirectoryServices;
using System.IO;
using System.Security.Principal;
using System.Xml;

public static class Util
{
    //public static System.IO.StreamWriter CreateLogFile()
    //{
    //    System.IO.StreamWriter file = null;
    //    try
    //    {
    //        string folder = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);

    //        // Combine the base folder with your specific folder....
    //        string specificFolder = Path.Combine(folder, "GetADobjects");

    //        // Check if folder exists and if not, create it
    //        if (!Directory.Exists(specificFolder))
    //            Directory.CreateDirectory(specificFolder);

    //        string filename = Path.Combine(specificFolder, "Log.txt");

    //        file = new System.IO.StreamWriter(filename);
    //    }
    //    catch (Exception ex)
    //    {
    //        string Msg = "Exception in CreateLogFile: " + ex.Message;
    //    }
    //    return file;
    //}

    public static PropertyValueCollection GetADproperty(DirectoryEntry item, string ADpropName)
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
    public static Int32 Get_userAccountControl(DirectoryEntry item, out Int64 PwdExpComputed)
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
            if (res.Properties.Contains("userAccountControl"))
                AC1 = Convert.ToInt32(res.Properties["userAccountControl"][0]);
            if (res.Properties.Contains("msDS-User-Account-Control-Computed"))
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
    /// and parses the ACE strings for "Change password" permission.
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
            // Get SDDL string for "The discretionary access control list" (DACL) section of the security descriptor from "user".
            ActiveDirectorySecurity userSecurity = user.ObjectSecurity;
            string userSDDL = userSecurity.GetSecurityDescriptorSddlForm(System.Security.AccessControl.AccessControlSections.Access);

            // Find ACE strings for "Change password" permission on "this" user.
            int ix_chgpwdGUID = -1, startIdx = 0;
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
                //if (ix_chgpwdGUID > 0)
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
            SqlContext.Pipe.Send("Warning: IsUserCannotChangePassword function failed for user (" + GetDistinguishedName(user) + ")"
                    + " Exception: " + ex.Message);
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

    public static Int64 GetFileTime(SearchResult sr, string ADpropName)
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

    public static string GetSID(DirectoryEntry item, string ADpropName)
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
    public static string GetDistinguishedName(DirectoryEntry item)
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
    public static string GetDistinguishedName(SearchResult sr)
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

    public static string GetGroupCategory(PropertyValueCollection ADgrptype)
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

    public static string GetGroupScope(PropertyValueCollection ADgrptype)
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
        catch (Exception ex)
        {
            SqlContext.Pipe.Send("Warning: SaveGroupMembersToXml (" + ParentDS + ") failed. Exception: " + ex.Message);
        }
    }
}   // endof: class ADhelpers
