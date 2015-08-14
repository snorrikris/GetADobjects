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
                    XmlElement group = doc.CreateElement(string.Empty, "Group", string.Empty);
                    group.SetAttribute("GrpDS", parent);
                    body.AppendChild(group);

                    foreach (Object obj in coll)
                    {
                        string GrpMember = (string)obj;
                        XmlElement member = doc.CreateElement(string.Empty, "Member", string.Empty);
                        member.SetAttribute("MemberDS", GrpMember);
                        group.AppendChild(member);
                    }
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
            // To prevent memory leaks, always call 
            // SearchResultCollection.Dispose() manually.
            if (null != results)
            {
                results.Dispose();
                results = null;
            }
        }
        file.Close();
    }   // endof: clr_GetADobjects

    [Microsoft.SqlServer.Server.SqlProcedure]
    public static void clr_GetADusersEx()
    {
        System.IO.StreamWriter file = CreateLogFile();

        SearchResultCollection results = null;

        try
        {
            //GroupsTableEx GroupsTblData = new GroupsTableEx();
            //DataTable tbl = GroupsTblData.CreateTable();

            string path = "LDAP://DC=veca,DC=is";
            DirectoryEntry entry = new DirectoryEntry(path);

            DirectorySearcher mySearcher = new DirectorySearcher(entry);

            //mySearcher.Filter = "(objectCategory=contact)";
            //mySearcher.Filter = "(&(objectCategory=person)(objectClass=contact)(|(sn=Smith)(sn=Johnson)))";
            mySearcher.Filter = "(&(objectCategory=person)(objectClass=user))";

            mySearcher.PageSize = 500;

            results = mySearcher.FindAll();

            //DataRow row;

            // Iterate through each SearchResult in the SearchResultCollection.
            foreach (SearchResult searchResult in results)
            {
                DirectoryEntry item = searchResult.GetDirectoryEntry();
                //row = tbl.NewRow();

                //row[11] = item.SchemaClassName; // "group";
                //row[12] = item.Guid;

                //SecurityIdentifier sid = new SecurityIdentifier(item.Properties["objectSid"][0] as byte[], 0);
                //row[14] = sid.ToString();

                //if (item.Properties.Contains("grouptype"))
                //    grouptype = (Int32)item.Properties["grouptype"].Value;

                //for (int i = 0; i < GroupsTblData.collist.Length; i++)
                //{
                //    TableColDef coldef = GroupsTblData.collist[i];
                //    if (coldef.IsMethod)
                //        continue;
                //    if (item.Properties.Contains(coldef.ADpropName))
                //    {
                //        try
                //        {
                //            row[i] = item.Properties[coldef.ADpropName].Value;
                //        }
                //        catch (Exception ex)
                //        {
                //            file.WriteLine("Exception on AD property (" + coldef.ADpropName + "). Error: " + ex.Message);
                //        }
                //    }
                //    else
                //    {
                //        file.WriteLine("Missing property (" + coldef.ADpropName + ") on group ");
                //    }
                //}
                //tbl.Rows.Add(row);

                // Iterate through each property name in each SearchResult.
                foreach (string propertyKey in searchResult.Properties.PropertyNames)
                {
                    // Retrieve the value assigned to that property name 
                    // in the ResultPropertyValueCollection.
                    ResultPropertyValueCollection valueCollection =
                        searchResult.Properties[propertyKey];

                    // Iterate through values for each property name in each 
                    // SearchResult.
                    foreach (Object propertyValue in valueCollection)
                    {
                        // Handle results. Be aware that the following 
                        // WriteLine only returns readable results for 
                        // properties that are strings.
                        file.WriteLine("{0}:{1}", propertyKey, propertyValue.ToString());
                    }
                }
            }
            //DataSetUtilities.SendDataTable(tbl);
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
            // To prevent memory leaks, always call 
            // SearchResultCollection.Dispose() manually.
            if (null != results)
            {
                results.Dispose();
                results = null;
            }
        }

        file.Close();
    }   // endof: clr_GetADusersEx

    [Microsoft.SqlServer.Server.SqlProcedure]
    public static void clr_GetADcontactsEx()
    {
        System.IO.StreamWriter file = CreateLogFile();

        SearchResultCollection results = null;

        try
        {
            //GroupsTableEx GroupsTblData = new GroupsTableEx();
            //DataTable tbl = GroupsTblData.CreateTable();

            string path = "LDAP://DC=veca,DC=is";
            DirectoryEntry entry = new DirectoryEntry(path);

            DirectorySearcher mySearcher = new DirectorySearcher(entry);

            //mySearcher.Filter = "(objectCategory=contact)";
            //mySearcher.Filter = "(&(objectCategory=person)(objectClass=contact)(|(sn=Smith)(sn=Johnson)))";
            mySearcher.Filter = "(&(objectCategory=person)(objectClass=contact))";

            mySearcher.PageSize = 500;

            results = mySearcher.FindAll();

            //DataRow row;

            // Iterate through each SearchResult in the SearchResultCollection.
            foreach (SearchResult searchResult in results)
            {
                DirectoryEntry item = searchResult.GetDirectoryEntry();
                //row = tbl.NewRow();

                //row[11] = item.SchemaClassName; // "group";
                //row[12] = item.Guid;

                //SecurityIdentifier sid = new SecurityIdentifier(item.Properties["objectSid"][0] as byte[], 0);
                //row[14] = sid.ToString();

                //if (item.Properties.Contains("grouptype"))
                //    grouptype = (Int32)item.Properties["grouptype"].Value;

                //for (int i = 0; i < GroupsTblData.collist.Length; i++)
                //{
                //    TableColDef coldef = GroupsTblData.collist[i];
                //    if (coldef.IsMethod)
                //        continue;
                //    if (item.Properties.Contains(coldef.ADpropName))
                //    {
                //        try
                //        {
                //            row[i] = item.Properties[coldef.ADpropName].Value;
                //        }
                //        catch (Exception ex)
                //        {
                //            file.WriteLine("Exception on AD property (" + coldef.ADpropName + "). Error: " + ex.Message);
                //        }
                //    }
                //    else
                //    {
                //        file.WriteLine("Missing property (" + coldef.ADpropName + ") on group ");
                //    }
                //}
                //tbl.Rows.Add(row);

                // Iterate through each property name in each SearchResult.
                foreach (string propertyKey in searchResult.Properties.PropertyNames)
                {
                    // Retrieve the value assigned to that property name 
                    // in the ResultPropertyValueCollection.
                    ResultPropertyValueCollection valueCollection =
                        searchResult.Properties[propertyKey];

                    // Iterate through values for each property name in each 
                    // SearchResult.
                    foreach (Object propertyValue in valueCollection)
                    {
                        // Handle results. Be aware that the following 
                        // WriteLine only returns readable results for 
                        // properties that are strings.
                        file.WriteLine("{0}:{1}", propertyKey, propertyValue.ToString());
                    }
                }
            }
            //DataSetUtilities.SendDataTable(tbl);
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
            // To prevent memory leaks, always call 
            // SearchResultCollection.Dispose() manually.
            if (null != results)
            {
                results.Dispose();
                results = null;
            }
        }

        file.Close();
    }   // endof: clr_GetADcontactsEx

    [Microsoft.SqlServer.Server.SqlProcedure]
    public static void clr_GetADcomputersEx()
    {
        System.IO.StreamWriter file = CreateLogFile();

        SearchResultCollection results = null;

        try
        {
            ComputerTableEx TblData = new ComputerTableEx();
            DataTable tbl = TblData.CreateTable();

            string path = "LDAP://DC=veca,DC=is";
            DirectoryEntry entry = new DirectoryEntry(path);

            DirectorySearcher mySearcher = new DirectorySearcher(entry);

            mySearcher.Filter = "(objectCategory=computer)";

            mySearcher.PageSize = 500;

            results = mySearcher.FindAll();

            DataRow row;

            // Iterate through each SearchResult in the SearchResultCollection.
            foreach (SearchResult searchResult in results)
            {
                DirectoryEntry item = searchResult.GetDirectoryEntry();
                row = tbl.NewRow();

                //if (item.Properties.Contains("grouptype"))
                //    grouptype = (Int32)item.Properties["grouptype"].Value;

                UACflags Item_UAC_flags = null;

                for (int i = 0; i < TblData.collist.Length; i++)
                {
                    TableColDefEx coldef = TblData.collist[i];
                    switch(coldef.OPtype)
                    {
                        case "Adprop":
                            if (item.Properties.Contains(coldef.ADpropName))
                            {
                                try
                                {
                                    row[i] = item.Properties[coldef.ADpropName].Value;
                                }
                                catch (Exception ex)
                                {
                                    file.WriteLine("Exception on AD property (" + coldef.ADpropName + "). Error: " + ex.Message);
                                }
                            }
                            break;
                        case "UAC":
                            if (Item_UAC_flags == null)
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
                                        file.WriteLine("Exception on AD property (" + coldef.ADpropName + "). Error: " + ex.Message);
                                    }
                                }
                                Item_UAC_flags = new UACflags(uac);
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
                            if (searchResult.Properties.Contains(coldef.ADpropName))
                            {
                                try
                                {
                                    time = (Int64)searchResult.Properties[coldef.ADpropName][0];
                                    if(time > 0 && time != 0x7fffffffffffffff)
                                    {
                                        row[i] = DateTime.FromFileTimeUtc(time);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    file.WriteLine("Exception on AD property (" + coldef.ADpropName + "). Error: " + ex.Message);
                                }
                            }
                            break;
                        case "SID":
                            if (item.Properties.Contains(coldef.ADpropName))
                            {
                                try
                                {
                                    SecurityIdentifier sid = new SecurityIdentifier(
                                        item.Properties[coldef.ADpropName][0] as byte[], 0);
                                    row[i] = sid.ToString();
                                }
                                catch (Exception ex)
                                {
                                    file.WriteLine("Exception on AD property (" + coldef.ADpropName + "). Error: " + ex.Message);
                                }
                            }
                            break;
                    }
                }
                tbl.Rows.Add(row);

                // Iterate through each property name in each SearchResult.
                //foreach (string propertyKey in searchResult.Properties.PropertyNames)
                //{
                //    // Retrieve the value assigned to that property name 
                //    // in the ResultPropertyValueCollection.
                //    ResultPropertyValueCollection valueCollection =
                //        searchResult.Properties[propertyKey];

                //    // Iterate through values for each property name in each 
                //    // SearchResult.
                //    foreach (Object propertyValue in valueCollection)
                //    {
                //        // Handle results. Be aware that the following 
                //        // WriteLine only returns readable results for 
                //        // properties that are strings.
                //        file.WriteLine("{0}:{1}", propertyKey, propertyValue.ToString());
                //    }
                //}
            }
            DataSetUtilities.SendDataTable(tbl);
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
            // To prevent memory leaks, always call 
            // SearchResultCollection.Dispose() manually.
            if (null != results)
            {
                results.Dispose();
                results = null;
            }
        }

        file.Close();
    }   // endof: clr_GetADcomputersEx

    [Microsoft.SqlServer.Server.SqlProcedure]
    public static void clr_GetADgroupsEx(out SqlXml MemberList)
    {
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

            GroupsTableEx GroupsTblData = new GroupsTableEx();
            DataTable tbl = GroupsTblData.CreateTable();

            string path = "LDAP://DC=veca,DC=is";
            DirectoryEntry entry = new DirectoryEntry(path);

            DirectorySearcher mySearcher = new DirectorySearcher(entry);

            mySearcher.Filter = "(objectCategory=group)";

            mySearcher.PageSize = 500;

            results = mySearcher.FindAll();

            DataRow row;

            // Iterate through each SearchResult in the SearchResultCollection.
            foreach (SearchResult searchResult in results)
            {
                DirectoryEntry item = searchResult.GetDirectoryEntry();
                row = tbl.NewRow();

                row[11] = item.SchemaClassName; // "group";
                row[12] = item.Guid;

                SecurityIdentifier sid = new SecurityIdentifier(item.Properties["objectSid"][0] as byte[], 0);
                row[14] = sid.ToString();

                // Source: https://msdn.microsoft.com/en-us/library/ms675935(v=vs.85).aspx
                Int32 grouptype = 0;
                if (item.Properties.Contains("grouptype"))
                    grouptype = (Int32)item.Properties["grouptype"].Value;
                string GroupCategory = "Distribution", GroupScope = "";
                bool IsSecurityGroup = ((grouptype & 0x80000000) > 0) ? true : false;
                if(IsSecurityGroup)
                    GroupCategory = "Security";
                if((grouptype & 0x1) == 0x1)
                    GroupScope = "BuiltIn";
                else if((grouptype & 0x2) == 0x2)
                    GroupScope = "Global";
                else if((grouptype & 0x4) == 0x4)
                    GroupScope = "DomainLocal";
                else if((grouptype & 0x8) == 0x8)
                    GroupScope = "Universal";
                row["GroupCategory"] = GroupCategory;
                row["GroupScope"] = GroupScope;

                for (int i = 0; i < GroupsTblData.collist.Length; i++)
                {
                    TableColDef coldef = GroupsTblData.collist[i];
                    if (coldef.IsMethod)
                        continue;
                    if (item.Properties.Contains(coldef.ADpropName))
                    {
                        try
                        {
                            row[i] = item.Properties[coldef.ADpropName].Value;
                        }
                        catch (Exception ex)
                        {
                            file.WriteLine("Exception on AD property (" + coldef.ADpropName + "). Error: " + ex.Message);
                        }
                    }
                    else
                    {
                        file.WriteLine("Missing property (" + coldef.ADpropName + ") on group ");
                    }
                }
                tbl.Rows.Add(row);

                if (item.Properties.Contains("member"))
                {
                    System.DirectoryServices.PropertyValueCollection coll =
                        item.Properties["member"];
                    string parent = (string)row["distinguishedname"];
                    XmlElement group = doc.CreateElement(string.Empty, "Group", string.Empty);
                    group.SetAttribute("GrpDS", parent);
                    body.AppendChild(group);

                    foreach(Object obj in coll)
                    {
                        string GrpMember = (string)obj;
                        XmlElement member = doc.CreateElement(string.Empty, "Member", string.Empty);
                        member.SetAttribute("MemberDS", GrpMember);
                        group.AppendChild(member);
                    }
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
            // To prevent memory leaks, always call 
            // SearchResultCollection.Dispose() manually.
            if (null != results)
            {
                results.Dispose();
                results = null;
            }
        }

        file.Close();
    }   // endof: clr_GetADgroupsEx

    [Microsoft.SqlServer.Server.SqlProcedure]
    public static void clr_GetADusers ()
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
            UserTable UserTblData = new UserTable();
            DataTable tbl = UserTblData.CreateTable();

            PrincipalContext oPrincipalContext = new PrincipalContext(ContextType.Domain, "veca.is", "DC=veca,DC=is", ContextOptions.Negotiate);

            UserPrincipal up = new UserPrincipal(oPrincipalContext);

            up.SamAccountName = "*";

            PrincipalSearcher ps = new PrincipalSearcher();
            ps.QueryFilter = up;

            // Set page size to overcome the 1000 items limit.
            DirectorySearcher dsrch = (DirectorySearcher)ps.GetUnderlyingSearcher();
            dsrch.PageSize = 500;

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

            // Set page size to overcome the 1000 items limit.
            DirectorySearcher dsrch = (DirectorySearcher)ps.GetUnderlyingSearcher();
            dsrch.PageSize = 500;

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

    [Microsoft.SqlServer.Server.SqlProcedure]
    public static void clr_GetADcomputers()
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
            ComputersTable CompuersTblData = new ComputersTable();
            DataTable tbl = CompuersTblData.CreateTable();

            PrincipalContext oPrincipalContext = new PrincipalContext(ContextType.Domain, "veca.is", "DC=veca,DC=is", ContextOptions.Negotiate);

            ComputerPrincipal up = new ComputerPrincipal(oPrincipalContext);

            up.SamAccountName = "*";

            PrincipalSearcher ps = new PrincipalSearcher();
            ps.QueryFilter = up;

            // Set page size to overcome the 1000 items limit.
            DirectorySearcher dsrch = (DirectorySearcher)ps.GetUnderlyingSearcher();
            dsrch.PageSize = 500;

            PrincipalSearchResult<Principal> results = ps.FindAll();

            if (results != null)
            {
                DataRow row;

                foreach (ComputerPrincipal computer in results)
                {
                    row = tbl.NewRow();

                    // COPY CODE from "AD_DW_Users table map to .net V4.xlsx".
                    // Filter on Method=true, IsUACflag=no
                    // Copy/Paste all filtered cells from "Methods" column.
                    if (computer.AccountExpirationDate != null) row[0] = computer.AccountExpirationDate;
                    if (computer.AccountLockoutTime != null) row[1] = computer.AccountLockoutTime;
                    row[3] = computer.AllowReversiblePasswordEncryption;
                    row[4] = computer.BadLogonCount;
                    row[5] = computer.UserCannotChangePassword;
                    row[8] = computer.Description;
                    row[9] = computer.DisplayName;
                    row[10] = computer.DistinguishedName;
                    row[13] = computer.Enabled;
                    if (computer.LastBadPasswordAttempt != null) row[14] = computer.LastBadPasswordAttempt;
                    if (computer.LastLogon != null) row[15] = computer.LastLogon;
                    row[17] = computer.IsAccountLockedOut();
                    row[20] = computer.Name;
                    row[22] = computer.StructuralObjectClass;
                    row[23] = computer.Guid;
                    if (computer.LastPasswordSet != null) row[25] = computer.LastPasswordSet;
                    row[26] = computer.PasswordNeverExpires;
                    row[27] = computer.PasswordNotRequired;
                    row[29] = computer.SamAccountName;
                    row[30] = computer.Sid;
                    row[31] = computer.SmartcardLogonRequired;
                    row[35] = computer.UserPrincipalName;


                    bool HasPropList = false;
                    DirectoryEntry directoryEntry = computer.GetUnderlyingObject() as DirectoryEntry;
                    for (int i = 0; i < CompuersTblData.collist.Length; i++)
                    {
                        TableColDef coldef = CompuersTblData.collist[i];
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
                            file.WriteLine("Missing property (" + coldef.ADpropName + ") on computer " + computer.SamAccountName);
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
                    row[12] = ((uac & DONT_REQ_PREAUTH) != 0) ? true : false;
                    row[24] = ((uac & PASSWORD_EXPIRED) != 0) ? true : false;
                    row[32] = ((uac & TRUSTED_FOR_DELEGATION) != 0) ? true : false;
                    row[33] = ((uac & TRUSTED_TO_AUTH_FOR_DELEGATION) != 0) ? true : false;
                    row[34] = ((uac & USE_DES_KEY_ONLY) != 0) ? true : false;


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
    }   // endof: clr_GetADcomputers

    [Microsoft.SqlServer.Server.SqlProcedure]
    public static void clr_GetADgroups()
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
            GroupsTable GroupsTblData = new GroupsTable();
            DataTable tbl = GroupsTblData.CreateTable();

            PrincipalContext oPrincipalContext = new PrincipalContext(ContextType.Domain, "veca.is", "DC=veca,DC=is", ContextOptions.Negotiate);

            GroupPrincipal up = new GroupPrincipal(oPrincipalContext);

            up.SamAccountName = "*";

            PrincipalSearcher ps = new PrincipalSearcher();
            ps.QueryFilter = up;

            // Set page size to overcome the 1000 items limit.
            DirectorySearcher dsrch = (DirectorySearcher)ps.GetUnderlyingSearcher();
            dsrch.PageSize = 500;

            PrincipalSearchResult<Principal> results = ps.FindAll();

            if (results != null)
            {
                DataRow row;

                foreach (GroupPrincipal group in results)
                {
                    row = tbl.NewRow();

                    // COPY CODE from "AD_DW_Users table map to .net V4.xlsx".
                    // Filter on Method=true, IsUACflag=no
                    // Copy/Paste all filtered cells from "Methods" column.
                    row[1] = group.Description;
                    row[2] = group.DisplayName;
                    row[3] = group.DistinguishedName;
                    row[7] = group.Name;
                    row[9] = group.StructuralObjectClass;
                    row[10] = group.Guid;
                    row[11] = group.SamAccountName;
                    row[12] = group.IsSecurityGroup;
                    row[13] = group.Sid;
                    row[14] = group.UserPrincipalName;


                    bool HasPropList = false;
                    DirectoryEntry directoryEntry = group.GetUnderlyingObject() as DirectoryEntry;
                    for (int i = 0; i < GroupsTblData.collist.Length; i++)
                    {
                        TableColDef coldef = GroupsTblData.collist[i];
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
                            file.WriteLine("Missing property (" + coldef.ADpropName + ") on group " + group.SamAccountName);
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
    }   // endof: clr_GetADcomputers

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
}   // endof: StoredProcedures partial class

public class UACflags
{
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
    //public string ObjType;
    public bool IsUser, IsContact, IsComputer, IsGroup;

    public ADcolsTable(string ADfilter)
    {
        IsUser = IsContact = IsComputer = IsGroup = false;
        // Use ADfilter parameter to determine the type of AD objects wanted.
        //ObjType = "";
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
        //ObjType = "user";
        IsUser = true;
        collist = new TableColDefEx[66];  // <-- SET number of elements to number of cells copied below.!

        // COPY CODE from "AD_DW_Users table map to .net V4.xlsx".
        // Copy/Paste all cells from "ColListDef" column.
        collist[0] = new TableColDefEx("AccountExpirationDate", typeof(DateTime), "accountexpires","filetime");
        collist[1] = new TableColDefEx("AccountLockoutTime", typeof(DateTime), "lockouttime","filetime");
        collist[2] = new TableColDefEx("AccountNotDelegated", typeof(Boolean), "NOT_DELEGATED","UAC");
        collist[3] = new TableColDefEx("AllowReversiblePasswordEncryption", typeof(Boolean), "ENCRYPTED_TEXT_PWD_ALLOWED","UAC");
        collist[4] = new TableColDefEx("BadLogonCount", typeof(Int32), "badpwdcount","Adprop");
        collist[5] = new TableColDefEx("CannotChangePassword", typeof(Boolean), "PASSWD_CANT_CHANGE","UAC");
        collist[6] = new TableColDefEx("City", typeof(String), "l","Adprop");
        collist[7] = new TableColDefEx("CN", typeof(String), "cn","Adprop");
        collist[8] = new TableColDefEx("Company", typeof(String), "company","Adprop");
        collist[9] = new TableColDefEx("Country", typeof(String), "co","Adprop");
        collist[10] = new TableColDefEx("Created", typeof(DateTime), "whencreated","Adprop");
        collist[11] = new TableColDefEx("Department", typeof(String), "department","Adprop");
        collist[12] = new TableColDefEx("Description", typeof(String), "description","Adprop");
        collist[13] = new TableColDefEx("DisplayName", typeof(String), "displayname","Adprop");
        collist[14] = new TableColDefEx("DistinguishedName", typeof(String), "distinguishedname","Adprop");
        collist[15] = new TableColDefEx("Division", typeof(String), "division","Adprop");
        collist[16] = new TableColDefEx("DoesNotRequirePreAuth", typeof(Boolean), "DONT_REQ_PREAUTH","UAC");
        collist[17] = new TableColDefEx("EmailAddress", typeof(String), "mail","Adprop");
        collist[18] = new TableColDefEx("EmployeeID", typeof(String), "employeeid","Adprop");
        collist[19] = new TableColDefEx("EmployeeNumber", typeof(String), "employeenumber","Adprop");
        collist[20] = new TableColDefEx("Enabled", typeof(Boolean), "ACCOUNTDISABLE","UAC");
        collist[21] = new TableColDefEx("Fax", typeof(String), "facsimileTelephoneNumber","Adprop");
        collist[22] = new TableColDefEx("GivenName", typeof(String), "givenname","Adprop");
        collist[23] = new TableColDefEx("HomeDirectory", typeof(String), "homedirectory","Adprop");
        collist[24] = new TableColDefEx("HomedirRequired", typeof(Boolean), "HOMEDIR_REQUIRED","UAC");
        collist[25] = new TableColDefEx("HomeDrive", typeof(String), "homedrive","Adprop");
        collist[26] = new TableColDefEx("HomePage", typeof(String), "wwwhomepage","Adprop");
        collist[27] = new TableColDefEx("HomePhone", typeof(String), "homephone","Adprop");
        collist[28] = new TableColDefEx("Initials", typeof(String), "initials","Adprop");
        collist[29] = new TableColDefEx("LastBadPasswordAttempt", typeof(DateTime), "badpasswordtime","filetime");
        collist[30] = new TableColDefEx("LastLogonDate", typeof(DateTime), "lastlogon","filetime");
        collist[31] = new TableColDefEx("LockedOut", typeof(Boolean), "LOCKOUT","UAC");
        collist[32] = new TableColDefEx("LogonCount", typeof(Int32), "logoncount","Adprop");
        collist[33] = new TableColDefEx("LogonWorkstations", typeof(String), "userworkstations","Adprop");
        collist[34] = new TableColDefEx("Manager", typeof(String), "manager","Adprop");
        collist[35] = new TableColDefEx("MNSLogonAccount", typeof(Boolean), "MNS_LOGON_ACCOUNT","UAC");
        collist[36] = new TableColDefEx("MobilePhone", typeof(String), "mobile","Adprop");
        collist[37] = new TableColDefEx("Modified", typeof(DateTime), "whenchanged","Adprop");
        collist[38] = new TableColDefEx("Name", typeof(String), "name","Adprop");
        collist[39] = new TableColDefEx("ObjectCategory", typeof(String), "objectcategory","Adprop");
        collist[40] = new TableColDefEx("ObjectClass", typeof(String), "SchemaClassName","ObjClass");
        collist[41] = new TableColDefEx("ObjectGUID", typeof(Guid), "Guid","ObjGuid");
        collist[42] = new TableColDefEx("Office", typeof(String), "physicalDeliveryOfficeName","Adprop");
        collist[43] = new TableColDefEx("OfficePhone", typeof(String), "telephonenumber","Adprop");
        collist[44] = new TableColDefEx("Pager", typeof(String), "pager","Adprop");
        collist[45] = new TableColDefEx("PasswordExpired", typeof(Boolean), "PASSWORD_EXPIRED","UAC");
        collist[46] = new TableColDefEx("PasswordLastSet", typeof(DateTime), "pwdlastset","filetime");
        collist[47] = new TableColDefEx("PasswordNeverExpires", typeof(Boolean), "DONT_EXPIRE_PASSWD","UAC");
        collist[48] = new TableColDefEx("PasswordNotRequired", typeof(Boolean), "PASSWD_NOTREQD","UAC");
        collist[49] = new TableColDefEx("POBox", typeof(String), "postofficebox","Adprop");
        collist[50] = new TableColDefEx("PostalCode", typeof(String), "postalcode","Adprop");
        collist[51] = new TableColDefEx("PrimaryGroupID", typeof(Int32), "primarygroupid","Adprop");
        collist[52] = new TableColDefEx("ProfilePath", typeof(String), "profilepath","Adprop");
        collist[53] = new TableColDefEx("SamAccountName", typeof(String), "samaccountname","Adprop");
        collist[54] = new TableColDefEx("ScriptPath", typeof(String), "scriptpath","Adprop");
        collist[55] = new TableColDefEx("SID", typeof(String), "objectsid","SID");
        collist[56] = new TableColDefEx("SmartcardLogonRequired", typeof(Boolean), "SMARTCARD_REQUIRED","UAC");
        collist[57] = new TableColDefEx("State", typeof(String), "st","Adprop");
        collist[58] = new TableColDefEx("StreetAddress", typeof(String), "streetaddress","Adprop");
        collist[59] = new TableColDefEx("Surname", typeof(String), "sn","Adprop");
        collist[60] = new TableColDefEx("Title", typeof(String), "title","Adprop");
        collist[61] = new TableColDefEx("TrustedForDelegation", typeof(Boolean), "TRUSTED_FOR_DELEGATION","UAC");
        collist[62] = new TableColDefEx("TrustedToAuthForDelegation", typeof(Boolean), "TRUSTED_TO_AUTH_FOR_DELEGATION","UAC");
        collist[63] = new TableColDefEx("UseDESKeyOnly", typeof(Boolean), "USE_DES_KEY_ONLY","UAC");
        collist[64] = new TableColDefEx("UserPrincipalName", typeof(String), "userprincipalname","Adprop");
        collist[65] = new TableColDefEx("userAccountControl", typeof(Int32), "useraccountcontrol","Adprop");
    }

    private void MakeContactColList()
    {
        //ObjType = "contact";
        IsContact = true;
    }

    private void MakeComputerColList()
    {
        //ObjType = "computer";
        IsComputer = true;
        collist = new TableColDefEx[41];  // <-- SET number of elements to number of cells copied below.!

        // COPY CODE from "AD_DW_Users table map to .net V4.xlsx".
        // Copy/Paste all cells from "ColListDef" column.
        collist[0] = new TableColDefEx("AccountExpirationDate", typeof(DateTime), "accountexpires", "filetime");
        collist[1] = new TableColDefEx("AccountLockoutTime", typeof(DateTime), "lockouttime", "filetime");
        collist[2] = new TableColDefEx("AccountNotDelegated", typeof(Boolean), "NOT_DELEGATED", "UAC");
        collist[3] = new TableColDefEx("AllowReversiblePasswordEncryption", typeof(Boolean), "ENCRYPTED_TEXT_PWD_ALLOWED", "UAC");
        collist[4] = new TableColDefEx("BadLogonCount", typeof(Int32), "badpwdcount", "Adprop");
        collist[5] = new TableColDefEx("CannotChangePassword", typeof(Boolean), "PASSWD_CANT_CHANGE", "UAC");
        collist[6] = new TableColDefEx("CN", typeof(String), "cn", "Adprop");
        collist[7] = new TableColDefEx("Created", typeof(DateTime), "whenCreated", "Adprop");
        collist[8] = new TableColDefEx("Description", typeof(String), "description", "Adprop");
        collist[9] = new TableColDefEx("DisplayName", typeof(String), "displayname", "Adprop");
        collist[10] = new TableColDefEx("DistinguishedName", typeof(String), "distinguishedname", "Adprop");
        collist[11] = new TableColDefEx("DNSHostName", typeof(String), "dnshostname", "Adprop");
        collist[12] = new TableColDefEx("DoesNotRequirePreAuth", typeof(Boolean), "DONT_REQ_PREAUTH", "UAC");
        collist[13] = new TableColDefEx("Enabled", typeof(Boolean), "ACCOUNTDISABLE", "UAC");
        collist[14] = new TableColDefEx("LastBadPasswordAttempt", typeof(DateTime), "badpasswordtime", "filetime");
        collist[15] = new TableColDefEx("LastLogonDate", typeof(DateTime), "lastlogontimestamp/lastlogon", "filetime");
        collist[16] = new TableColDefEx("Location", typeof(String), "location", "Adprop");
        collist[17] = new TableColDefEx("LockedOut", typeof(Boolean), "LOCKOUT", "UAC");
        collist[18] = new TableColDefEx("logonCount", typeof(Int32), "logoncount", "Adprop");
        collist[19] = new TableColDefEx("ManagedBy", typeof(String), "managedby", "Adprop");
        collist[20] = new TableColDefEx("Modified", typeof(DateTime), "whenChanged", "Adprop");
        collist[21] = new TableColDefEx("MNSLogonAccount", typeof(Boolean), "MNS_LOGON_ACCOUNT", "UAC");
        collist[22] = new TableColDefEx("Name", typeof(String), "name", "Adprop");
        collist[23] = new TableColDefEx("ObjectCategory", typeof(String), "objectcategory", "Adprop");
        collist[24] = new TableColDefEx("ObjectClass", typeof(String), "SchemaClassName", "ObjClass");
        collist[25] = new TableColDefEx("ObjectGUID", typeof(Guid), "Guid", "ObjGuid");
        collist[26] = new TableColDefEx("OperatingSystem", typeof(String), "operatingsystem", "Adprop");
        collist[27] = new TableColDefEx("OperatingSystemServicePack", typeof(String), "operatingsystemservicepack", "Adprop");
        collist[28] = new TableColDefEx("OperatingSystemVersion", typeof(String), "operatingsystemversion", "Adprop");
        collist[29] = new TableColDefEx("PasswordExpired", typeof(Boolean), "PASSWORD_EXPIRED", "UAC");
        collist[30] = new TableColDefEx("PasswordLastSet", typeof(DateTime), "pwdlastset", "filetime");
        collist[31] = new TableColDefEx("PasswordNeverExpires", typeof(Boolean), "DONT_EXPIRE_PASSWD", "UAC");
        collist[32] = new TableColDefEx("PasswordNotRequired", typeof(Boolean), "PASSWD_NOTREQD", "UAC");
        collist[33] = new TableColDefEx("PrimaryGroupID", typeof(Int32), "primarygroupid", "Adprop");
        collist[34] = new TableColDefEx("SamAccountName", typeof(String), "samaccountname", "Adprop");
        collist[35] = new TableColDefEx("SID", typeof(String), "objectsid", "SID");
        collist[36] = new TableColDefEx("SmartcardLogonRequired", typeof(Boolean), "SMARTCARD_REQUIRED", "UAC");
        collist[37] = new TableColDefEx("TrustedForDelegation", typeof(Boolean), "TRUSTED_FOR_DELEGATION", "UAC");
        collist[38] = new TableColDefEx("TrustedToAuthForDelegation", typeof(Boolean), "TRUSTED_TO_AUTH_FOR_DELEGATION", "UAC");
        collist[39] = new TableColDefEx("UseDESKeyOnly", typeof(Boolean), "USE_DES_KEY_ONLY", "UAC");
        collist[40] = new TableColDefEx("userAccountControl", typeof(Int32), "useraccountcontrol", "Adprop");
    }

    private void MakeGroupColList()
    {
        //ObjType = "group";
        IsGroup = true;
        collist = new TableColDefEx[15];  // <-- SET number of elements to number of cells copied below.!

        // COPY CODE from "AD_DW_Users table map to .net V4.xlsx".
        // Copy/Paste all cells from "ColListDef" column.
        collist[0] = new TableColDefEx("Created", typeof(DateTime), "whenCreated", "Adprop");
        collist[1] = new TableColDefEx("Description", typeof(String), "description", "Adprop");
        collist[2] = new TableColDefEx("DisplayName", typeof(String), "displayname", "Adprop");
        collist[3] = new TableColDefEx("DistinguishedName", typeof(String), "distinguishedname", "Adprop");
        collist[4] = new TableColDefEx("EmailAddress", typeof(String), "mail", "Adprop");
        collist[5] = new TableColDefEx("GroupCategory", typeof(String), "grouptype", "GrpCat");
        collist[6] = new TableColDefEx("GroupScope", typeof(String), "grouptype", "GrpScope");
        collist[7] = new TableColDefEx("ManagedBy", typeof(String), "managedby", "Adprop");
        collist[8] = new TableColDefEx("Modified", typeof(DateTime), "whenChanged", "Adprop");
        collist[9] = new TableColDefEx("Name", typeof(String), "name", "Adprop");
        collist[10] = new TableColDefEx("ObjectCategory", typeof(String), "objectcategory", "Adprop");
        collist[11] = new TableColDefEx("ObjectClass", typeof(String), "SchemaClassName", "ObjClass");
        collist[12] = new TableColDefEx("ObjectGUID", typeof(Guid), "Guid", "ObjGuid");
        collist[13] = new TableColDefEx("SamAccountName", typeof(String), "samaccountname", "Adprop");
        collist[14] = new TableColDefEx("SID", typeof(String), "objectsid", "SID");
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

public class ComputerTableEx
{
    public TableColDefEx[] collist;

    public ComputerTableEx()
    {
        collist = new TableColDefEx[41];  // <-- SET number of elements to number of cells copied below.!

        // COPY CODE from "AD_DW_Users table map to .net V4.xlsx".
        // Copy/Paste all cells from "ColListDef" column.
        collist[0] = new TableColDefEx("AccountExpirationDate", typeof(DateTime), "accountexpires", "filetime");
        collist[1] = new TableColDefEx("AccountLockoutTime", typeof(DateTime), "lockouttime", "filetime");
        collist[2] = new TableColDefEx("AccountNotDelegated", typeof(Boolean), "NOT_DELEGATED", "UAC");
        collist[3] = new TableColDefEx("AllowReversiblePasswordEncryption", typeof(Boolean), "ENCRYPTED_TEXT_PWD_ALLOWED", "UAC");
        collist[4] = new TableColDefEx("BadLogonCount", typeof(Int32), "badpwdcount", "Adprop");
        collist[5] = new TableColDefEx("CannotChangePassword", typeof(Boolean), "PASSWD_CANT_CHANGE", "UAC");
        collist[6] = new TableColDefEx("CN", typeof(String), "cn", "Adprop");
        collist[7] = new TableColDefEx("Created", typeof(DateTime), "whenCreated", "Adprop");
        collist[8] = new TableColDefEx("Description", typeof(String), "description", "Adprop");
        collist[9] = new TableColDefEx("DisplayName", typeof(String), "displayname", "Adprop");
        collist[10] = new TableColDefEx("DistinguishedName", typeof(String), "distinguishedname", "Adprop");
        collist[11] = new TableColDefEx("DNSHostName", typeof(String), "dnshostname", "Adprop");
        collist[12] = new TableColDefEx("DoesNotRequirePreAuth", typeof(Boolean), "DONT_REQ_PREAUTH", "UAC");
        collist[13] = new TableColDefEx("Enabled", typeof(Boolean), "ACCOUNTDISABLE", "UAC");
        collist[14] = new TableColDefEx("LastBadPasswordAttempt", typeof(DateTime), "badpasswordtime", "filetime");
        collist[15] = new TableColDefEx("LastLogonDate", typeof(DateTime), "lastlogontimestamp/lastlogon", "filetime");
        collist[16] = new TableColDefEx("Location", typeof(String), "location", "Adprop");
        collist[17] = new TableColDefEx("LockedOut", typeof(Boolean), "LOCKOUT", "UAC");
        collist[18] = new TableColDefEx("logonCount", typeof(Int32), "logoncount", "Adprop");
        collist[19] = new TableColDefEx("ManagedBy", typeof(String), "managedby", "Adprop");
        collist[20] = new TableColDefEx("Modified", typeof(DateTime), "whenChanged", "Adprop");
        collist[21] = new TableColDefEx("MNSLogonAccount", typeof(Boolean), "MNS_LOGON_ACCOUNT", "UAC");
        collist[22] = new TableColDefEx("Name", typeof(String), "name", "Adprop");
        collist[23] = new TableColDefEx("ObjectCategory", typeof(String), "objectcategory", "Adprop");
        collist[24] = new TableColDefEx("ObjectClass", typeof(String), "SchemaClassName", "ObjClass");
        collist[25] = new TableColDefEx("ObjectGUID", typeof(Guid), "Guid", "ObjGuid");
        collist[26] = new TableColDefEx("OperatingSystem", typeof(String), "operatingsystem", "Adprop");
        collist[27] = new TableColDefEx("OperatingSystemServicePack", typeof(String), "operatingsystemservicepack", "Adprop");
        collist[28] = new TableColDefEx("OperatingSystemVersion", typeof(String), "operatingsystemversion", "Adprop");
        collist[29] = new TableColDefEx("PasswordExpired", typeof(Boolean), "PASSWORD_EXPIRED", "UAC");
        collist[30] = new TableColDefEx("PasswordLastSet", typeof(DateTime), "pwdlastset", "filetime");
        collist[31] = new TableColDefEx("PasswordNeverExpires", typeof(Boolean), "DONT_EXPIRE_PASSWD", "UAC");
        collist[32] = new TableColDefEx("PasswordNotRequired", typeof(Boolean), "PASSWD_NOTREQD", "UAC");
        collist[33] = new TableColDefEx("PrimaryGroupID", typeof(Int32), "primarygroupid", "Adprop");
        collist[34] = new TableColDefEx("SamAccountName", typeof(String), "samaccountname", "Adprop");
        collist[35] = new TableColDefEx("SID", typeof(String), "objectsid", "SID");
        collist[36] = new TableColDefEx("SmartcardLogonRequired", typeof(Boolean), "SMARTCARD_REQUIRED", "UAC");
        collist[37] = new TableColDefEx("TrustedForDelegation", typeof(Boolean), "TRUSTED_FOR_DELEGATION", "UAC");
        collist[38] = new TableColDefEx("TrustedToAuthForDelegation", typeof(Boolean), "TRUSTED_TO_AUTH_FOR_DELEGATION", "UAC");
        collist[39] = new TableColDefEx("UseDESKeyOnly", typeof(Boolean), "USE_DES_KEY_ONLY", "UAC");
        collist[40] = new TableColDefEx("userAccountControl", typeof(Int32), "useraccountcontrol", "Adprop");
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

public class ComputersTable
{
    public TableColDef[] collist;

    public ComputersTable()
    {
        collist = new TableColDef[37];  // <-- SET number of elements to number of cells copied below.!

        // COPY CODE from "AD_DW_Users table map to .net V4.xlsx".
        // Copy/Paste all cells from "ColListDef" column in "Contacts" sheet.
        collist[0] = new TableColDef("AccountExpirationDate", typeof(DateTime), "AccountExpirationDate", true);
        collist[1] = new TableColDef("AccountLockoutTime", typeof(DateTime), "AccountLockoutTime", true);
        collist[2] = new TableColDef("AccountNotDelegated", typeof(Boolean), "NOT_DELEGATED", true);
        collist[3] = new TableColDef("AllowReversiblePasswordEncryption", typeof(Boolean), "AllowReversiblePasswordEncryption", true);
        collist[4] = new TableColDef("BadLogonCount", typeof(Int32), "BadLogonCount", true);
        collist[5] = new TableColDef("CannotChangePassword", typeof(Boolean), "UserCannotChangePassword", true);
        collist[6] = new TableColDef("CN", typeof(String), "CN", false);
        collist[7] = new TableColDef("Created", typeof(DateTime), "whenCreated", false);
        collist[8] = new TableColDef("Description", typeof(String), "Description", true);
        collist[9] = new TableColDef("DisplayName", typeof(String), "DisplayName", true);
        collist[10] = new TableColDef("DistinguishedName", typeof(String), "DistinguishedName", true);
        collist[11] = new TableColDef("DNSHostName", typeof(String), "dNSHostName", false);
        collist[12] = new TableColDef("DoesNotRequirePreAuth", typeof(Boolean), "DONT_REQ_PREAUTH", true);
        collist[13] = new TableColDef("Enabled", typeof(Boolean), "Enabled", true);
        collist[14] = new TableColDef("LastBadPasswordAttempt", typeof(DateTime), "LastBadPasswordAttempt", true);
        collist[15] = new TableColDef("LastLogonDate", typeof(DateTime), "LastLogon", true);
        collist[16] = new TableColDef("Location", typeof(String), "location", false);
        collist[17] = new TableColDef("LockedOut", typeof(Boolean), "IsAccountLockedOut", true);
        collist[18] = new TableColDef("ManagedBy", typeof(String), "ManagedBy", false);
        collist[19] = new TableColDef("Modified", typeof(DateTime), "whenChanged", false);
        collist[20] = new TableColDef("Name", typeof(String), "Name", true);
        collist[21] = new TableColDef("ObjectCategory", typeof(String), "ObjectCategory", false);
        collist[22] = new TableColDef("ObjectClass", typeof(String), "StructuralObjectClass", true);
        collist[23] = new TableColDef("ObjectGUID", typeof(Guid), "Guid", true);
        collist[24] = new TableColDef("PasswordExpired", typeof(Boolean), "PASSWORD_EXPIRED", true);
        collist[25] = new TableColDef("PasswordLastSet", typeof(DateTime), "LastPasswordSet", true);
        collist[26] = new TableColDef("PasswordNeverExpires", typeof(Boolean), "PasswordNeverExpires", true);
        collist[27] = new TableColDef("PasswordNotRequired", typeof(Boolean), "PasswordNotRequired", true);
        collist[28] = new TableColDef("PrimaryGroupID", typeof(Int32), "PrimaryGroupID", false);
        collist[29] = new TableColDef("SamAccountName", typeof(String), "SamAccountName", true);
        collist[30] = new TableColDef("SID", typeof(String), "Sid", true);
        collist[31] = new TableColDef("SmartcardLogonRequired", typeof(Boolean), "SmartcardLogonRequired", true);
        collist[32] = new TableColDef("TrustedForDelegation", typeof(Boolean), "TRUSTED_FOR_DELEGATION", true);
        collist[33] = new TableColDef("TrustedToAuthForDelegation", typeof(Boolean), "TRUSTED_TO_AUTH_FOR_DELEGATION", true);
        collist[34] = new TableColDef("UseDESKeyOnly", typeof(Boolean), "USE_DES_KEY_ONLY", true);
        collist[35] = new TableColDef("UserPrincipalName", typeof(String), "UserPrincipalName", true);
        collist[36] = new TableColDef("userAccountControl", typeof(Int32), "userAccountControl", false);
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
}   // endof: ComputersTable class

public class GroupsTable
{
    public TableColDef[] collist;

    public GroupsTable()
    {
        collist = new TableColDef[15];  // <-- SET number of elements to number of cells copied below.!

        // COPY CODE from "AD_DW_Users table map to .net V4.xlsx".
        // Copy/Paste all cells from "ColListDef" column in "Contacts" sheet.
        collist[0] = new TableColDef("Created", typeof(DateTime), "whenCreated", false);
        collist[1] = new TableColDef("Description", typeof(String), "Description", true);
        collist[2] = new TableColDef("DisplayName", typeof(String), "DisplayName", true);
        collist[3] = new TableColDef("DistinguishedName", typeof(String), "DistinguishedName", true);
        collist[4] = new TableColDef("EmailAddress", typeof(String), "mail", false);
        collist[5] = new TableColDef("ManagedBy", typeof(String), "ManagedBy", false);
        collist[6] = new TableColDef("Modified", typeof(DateTime), "whenChanged", false);
        collist[7] = new TableColDef("Name", typeof(String), "Name", true);
        collist[8] = new TableColDef("ObjectCategory", typeof(String), "ObjectCategory", false);
        collist[9] = new TableColDef("ObjectClass", typeof(String), "StructuralObjectClass", true);
        collist[10] = new TableColDef("ObjectGUID", typeof(Guid), "Guid", true);
        collist[11] = new TableColDef("SamAccountName", typeof(String), "SamAccountName", true);
        collist[12] = new TableColDef("SecurityEnabled", typeof(Boolean), "IsSecurityGroup", true);
        collist[13] = new TableColDef("SID", typeof(String), "Sid", true);
        collist[14] = new TableColDef("UserPrincipalName", typeof(String), "UserPrincipalName", true);
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
}   // endof: GroupsTable class

public class GroupsTableEx
{
    public TableColDef[] collist;

    public GroupsTableEx()
    {
        collist = new TableColDef[15];  // <-- SET number of elements to number of cells copied below.!

        // COPY CODE from "AD_DW_Users table map to .net V4.xlsx".
        // Copy/Paste all cells from "ColListDef" column in "Contacts" sheet.
        collist[0] = new TableColDef("Created", typeof(DateTime), "whenCreated", false);
        collist[1] = new TableColDef("Description", typeof(String), "description", false);
        collist[2] = new TableColDef("DisplayName", typeof(String), "displayname", false);
        collist[3] = new TableColDef("DistinguishedName", typeof(String), "distinguishedname", false);
        collist[4] = new TableColDef("EmailAddress", typeof(String), "mail", false);
        collist[5] = new TableColDef("GroupCategory", typeof(String), "grouptype", true);
        collist[6] = new TableColDef("GroupScope", typeof(String), "grouptype", true);
        collist[7] = new TableColDef("ManagedBy", typeof(String), "managedby", false);
        collist[8] = new TableColDef("Modified", typeof(DateTime), "whenChanged", false);
        collist[9] = new TableColDef("Name", typeof(String), "name", false);
        collist[10] = new TableColDef("ObjectCategory", typeof(String), "objectcategory", false);
        collist[11] = new TableColDef("ObjectClass", typeof(String), "SchemaClassName", true);
        collist[12] = new TableColDef("ObjectGUID", typeof(Guid), "Guid", true);
        collist[13] = new TableColDef("SamAccountName", typeof(String), "samaccountname", false);
        collist[14] = new TableColDef("SID", typeof(String), "objectsid", true);
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
}   // endof: GroupsTableEx class

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

