using Microsoft.SqlServer.Server;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlTypes;
using System.DirectoryServices;
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

        System.IO.StreamWriter file = Util.CreateLogFile();

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

            // Create key/value collection - key is (user) distinguishedname, value is object GUID.
            Dictionary<string, Guid> UserDStoGUID = new Dictionary<string, Guid>();

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
                            PropertyValueCollection prop = Util.GetADproperty(item, coldef.ADpropName);
                            if (prop != null)
                                row[i] = prop.Value;
                            break;

                        case "UAC":
                            if (Item_UAC_flags == null)
                            {   // Get UAC flags only once per AD object.
                                Item_UAC_flags = new UACflags(Util.Get_userAccountControl(item, out UserPasswordExpiryTimeComputed));
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
                                time = Util.GetFileTime(searchResult, coldef.ADpropName);
                            if(time > 0 && time != 0x7fffffffffffffff && time != -1)
                            {
                                //row[i] = DateTime.FromFileTimeUtc(time);
                                row[i] = DateTime.FromFileTime(time);       // Convert UTC to local time.
                            }
                            break;

                        case "SID":
                            row[i] = Util.GetSID(item, coldef.ADpropName);
                            break;

                        case "GrpCat":
                            if (ADGroupType == null)
                                ADGroupType = Util.GetADproperty(item, "grouptype");
                            row[i] = Util.GetGroupCategory(ADGroupType);
                            break;

                        case "GrpScope":
                            if (ADGroupType == null)
                                ADGroupType = Util.GetADproperty(item, "grouptype");
                            row[i] = Util.GetGroupScope(ADGroupType);
                            break;
                    }
                }
                tbl.Rows.Add(row);

                if (TblData.IsUser)
                {
                    // Set UserMustChangePasswordAtNextLogon column value (for user objects).
                    bool IsUsrChgPwd = false;
                    if (row.IsNull("PasswordLastSet") && !(bool)row["PasswordNeverExpires"]
                        && !(bool)row["PasswordNotRequired"])
                        IsUsrChgPwd = true;
                    row["UserMustChangePasswordAtNextLogon"] = IsUsrChgPwd;

                    // Collect user distinguishedname into dictionary, value is object GUID.
                    // This is needed later to set ManagerGUID column.
                    UserDStoGUID.Add((string)row["distinguishedname"], (Guid)row["ObjectGUID"]);
                }

                // Save group members into the Xml document.
                if (TblData.IsGroup && item.Properties.Contains("member"))
                {
                    PropertyValueCollection coll = Util.GetADproperty(item, "member");
                    string parent = (string)row["distinguishedname"];
                    Util.SaveGroupMembersToXml(doc, body, parent, coll);
                }
            }
            // All rows have been added to the dataset.

            // set ManagerGUID column for user objects.
            if (TblData.IsUser)
            {
                foreach (DataRow rowUsr in tbl.Rows)
                {
                    object manager = rowUsr["Manager"]; // distinguishedname of Manager.
                    if (manager == DBNull.Value)
                        continue;
                    Guid ManagerGUID;
                    if (UserDStoGUID.TryGetValue((string)manager, out ManagerGUID))
                        rowUsr["ManagerGUID"] = ManagerGUID;
                }
            }

            // Return set dataset to SQL server.
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
        System.IO.StreamWriter file = Util.CreateLogFile();

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

                PropertyValueCollection prop = Util.GetADproperty(item, "thumbnailphoto");
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
                    SqlContext.Pipe.Send("Warning: Get image size failed for user (" + Util.GetDistinguishedName(item) + ")"
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
}   // endof: StoredProcedures partial class
