using Microsoft.SqlServer.Server;
using System;
using System.Collections.Generic;

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

