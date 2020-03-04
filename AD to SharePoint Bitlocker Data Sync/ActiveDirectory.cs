using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices;
using System.Runtime.InteropServices;

namespace AD_to_SharePoint_Bitlocker_Data_Sync
{
    public class BitlockerData{
        public string recoveryPassword { get; set; }
        public string recoveryGuid { get; set; }
        public DateTime date { get; set; }
    }
    class ActiveDirectory
    {
        /*Connecting to Active Directory, searching for the current object's computerName property, determining if the AD computer object
         * has Bitlocker data associated with it, if Bitlocker data is present, the Bitlocker Recovery Key, Recovery Guid and Date are assigned to properties
         * of a BitlockerData object, that object is then returned to the calling function*/
        public BitlockerData getBitlockerInfo(ComputerInfo c)
        {
            BitlockerData bitlockerInfoObject = new BitlockerData();
            using (DirectoryEntry parent = new DirectoryEntry("LDAP://wcc.local:636"))
            {
                using (DirectorySearcher LdapSearcher = new DirectorySearcher(parent))
                {
                    LdapSearcher.Filter = string.Concat("(&(objectClass=computer)(name=", c.computerName, "))");
                    SearchResult srcComp = LdapSearcher.FindOne();
                    if (srcComp != null)
                    {
                        using (DirectoryEntry compEntry = srcComp.GetDirectoryEntry())
                        {
                            try
                            {
                                Object objValue = Marshal.BindToMoniker(srcComp.GetDirectoryEntry().Path.Replace("GC://", "LDAP://"));
                                Type tType = objValue.GetType();
                                tType.InvokeMember("Filter",
                                System.Reflection.BindingFlags.SetProperty | System.Reflection.BindingFlags.Public, null,
                                objValue, new Object[] { "msFVE-RecoveryInformation" });
                                foreach (Object obj in (IEnumerable)objValue)
                                {
                                    Guid gRecoveryGUID = new Guid((Byte[])obj.GetType().InvokeMember("msFVE-RecoveryGuid", System.Reflection.BindingFlags.GetProperty | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance, null, obj, null, null, null, null));
                                    string name = obj.GetType().InvokeMember("name", System.Reflection.BindingFlags.GetProperty | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance, null, obj, null, null, null, null).ToString();
                                    string dateString = name.Substring(3, name.IndexOf("T", System.StringComparison.Ordinal) - 3);
                                    DateTime date = Convert.ToDateTime(dateString);
                                    string dateOnly = date.ToString().Substring(0, date.ToString().IndexOf(" "));
                                    string time = name.Substring(name.IndexOf("T", System.StringComparison.Ordinal) + 1, name.IndexOf("{", System.StringComparison.Ordinal) - 20);
                                    string objTime = DateTime.Parse(time).ToString("h:mm:ss tt");
                                    time = objTime;
                                    DateTime dateTime = Convert.ToDateTime(dateOnly + " " + time);
                                    if (gRecoveryGUID != null)
                                    {
                                        bitlockerInfoObject.recoveryGuid = gRecoveryGUID.ToString().ToUpper();
                                        bitlockerInfoObject.recoveryPassword = obj.GetType().InvokeMember("msFVE-RecoveryPassword", System.Reflection.BindingFlags.GetProperty | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance, null, obj, null, null, null, null).ToString();
                                        bitlockerInfoObject.date = dateTime;
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                }
            }
            return bitlockerInfoObject;
        }
    }
}
