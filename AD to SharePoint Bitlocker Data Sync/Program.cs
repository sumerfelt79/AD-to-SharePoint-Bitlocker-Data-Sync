using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AD_to_SharePoint_Bitlocker_Data_Sync
{
    public class ComputerInfo
    {
        public string computerName { get; set; }
        public string computerOwner { get; set; }
        public static List<ComputerInfo> computerNames = new List<ComputerInfo>();
    }
    class Program
    {
        static void Main(string[] args)
        {
            SharePoint spMethods = new SharePoint();
            ActiveDirectory adMethods = new ActiveDirectory();
            spMethods.getComputerAssets();
            foreach (ComputerInfo c in ComputerInfo.computerNames)
            {
                BitlockerData bitlockerData =  adMethods.getBitlockerInfo(c);
                if (bitlockerData.recoveryGuid != null)
                {
                    spMethods.updateBitlockerList(c, bitlockerData);
                }
            }
        }
    }
}
