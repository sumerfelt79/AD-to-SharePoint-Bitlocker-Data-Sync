using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;

namespace AD_to_SharePoint_Bitlocker_Data_Sync
{
    class SharePoint
    {
        /*using the IT Assets SharePoint list as a data source, I'm creating a list of ComputerInfo objects,
         * that list will be iterated over to determined which IT Assets have Bitlocker information in AD*/
        public void getComputerAssets()
        {
            using (ClientContext context = new ClientContext("https://sharepoint.wilsonconst.com/"))
            {
                List assets = context.Web.Lists.GetByTitle("IT Assets");
                CamlQuery query = new CamlQuery();
                query.ViewXml = "<View><Query><OrderBy><FieldRef Name='Title' Ascending='True' /></OrderBy></Query></View>";
                ListItemCollection collection = assets.GetItems(query);
                context.Load(collection);
                context.ExecuteQuery();
                foreach (ListItem item in collection)
                {
                    if (item["Title"] != null && item["Title"].ToString() != "" && item["Status"].ToString().ToLower() == "active")
                    {
                        if (item["category"].ToString().ToLower() == "laptop" || item["category"].ToString().ToLower() == "desktop")
                        {
                            FieldLookupValue owner = item["Assigned"] as FieldLookupValue;
                            //determing if the Assigned column for current list is null
                            if (owner != null)
                            {
                                string computerOwner = owner.LookupValue;
                                if (computerOwner == "")
                                {
                                    computerOwner = "Not Assigned";
                                }
                                ComputerInfo active = new ComputerInfo { computerName = item["Title"].ToString(), computerOwner = computerOwner };
                                ComputerInfo.computerNames.Add(active);
                            }
                            else
                            {
                                ComputerInfo active = new ComputerInfo { computerName = item["Title"].ToString(), computerOwner = "" };
                                ComputerInfo.computerNames.Add(active);
                            }
                        }
                    }
                }
            }
        }
        //adding Bitlocker information to the Bitlocker list on SharePoint
        public void updateBitlockerList(ComputerInfo c, BitlockerData b)
        {
            /*determing if the Bitlocker list already contains the current iteration's data. This is determined using a CamlQuery
             * that filters out all list items except those share the computer name and owner of the of the current ComputerInfo object.
             * Then, we iterate over the collection of list items gathered with the CamlQuery, comapring the Date Added field the date property of
             * the BitlockerData object, if any of the dates in the list item collection match the date in the BitlockerData object, a exists bool is set to true,
             * and the data will not be added, if no match is found, a new list item will be added to the Bitlocker list*/
            using (ClientContext context = new ClientContext("https://sharepoint.wilsonconst.com/it-site"))
            {
                List assetsList = context.Web.Lists.GetByTitle("Bitlocker");
                CamlQuery query = new CamlQuery() { ViewXml = "<View><Query><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>" + c.computerName + "</Value></Eq><Eq><FieldRef Name='User2' /><Value Type='Text'>" + c.computerOwner + "</Value></Eq></And></Where></Query></View>" };
                ListItemCollection collection = assetsList.GetItems(query);
                context.Load(collection);
                context.ExecuteQuery();
                bool exists = false;
                foreach (ListItem i in collection)
                {
                    DateTime spDateTime = Convert.ToDateTime(i["Date_x0020_Added"].ToString()).ToLocalTime();
                    if (spDateTime == b.date)
                    {
                        exists = true;
                    }
                }

                if (!exists)
                {
                    ListItemCreationInformation creationInfo = new ListItemCreationInformation();
                    ListItem newItem = assetsList.AddItem(creationInfo);
                    newItem["Title"] = c.computerName;
                    newItem["User2"] = c.computerOwner;
                    newItem["Identifier"] = b.recoveryGuid;
                    newItem["Recovery_x0020_Key"] = b.recoveryPassword;
                    newItem["Date_x0020_Added"] = b.date;
                    newItem.Update();
                    context.ExecuteQuery();
                }
            }
        }
    }
}
