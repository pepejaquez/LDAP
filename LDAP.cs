using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LDAPProperties
{
    class LDAP
    {

        DirectoryEntry de = new DirectoryEntry("LDAP://pdc1/OU=Students,dc=pencol,dc=local");

        // Additions to this list are the children of 'OU=Students' (e.g. "OU=CSE,OU=Students")
        public List<string> DirectoryChildren = new List<string>(){ "OU=Students" };

        string serverString = "LDAP://pdc1/";
        string location = ",dc=pencol,dc=local";

        public DataRow row;


        public LDAP()
        {
            // Populate the DirectoryChildren List with the 'OU' entries of OU=Students.
            foreach (DirectoryEntry entry in de.Children)
            {
                if (entry.Name.ToString().Substring(0, 2) == "OU")
                {
                    DirectoryChildren.Add(entry.Name + ",OU=Students");
                }
            }
        }
        
        public void GetNewStudents(Dictionary<string, bool> dict, DataTable dt, OleDbDataAdapter da, DataSet ds)
        {
            foreach (string s in DirectoryChildren)
            {
                DirectoryEntry de = new DirectoryEntry(serverString + s + location);

                foreach (DirectoryEntry entry in de.Children)
                {
                    if (!(entry.Properties["cn"].Value == null) && entry.Properties["cn"].Value.ToString()[0] == 's')
                    {
                        // Add the entries that do not already have their key in the dictionary.
                        if (!dict.ContainsKey(entry.Properties["cn"].Value.ToString().Substring(1, 6)))
                        {
                            row = dt.NewRow();
                            row.BeginEdit();
                            int totalRec = dt.Rows.Count;
                            int currRec = totalRec - 1;

                            // Strip the 's' off the id.
                            row["id"] = entry.Properties["cn"].Value.ToString().Substring(1, 6);

                            // Change the names to propercase.
                            row["lastName"] = ProperCase(entry.Properties["sn"].Value.ToString().ToLower());
                            row["firstName"] = ProperCase(entry.Properties["givenName"].Value.ToString().ToLower());
                            dt.Rows.Add(row);
                        }
                    }
                }                
            }
            da.Update(ds, "tblStudents");
        }

        private string ProperCase(string str)
        {
            StringBuilder sb = new StringBuilder(str);

            // Change the first character to uppercase.
            sb[0] = char.ToUpper(sb[0]);

            // Change the third character to uppercase. (e.g. Mccoy is now McCoy)
            if (str.Substring(0, 2) == "mc")
            {                  
                sb[2] = char.ToUpper(sb[2]);
            }

            // Change the character after the space to uppercase. (e.g. Bickle biggs is now Bickle Biggs)
            if (str.Contains(' '))
            {
                sb[str.IndexOf(' ') + 1] = char.ToUpper(sb[str.IndexOf(' ') + 1]);
            }

            // Change the character after the hyphen to uppercase. (e.g. Graham-harvey is now Graham-Harvey)
            if (str.Contains('-'))
            {
                sb[str.IndexOf('-') + 1] = char.ToUpper(sb[str.IndexOf('-') + 1]);
            }
            
            return sb.ToString();
        }        
    }
}
