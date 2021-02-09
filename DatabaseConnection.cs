using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LDAPProperties
{
    class DatabaseConnection
    {
        
        string connString = "Provider=Microsoft.ACE.OLEDB.12.0;data source=T:\\Laptops\\Laptop Database.accdb";

        public Dictionary<string, bool> studentKeys = new Dictionary<string, bool>();

        public OleDbConnection dbconn;
        public OleDbDataAdapter dataAdapter;
        public DataTable dataTable;       
        public DataSet ds;
        public int currRec = 0;
        public int totalRec = 0;

        public DatabaseConnection()
        {
            GetDatabaseData();
            InitializeCommands();
        }

        public void GetDatabaseData()
        {
            // Retrieve the student id's from the database to use as keys in the studentKeys dictionary.
            dbconn = new OleDbConnection(connString);
            string commandString = "SELECT * from tblStudents";
            dataAdapter = new OleDbDataAdapter(commandString, dbconn);
            ds = new DataSet();
            dataAdapter.Fill(ds, "tblStudents");

            dataTable = ds.Tables["tblStudents"];
            currRec = 0;
            totalRec = dataTable.Rows.Count;

            foreach(DataRow row in dataTable.Rows)
            {
                object data = row["id"];
                string strData = data.ToString();
                studentKeys.Add(strData, true);
            }            
        }  

        public void InitializeCommands()
        {
            dataAdapter.InsertCommand = dbconn.CreateCommand();
            dataAdapter.InsertCommand.CommandText = "INSERT INTO tblStudents(id, lastName, firstName) VALUES(@id, @lastName, @firstName)";
            AddParams(dataAdapter.InsertCommand, "id", "lastName", "firstName");
        }

        private void AddParams(OleDbCommand cmd, params string[] cols)
        {
            foreach(string col in cols)
            {
                cmd.Parameters.Add("@" + col, OleDbType.Char, 0, col);
            }
        }
    }
}
