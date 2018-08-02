using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HizbeJamali.ZaereenDataImportApp
{
    /// <summary>
    /// Create seperate excel sheet and place the data in it, it should be in the below format.
    /// ITS, Name, Mobile, Age, Location, Occupation, TripExp, Remarks. In case there are extra fields, then the code should be amended before executing.
    /// it should start from the first column i.e.: A1
    /// first row should be the column names and should match the above names.
    /// 2 arguments should be passed to the command line (use command prompt to execute the application, enter the path of the application and pass the 2 parameters).
    /// i.e.: Excel Sheet (the one created above) path and the database (create a copy of production database on local) path
    /// both the paths should have double quotes around them and both the paths should be separated by space.
    /// Before attempting to run ensure the Account_No field in the production database is auto-generated, if not, then amend the code to generate account_no dynamically.
    /// Run the application and it will insert the records in the database.
    /// </summary>
    class ImportZaereenDataFromExcel
    {

        private DataTable GetDataFromExcelSheet(string excelFileLocation)
        {
            if (!string.IsNullOrEmpty(excelFileLocation))
            {
                DataTable newDt = new DataTable("StructuredData");
                var oleExcelConnectionString = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0;HDR=No;IMEX=1""", excelFileLocation);
                using (OleDbConnection oleExcelConnection = new OleDbConnection(oleExcelConnectionString))
                {
                    OleDbDataAdapter oleExcelDataAdapter = new OleDbDataAdapter("select * from [Sheet1$]", oleExcelConnection);
                    DataTable dt = new DataTable();
                    oleExcelDataAdapter.Fill(dt);

                    for (int row = 0; row < dt.Rows.Count; row++)
                    {
                        if (row == 0)
                        {
                            for (int col = 0; col < dt.Columns.Count; col++)
                            {
                                newDt.Columns.Add(dt.Rows[row][col].ToString());
                            }
                            continue;
                        }
                        DataRow _currentRow = newDt.NewRow();
                        _currentRow["ITS"] = dt.Rows[row]["F1"];
                        _currentRow["Name"] = dt.Rows[row]["F2"];
                        _currentRow["Mobile"] = dt.Rows[row]["F3"];
                        _currentRow["Age"] = dt.Rows[row]["F4"];
                        _currentRow["Location"] = dt.Rows[row]["F5"];
                        _currentRow["Occupation"] = dt.Rows[row]["F6"];
                        _currentRow["TripExp"] = dt.Rows[row]["F7"];
                        _currentRow["Remarks"] = dt.Rows[row]["F8"];
                        newDt.Rows.Add(_currentRow);
                    }

                }
                return newDt;
            }
            return null;
        }

        public int GetLastAccountNumber(string dbPath)
        {
            int account_number = 0;
            using (OleDbConnection connection = new OleDbConnection(string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};", dbPath)))
            {
                OleDbCommand command = new OleDbCommand("select top 1 Account_No from ZaereenLedger order by account_no desc", connection);
                command.CommandType = CommandType.Text;
                connection.Open();
                account_number = Convert.ToInt32(command.ExecuteScalar());
                connection.Close();
            }
            return account_number;
        }

        private void PushDataToAccessDatabase(string dbPath, DataTable dataToInsert)
        {
            using (OleDbConnection connection = new OleDbConnection(string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};", dbPath)))
            {
                int count = 1;
                foreach(DataRow row in dataToInsert.Rows)
                {
                    count++;
                    string insertQuery = string.Format("insert into zaereenledger (Account_No, Zaereen_Name, Age, Ejamaat, Mobile, Occupation, Address, TripExp, Remarks) values ('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', {6}, '{7}', '{8}')",
                        GetLastAccountNumber(dbPath) + count, row["Name"], row["Age"], row["ITS"], row["Mobile"], row["Occupation"], row["Location"], row["TripExp"], row["Remarks"]);

                    OleDbCommand command = new OleDbCommand(insertQuery, connection);
                    command.CommandType = CommandType.Text;
                    connection.Open();
                    command.ExecuteNonQuery();
                    connection.Close();
                }
            }
        }

        static void Main(string[] args)
        {
            ImportZaereenDataFromExcel p = new ImportZaereenDataFromExcel();
            DataTable dt = p.GetDataFromExcelSheet(args[0]);
            p.PushDataToAccessDatabase(args[1], dt);
            Console.WriteLine("Records inserted successfully");
        }
    }
}
