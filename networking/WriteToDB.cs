using SpreadsheetReader.models;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;

namespace SpreadsheetReader.networking
{
    public class WriteToDB
    {
        public string sql { set; get; }
        public string connectionString { set; get; }
        public WriteToDB(string sql, string connectionString)
        {
            this.sql = sql;
            this.connectionString = connectionString;
        }

        public void Write()
        {
            string returnValue = "";
            using (SqlConnection connection = new SqlConnection(this.connectionString))
            {
                try
                {
                    connection.Open();

                    using (SqlCommand cmd = new SqlCommand(sql, connection))
                    {
                        int rowsAdded = cmd.ExecuteNonQuery();

                        if (rowsAdded > 0)
                        {
                            returnValue += (rowsAdded + " rows have been added to the database...");
                            new NetworkLogger(returnValue, false);
                        }

                        else
                        {
                            returnValue += (rowsAdded + " rows have been added to the database...");
                            new NetworkLogger(returnValue, false);
                        }
                            

                    }
                }
                catch (Exception ex)
                {
                    returnValue += ("NETWORKING ERROR: " + ex.Message + "\n Error exception - LINE 43, WriteToDB");
                    new NetworkLogger(returnValue, true);
                }
                
            }

            
        }
    }


}
