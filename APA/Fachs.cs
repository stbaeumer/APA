using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace APA
{
    public class Fachs : List<Fach>
    {
        public Fachs()
        {
            using (SqlConnection sqlConnection = new SqlConnection(Global.ConnectionStringUntis))
            {
                try
                {
                    string queryString = @"SELECT DISTINCT 
Subjects.Subject_ID,
Subjects.Name,
Subjects.Longname,
Subjects.Text,
Description.Name
FROM Description RIGHT JOIN Subjects ON Description.DESCRIPTION_ID = Subjects.DESCRIPTION_ID
WHERE Subjects.Schoolyear_id = " + Global.AktSjUnt + " AND (Subjects.Deleted='false')  AND ((Subjects.SCHOOL_ID)=177659) ORDER BY Subjects.Name;";

                    SqlCommand odbcCommand = new SqlCommand(queryString, sqlConnection);
                    sqlConnection.Open();
                    SqlDataReader sqlDataReader = odbcCommand.ExecuteReader();

                    while (sqlDataReader.Read())
                    {
                        Fach fach = new Fach()
                        {
                            IdUntis = sqlDataReader.GetInt32(0),
                            KürzelUntis = Global.SafeGetString(sqlDataReader, 1)
                        };

                        this.Add(fach);
                    };

                    Console.WriteLine(("Fächer " + ".".PadRight(this.Count / 150, '.')).PadRight(48, '.') + (" " + this.Count).ToString().PadLeft(30), '.');

                    sqlDataReader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
                finally
                {
                    sqlConnection.Close();
                }
            }
        }
    }
}