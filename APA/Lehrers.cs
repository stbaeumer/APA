using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;

namespace APA
{
    public class Lehrers : List<Lehrer>
    {
        private string connectionStringAtlantis;
        private List<string> aktSj;

        public Lehrers()
        {
        }

        public Lehrers(string connectionStringAtlantis, List<string> aktSj)
        {
            try
            {
                var typ = (DateTime.Now.Month > 2 && DateTime.Now.Month <= 9) ? "JZ" : "HZ";

                Console.Write(("Lehrer aus Atlantis (" + typ + ")").PadRight(71, '.'));

                using (OdbcConnection connection = new OdbcConnection(connectionStringAtlantis))
                {
                    DataSet dataSet = new DataSet();
                    OdbcDataAdapter schuelerAdapter = new OdbcDataAdapter(@"
SELECT DBA.lehr_sc.ls_id As Id,
DBA.lehrer.le_kuerzel As Kuerzel,
DBA.adresse.email As Mail
FROM(DBA.lehr_sc JOIN DBA.lehrer ON DBA.lehr_sc.le_id = DBA.lehrer.le_id) JOIN DBA.adresse ON DBA.lehrer.le_id = DBA.adresse.le_id
WHERE vorgang_schuljahr = '" + (Convert.ToInt32(aktSj[0]) - 0) + "/" + (Convert.ToInt32(aktSj[1]) - 0) + @"'; ", connection);
                    connection.Open();
                    schuelerAdapter.Fill(dataSet, "DBA.leistungsdaten");

                    foreach (DataRow theRow in dataSet.Tables["DBA.leistungsdaten"].Rows)
                    {
                        Lehrer lehrer = new Lehrer();
                        try
                        {
                            lehrer.Kürzel = theRow["Kuerzel"].ToString();
                            lehrer.AtlantisId = Convert.ToInt32(theRow["Id"]);
                            lehrer.Mail = theRow["Mail"].ToString();
                        }
                        catch (Exception ex)
                        {
                        }
                        if (lehrer.Mail.Contains("@berufskolleg-borken.de"))
                        {
                            this.Add(lehrer);
                        }
                    }
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            Console.WriteLine((" " + this.Count.ToString()).PadLeft(30, '.'));
        }

        public Lehrers(Periodes periodes, string ConnectionStringAtlantis, List<string> aktSj)
        {
            Lehrers alleAtlantisLehrer = new Lehrers(ConnectionStringAtlantis, aktSj);

            using (SqlConnection sqlConnection = new SqlConnection(Global.ConnectionStringUntis))
            {
                try
                {
                    string queryString = @"SELECT DISTINCT 
Teacher.Teacher_ID, 
Teacher.Name, 
Teacher.Longname, 
Teacher.FirstName,
Teacher.Email,
Teacher.Flags,
Teacher.Title,
Teacher.ROOM_ID,
Teacher.Text2,
Teacher.Text3,
Teacher.PlannedWeek
FROM Teacher 
WHERE (((SCHOOLYEAR_ID)= " + Global.AktSjUnt + ") AND  ((TERM_ID)=" + periodes.Count + ") AND ((Teacher.SCHOOL_ID)=177659) AND (((Teacher.Deleted)='false'))) ORDER BY Teacher.Name;";

                    SqlCommand odbcCommand = new SqlCommand(queryString, sqlConnection);
                    sqlConnection.Open();
                    SqlDataReader sqlDataReader = odbcCommand.ExecuteReader();

                    while (sqlDataReader.Read())
                    {
                        Lehrer lehrer = new Lehrer()
                        {
                            IdUntis = sqlDataReader.GetInt32(0),
                            Kürzel = Global.SafeGetString(sqlDataReader, 1),
                            Nachname = Global.SafeGetString(sqlDataReader, 2),
                            Vorname = Global.SafeGetString(sqlDataReader, 3),
                            Mail = Global.SafeGetString(sqlDataReader, 4),
                            Anrede = Global.SafeGetString(sqlDataReader, 5) == "n" ? "Herr" : Global.SafeGetString(sqlDataReader, 5) == "W" ? "Frau" : "",
                            Titel = Global.SafeGetString(sqlDataReader, 6),
                            Funktion = Global.SafeGetString(sqlDataReader, 8),
                            Dienstgrad = Global.SafeGetString(sqlDataReader, 9)
                        };

                        if (!lehrer.Mail.EndsWith("@berufskolleg-borken.de") && lehrer.Kürzel != "LAT" && lehrer.Kürzel != "?")
                            Console.WriteLine("Untis2Exchange Fehlermeldung: Der Lehrer " + lehrer.Kürzel + " hat keine Mail-Adresse in Untis. Bitte in Untis eintragen.");
                        if (lehrer.Anrede == "" && lehrer.Kürzel != "LAT" && lehrer.Kürzel != "?")
                            Console.WriteLine("Untis2Exchange Fehlermeldung: Der Lehrer " + lehrer.Kürzel + " hat kein Geschlecht in Untis. Bitte in Untis eintragen.");

                        lehrer.AtlantisId = (from l in alleAtlantisLehrer where l.Kürzel == lehrer.Kürzel select l.AtlantisId).FirstOrDefault();

                        this.Add(lehrer);
                    };

                    Console.WriteLine(("Lehrer " + ".".PadRight(this.Count / 150, '.')).PadRight(48, '.') + (" " + this.Count).ToString().PadLeft(30), '.');

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
        
        internal void FehlendeUndDoppelteEinträge(Schuelers schuelers)
        {
            foreach (var lehrer in this)
            {
                if (lehrer.Mail != null && lehrer.Mail != "")
                {
                    var schuelerOhneNoten = (from s in schuelers
                                             from f in s.Fächer
                                             where f.Note == null || f.Note == ""
                                             where f.Lehrerkürzel == lehrer.Kürzel
                                             select s).ToList();

                    List<Schueler> schuelerMitDoppelterNote = new List<Schueler>();
                    
                    if (schuelerOhneNoten.Count > 0)
                    {
                        lehrer.Mailen(schuelerOhneNoten, schuelerMitDoppelterNote);
                    }
                }                
            }
        }
    }
}