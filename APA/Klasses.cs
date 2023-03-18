using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;

namespace APA
{
    public class Klasses : List<Klasse>
    {
        public Lehrers Lehrers { get; set; }

        public Klasses(Lehrers lehrers, Periodes periodes, List<string> interessierendeKlassen)
        {
            Lehrers = lehrers;

            using (SqlConnection sqlConnection = new SqlConnection(Global.ConnectionStringUntis))
            {
                try
                {
                    string queryString = @"SELECT 
Class.CLASS_ID, 
Class.Name, 
Class.TeacherIds, 
Class.Longname, 
Teacher.Name,
Class.ClassLevel, 
Class.PERIODS_TABLE_ID,
Department.Name,
Class.TimeRequest,
Class.ROOM_ID,
Class.Text
FROM (Class LEFT JOIN Department ON Class.DEPARTMENT_ID = Department.DEPARTMENT_ID) LEFT JOIN Teacher ON Class.TEACHER_ID = Teacher.TEACHER_ID
WHERE (((Class.SCHOOL_ID)=177659) AND ((Class.TERM_ID)=" + periodes.Count + ") AND ((Class.Deleted)='false') AND ((Class.TERM_ID)=" + periodes.Count + ") AND ((Class.SCHOOLYEAR_ID)=" + Global.AktSjUnt + ") AND ((Department.SCHOOL_ID)=177659) AND ((Department.SCHOOLYEAR_ID)=" + Global.AktSjUnt + ") AND ((Teacher.SCHOOL_ID)=177659) AND ((Teacher.SCHOOLYEAR_ID)=" + Global.AktSjUnt + ") AND ((Teacher.TERM_ID)=" + periodes.Count + "))ORDER BY Class.Name ASC; ";

                    SqlCommand odbcCommand = new SqlCommand(queryString, sqlConnection);
                    sqlConnection.Open();
                    SqlDataReader sqlDataReader = odbcCommand.ExecuteReader();

                    while (sqlDataReader.Read())
                    {
                        List<Lehrer> klassenleitungen = new List<Lehrer>();

                        foreach (var item in (Global.SafeGetString(sqlDataReader, 2)).Split(','))
                        {
                            klassenleitungen.Add((from l in lehrers
                                                  where l.IdUntis.ToString() == item
                                                  select l).FirstOrDefault());
                        }

                        var klasseName = Global.SafeGetString(sqlDataReader, 1);

                        Klasse klasse = new Klasse()
                        {
                            IdUntis = sqlDataReader.GetInt32(0),
                            NameUntis = klasseName,
                            Klassenleitungen = klassenleitungen,
                            Jahrgang = Global.SafeGetString(sqlDataReader, 5),
                            Bereichsleitung = Global.SafeGetString(sqlDataReader, 7),
                            Beschreibung = Global.SafeGetString(sqlDataReader, 3),
                            Url = "https://www.berufskolleg-borken.de/bildungsgange/" + Global.SafeGetString(sqlDataReader, 10)
                        };

                        if (interessierendeKlassen.Contains(klasse.NameUntis)) 
                        {
                            this.Add(klasse);
                        }                       
                    };

                    Console.WriteLine(("Klassen " + ".".PadRight(this.Count / 150, '.')).PadRight(48, '.') + (" " + this.Count).ToString().PadLeft(30), '.');

                    sqlDataReader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                    throw new Exception(ex.ToString());
                }
                finally
                {
                    sqlConnection.Close();
                }
            }
        }

        internal List<string> Dokumente()
        {
            var x = new List<string>();

            foreach (var item in (from k in this select Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + k.NameUntis + ".pdf"))
            {
                x.Add(item);
            }

            x.Add(Global.Ziel);

            return x;
        }

        public Klasses()
        {
        }

        public Excelzeilen Notenlisten(Schuelers schuelers, Lehrers lehrers)
        {
            Excelzeilen excelzeilen = new Excelzeilen();

            string quelle = "APA.xlsx";
            
            System.IO.File.Copy(quelle, Global.Ziel, true);

            Application application = new Application();
            Workbook workbook = application.Workbooks.Open(Global.Ziel);
            try
            {
                foreach (var klasse in this)
                {
                    excelzeilen.Add(klasse.Notenliste(application, workbook, (from s in schuelers
                                                                              where s.Klasse != null
                                                                              where s.Klasse.NameUntis == klasse.NameUntis
                                                              select s).ToList(), lehrers));
                }
                return excelzeilen;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);                
            }
            finally
            {
                workbook.Save();
                workbook.Close();
                application.Quit();                
            }
            return null;
        }
    }
}