﻿using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace APA
{
    public class Periodes : List<Periode>
    {
        public int AktuellePeriode { get; private set; }

        public Periodes()
        {
            using (SqlConnection sqlConnection = new SqlConnection(Global.ConnectionStringUntis))
            {
                try
                {
                    string queryString = @"SELECT DISTINCT
Terms.TERM_ID, 
Terms.Name, 
Terms.Longname, 
Terms.DateFrom, 
Terms.DateTo
FROM Terms
WHERE (((Terms.SCHOOLYEAR_ID)= " + Global.AktSjUnt + ")  AND ((Terms.SCHOOL_ID)=177659)) ORDER BY Terms.TERM_ID;";

                    SqlCommand odbcCommand = new SqlCommand(queryString, sqlConnection);
                    sqlConnection.Open();
                    SqlDataReader sqlDataReader = odbcCommand.ExecuteReader();

                    Console.WriteLine("Perioden");
                    Console.WriteLine("--------");

                    while (sqlDataReader.Read())
                    {
                        Periode periode = new Periode()
                        {
                            IdUntis = sqlDataReader.GetInt32(0),
                            Name = Global.SafeGetString(sqlDataReader, 1),
                            Langname = Global.SafeGetString(sqlDataReader, 2),
                            Von = DateTime.ParseExact((sqlDataReader.GetInt32(3)).ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture),
                            Bis = DateTime.ParseExact((sqlDataReader.GetInt32(4)).ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture)
                        };

                        if (DateTime.Now > periode.Von && DateTime.Now < periode.Bis)
                            this.AktuellePeriode = periode.IdUntis;

                        Console.WriteLine(" " + periode.Name.ToString().PadRight(25) + " " + periode.Von.ToShortDateString() + " - " + periode.Bis.ToShortDateString());


                        this.Add(periode);
                    };

                    // Korrektur des Periodenendes

                    for (int i = 0; i < this.Count - 1; i++)
                    {
                        this[i].Bis = this[i + 1].Von.AddDays(-1);
                    }

                    sqlDataReader.Close();

                    Console.WriteLine("");

                    if (this.AktuellePeriode == 0)
                    {
                        Console.WriteLine("Es kann keine aktuelle Periode ermittelt werden. Das ist z. B. während der Sommerferien der Fall. Es wird die Periode " + this.Count + " als aktuelle Periode angenommen.");
                        this.AktuellePeriode = this.Count;
                    }
                    else
                    {
                        Console.WriteLine(" Aktuelle Periode: " + this.AktuellePeriode);
                    }

                    Console.WriteLine("");
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