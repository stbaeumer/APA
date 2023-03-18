using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace APA
{
    public class Leistungen : List<Leistung>
    {
        public Leistungen()
        {
        }

        public Leistungen(string connetionstringAtlantis, List<string> aktSj, string user, Lehrers lehrers)
        {
            try
            {
                var typ = (DateTime.Now.Month > 2 && DateTime.Now.Month <= 9) ? "JZ" : "HZ";

                Console.Write(("Leistungsdaten aus Atlantis (" + typ + ")").PadRight(71, '.'));

                using (OdbcConnection connection = new OdbcConnection(connetionstringAtlantis))
                {
                    DataSet dataSet = new DataSet();
                    OdbcDataAdapter schuelerAdapter = new OdbcDataAdapter(@"
SELECT DBA.noten_einzel.noe_id AS LeistungId,
DBA.noten_einzel.fa_id,
DBA.noten_einzel.s_art_fach,
DBA.noten_einzel.kurztext AS Fach,
DBA.noten_einzel.zeugnistext AS Zeugnistext,
DBA.noten_einzel.s_note AS Note,
DBA.noten_einzel.punkte AS Punkte,
DBA.noten_einzel.s_tendenz AS Tendenz,
DBA.noten_einzel.s_einheit AS Einheit,
DBA.noten_einzel.ls_id_1 AS LehrkraftAtlantisId,
DBA.schueler.name_1 AS Nachname,
DBA.schueler.name_2 AS Vorname,
DBA.schueler.pu_id AS SchlüsselExtern,
DBA.schue_sj.s_religions_unterricht AS Religion,
DBA.schue_sj.dat_austritt AS ausgetreten,
DBA.schue_sj.vorgang_akt_satz_jn AS SchuelerAktivInDieserKlasse,
DBA.schue_sj.vorgang_schuljahr AS Schuljahr,
(substr(schue_sj.s_berufs_nr,4,5)) AS Fachklasse,
DBA.klasse.s_klasse_art AS Anlage,
DBA.klasse.jahrgang AS Jahrgang,
DBA.schue_sj.s_gliederungsplan_kl AS Gliederung,
DBA.noten_kopf.s_typ_nok AS HzJz,
DBA.noten_kopf.nok_id AS NOK_ID,
DBA.noten_kopf.s_art_nok AS Zeugnisart,
DBA.noten_kopf.bemerkung_block_1 AS Bemerkung1,
DBA.noten_kopf.bemerkung_block_2 AS Bemerkung2,
DBA.noten_kopf.bemerkung_block_3 AS Bemerkung3,
DBA.noten_kopf.dat_notenkonferenz AS Konferenzdatum,
DBA.klasse.klasse AS Klasse
FROM(((DBA.noten_kopf JOIN DBA.schue_sj ON DBA.noten_kopf.pj_id = DBA.schue_sj.pj_id) JOIN DBA.klasse ON DBA.schue_sj.kl_id = DBA.klasse.kl_id) JOIN DBA.noten_einzel ON DBA.noten_kopf.nok_id = DBA.noten_einzel.nok_id ) JOIN DBA.schueler ON DBA.noten_einzel.pu_id = DBA.schueler.pu_id
WHERE schue_sj.s_typ_vorgang = 'A' AND (s_typ_nok = 'JZ' OR s_typ_nok = 'HZ') AND
(  
    (vorgang_schuljahr = '" + (Convert.ToInt32(aktSj[0])) + "/" + (Convert.ToInt32(aktSj[1])) + @"' AND klasse.jahrgang = 'C032'  AND s_note IS NOT NULL)
OR
  (vorgang_schuljahr = '" + (Convert.ToInt32(aktSj[0])) + "/" + (Convert.ToInt32(aktSj[1])) + @"' AND (klasse.jahrgang = 'B082') AND s_note IS NOT NULL)
OR
  (vorgang_schuljahr = '" + (Convert.ToInt32(aktSj[0])) + "/" + (Convert.ToInt32(aktSj[1])) + @"' AND (klasse.jahrgang = 'C081') AND s_note IS NOT NULL)
OR
  (vorgang_schuljahr = '" + (Convert.ToInt32(aktSj[0])) + "/" + (Convert.ToInt32(aktSj[1])) + @"' AND (klasse.jahrgang = 'C081') AND s_note IS NOT NULL)
OR
  (vorgang_schuljahr = '" + (Convert.ToInt32(aktSj[0])) + "/" + (Convert.ToInt32(aktSj[1])) + @"' AND (klasse.jahrgang = 'C061') AND s_note IS NOT NULL)
OR
  (vorgang_schuljahr = '" + (Convert.ToInt32(aktSj[0])) + "/" + (Convert.ToInt32(aktSj[1])) + @"' AND (klasse.jahrgang = 'C132') AND s_note IS NOT NULL)
)
ORDER BY DBA.noten_einzel.nok_id ASC , DBA.noten_einzel.position_1 ASC, DBA.schue_sj.vorgang_schuljahr DESC; ", connection);


                    connection.Open();
                    schuelerAdapter.Fill(dataSet, "DBA.leistungsdaten");

                    string bereich = "";

                    foreach (DataRow theRow in dataSet.Tables["DBA.leistungsdaten"].Rows)
                    {
                        // Leistungen des aktuellen Abschnitts sind abhängig vom aktuellen Monat.
                        // Leistungen vergangener Jahre sind immer "JZ"

                        if (theRow["s_art_fach"].ToString() == "U")
                        {
                            bereich = theRow["Zeugnistext"].ToString();
                        }
                        else
                        {
                            if (typ == theRow["HzJz"].ToString() || (theRow["Schuljahr"].ToString() != (Convert.ToInt32(aktSj[0]) - 0) + "/" + (Convert.ToInt32(aktSj[1]) - 0) && theRow["HzJz"].ToString() == "JZ"))
                            {
                                var sss = theRow["s_art_fach"].ToString();

                                DateTime austrittsdatum = theRow["ausgetreten"].ToString().Length < 3 ? new DateTime() : DateTime.ParseExact(theRow["ausgetreten"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);

                                Leistung leistung = new Leistung();

                                try
                                {
                                    // Wenn der Schüler nicht in diesem Schuljahr ausgetreten ist ...

                                    if (!(austrittsdatum > new DateTime(DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1, 8, 1) && austrittsdatum < DateTime.Now))
                                    {
                                        leistung.LeistungId = Convert.ToInt32(theRow["LeistungId"]);
                                        leistung.ReligionAbgewählt = theRow["Religion"].ToString() == "N";
                                        leistung.Schuljahr = theRow["Schuljahr"].ToString();
                                        leistung.Bereich = bereich;
                                        leistung.Gliederung = theRow["Gliederung"].ToString();
                                        leistung.HatBemerkung = (theRow["Bemerkung1"].ToString() + theRow["Bemerkung2"].ToString() + theRow["Bemerkung3"].ToString()).Contains("Fehlzeiten") ? true : false;
                                        leistung.Jahrgang = Convert.ToInt32(theRow["Jahrgang"].ToString().Substring(3, 1));
                                        leistung.Name = theRow["Nachname"] + " " + theRow["Vorname"];
                                        leistung.Klasse = theRow["Klasse"].ToString();
                                        leistung.Fach = theRow["Fach"] == null ? "" : theRow["Fach"].ToString();
                                        leistung.Gesamtnote = theRow["Note"].ToString() == "" ? null : theRow["Note"].ToString() == "Attest" ? "A" : theRow["Note"].ToString();
                                        leistung.Gesamtpunkte = theRow["Punkte"].ToString() == "" ? null : (theRow["Punkte"].ToString()).Split(',')[0];
                                        leistung.Tendenz = theRow["Tendenz"].ToString() == "" ? null : theRow["Tendenz"].ToString();
                                        leistung.EinheitNP = theRow["Einheit"].ToString() == "" ? "N" : theRow["Einheit"].ToString();
                                        leistung.SchlüsselExtern = Convert.ToInt32(theRow["SchlüsselExtern"].ToString());
                                        leistung.HzJz = theRow["HzJz"].ToString();
                                        leistung.Anlage = theRow["Anlage"].ToString();
                                        leistung.Zeugnisart = theRow["Zeugnisart"].ToString();
                                        leistung.Zeugnistext = theRow["Zeugnistext"].ToString();
                                        leistung.Konferenzdatum = theRow["Konferenzdatum"].ToString().Length < 3 ? new DateTime() : (DateTime.ParseExact(theRow["Konferenzdatum"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture)).AddHours(15);
                                        leistung.SchuelerAktivInDieserKlasse = theRow["SchuelerAktivInDieserKlasse"].ToString() == "J";
                                        leistung.Beschreibung = "";
                                        leistung.GeholteNote = false;
                                        if ((theRow["LehrkraftAtlantisId"]).ToString() != "")
                                        {
                                            leistung.LehrkraftAtlantisId = Convert.ToInt32(theRow["LehrkraftAtlantisId"]);
                                            leistung.LehrerKürzel = (from l in lehrers where l.AtlantisId == leistung.LehrkraftAtlantisId select l.Kürzel).FirstOrDefault();
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("Fehler beim Einlesen der Atlantis-Leistungsdatensätze: ENTER" + ex);
                                    Console.ReadKey();
                                }

                                // Nur Fächer mit Note werden hinzugefügt. Ein Fach wird kein zweites Mal hinzugefügt.

                                if (leistung.SchlüsselExtern == 152113 && leistung.Fach == "BWR")
                                {
                                    string aa = "";
                                }

                                if (leistung.Gesamtnote != "" && !(from x in this where x.SchlüsselExtern == leistung.SchlüsselExtern where x.Fach.Replace("B1", "").Replace("B2", "") == leistung.Fach.Replace("B1", "").Replace("B2", "") select x).Any())
                                {

                                    this.Add(leistung);
                                }
                            }
                        }
                    }
                    connection.Close();

                    foreach (var item in this)
                    {
                        if (item.SchlüsselExtern == 152113)
                        {
                            Console.WriteLine(item.SchlüsselExtern + " " + item.Fach + " " + item.Schuljahr + " " + item.Gesamtnote);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            Console.WriteLine((" " + this.Count.ToString()).PadLeft(30, '.'));            
        }

        //public Leistungen(Fachs fachs, Klasses klasses)
    //    {
    //        using (StreamReader reader = new StreamReader(Global.InputNotenCsv))
    //        {
    //            string überschrift = reader.ReadLine();

    //            int i = 1;

    //            Leistung leistung = new Leistung();

    //            while (true)
    //            {
    //                string line = reader.ReadLine();

    //                try
    //                {
    //                    if (line != null)
    //                    {
    //                        var x = line.Split('\t');
    //                        i++;

    //                        if (i == 2629)
    //                        {
    //                            string a = "";
    //                        }
    //                        if (x.Length == 10)
    //                        {
    //                            leistung = new Leistung();
    //                            leistung.Datum = DateTime.ParseExact(x[0], "dd.MM.yyyy", System.Globalization.CultureInfo.InvariantCulture);
    //                            leistung.Name = x[1];
    //                            leistung.Klasse = x[2];
    //                            leistung.Fach = (from f in fachs where f.KürzelUntis.ToString() == x[3] select f).FirstOrDefault();
    //                            leistung.Prüfungsart = x[4];
    //                            leistung.Gesamtpunkte = x[9];
    //                            leistung.Bemerkung = x[6];
    //                            leistung.Benutzer = x[7];
    //                            leistung.SchlüsselExtern = Convert.ToInt32(x[8]);
    //                        }

    //                        // Wenn in den Bemerkungen eine zusätzlicher Umbruch eingebaut wurde:

    //                        if (x.Length == 7)
    //                        {
    //                            leistung = new Leistung();
    //                            leistung.Datum = DateTime.ParseExact(x[0], "dd.MM.yyyy", System.Globalization.CultureInfo.InvariantCulture);
    //                            leistung.Name = x[1];
    //                            leistung.Klasse = x[2];
    //                            leistung.Fach = (from f in fachs where f.KürzelUntis.ToString() == x[3] select f).FirstOrDefault();
    //                            leistung.Bemerkung = x[6];
    //                            Console.WriteLine("\n\n  [!] Achtung: In den Zeilen " + (i - 1) + "-" + i + " hat vermutlich die Lehrkraft eine Bemerkung mit einem Zeilen-");
    //                            Console.Write("      umbruch eingebaut. Es wird nun versucht trotzdem korrekt zu importieren ... ");
    //                        }

    //                        if (x.Length == 4)
    //                        {
    //                            leistung.Benutzer = x[1];
    //                            leistung.SchlüsselExtern = Convert.ToInt32(x[2]);
    //                            leistung.Gesamtpunkte = x[3];
    //                            Console.WriteLine("hat geklappt.\n");
    //                        }

    //                        if (x.Length < 4)
    //                        {
    //                            Console.WriteLine("\n\n[!] MarksPerLesson.CSV: In der Zeile " + i + " stimmt die Anzahl der Spalten nicht. Das kann passieren, wenn z. B. die Lehrkraft bei einer Bemerkung einen Umbruch eingibt. Mit Suchen & Ersetzen kann die Datei MarksPerLesson.CSV korrigiert werden.");
    //                            Console.ReadKey();
    //                            throw new Exception("\n\n[!] MarksPerLesson.CSV: In der Zeile " + i + " stimmt die Anzahl der Spalten nicht. Das kann passieren, wenn z. B. die Lehrkraft bei einer Bemerkung einen Umbruch eingibt. Mit Suchen & Ersetzen kann die Datei MarksPerLesson.CSV korrigiert werden.");
    //                        }

    //                        // Nur Halbjahresnoten und Blaue Briefe sind relevant. Differenzierungsbereich zählt nicht.

    //                        if (Global.Mangelhaft.Contains(leistung.BlauerBriefNote) || Global.Ungenügend.Contains(leistung.BlauerBriefNote))
    //                        {

    //                            if (leistung.Prüfungsart == Global.BlaueBriefe)
    //                            {
    //                                if (leistung.Fach != null)
    //                                {
    //                                    if (leistung.IstKeinDiff(klasses))
    //                                    {
    //                                        this.Add(leistung);
    //                                    }
    //                                    else
    //                                    {
    //                                        Console.WriteLine("ACHTUNG: Mahnung im Diff-Bereich. " + leistung.Klasse + ": " + leistung.Fach.BezeichnungImZeugnis + " [ENTER]");
    //                                        Console.ReadKey();
    //                                    }
    //                                }
    //                                else
    //                                {
    //                                    Console.WriteLine("ACHTUNG: Blauer Brief ohne Fach bei Zeile: " + i);
    //                                }
    //                            }
    //                        }
    //                    }
    //                }
    //                catch (Exception ex)
    //                {
    //                    throw ex;
    //                }

    //                if (line == null)
    //                {
    //                    break;
    //                }
    //            }
    //            Console.WriteLine(("Leistungsdaten " + ".".PadRight(this.Count / 150, '.')).PadRight(48, '.') + (" " + this.Count).ToString().PadLeft(4), '.');
    //        }
    //    }
    }
}
