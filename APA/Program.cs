using ICSharpCode.SharpZipLib.Core;
using ICSharpCode.SharpZipLib.Zip;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;

namespace APA
{
    class Program
    {
        [STAThread]

        static void Main(string[] args)
        {
            System.Net.ServicePointManager.ServerCertificateValidationCallback = ((sender, certificate, chain, sslPolicyErrors) => true);

            string ConnectionStringAtlantis = @"Dsn=Atlantis9;uid=";
            string User = System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToUpper().Split('\\')[1];
            string Zeitstempel = DateTime.Now.ToString("yyMMdd-HHmmss");
            List<string> AktSj = new List<string>
            {
                (DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1).ToString(),
                (DateTime.Now.Month >= 8 ? DateTime.Now.Year + 1 - 2000 : DateTime.Now.Year - 2000).ToString()
            };

            try
            {
                var prds = new Periodes();
                var fchs = new Fachs();
                var lehs = new Lehrers(prds, ConnectionStringAtlantis, AktSj);

                Console.WriteLine("Bitte die interessierenden Klassen kommasepariert angeben [" + Properties.Settings.Default.InteressierendeKlassen + "] :");
                List<string> interessierendeKlassen = new List<string>();
                var x = Console.ReadLine();

                if (x == "")
                {
                    interessierendeKlassen.AddRange(Properties.Settings.Default.InteressierendeKlassen.Split(','));
                    x = Properties.Settings.Default.InteressierendeKlassen;
                }
                else
                {
                    interessierendeKlassen.AddRange(x.Split(','));
                    Properties.Settings.Default.InteressierendeKlassen = x;
                    Properties.Settings.Default.Save();
                }

                var klss = new Klasses(lehs, prds, interessierendeKlassen);

                var schuelers = new Schuelers(klss, lehs, interessierendeKlassen);
                Excelzeilen excelzeilen = new Excelzeilen();
                schuelers.AddLeistungen(ConnectionStringAtlantis + "dba", AktSj, User, lehs);

                excelzeilen.AddRange(klss.Notenlisten(schuelers, lehs));
                //excelzeilen.ToExchange(lehs);
                //lehs.FehlendeUndDoppelteEinträge(schuelers);                
                System.Diagnostics.Process.Start(Global.Ziel);
                //Global.MailSenden(new Klasse(), new Lehrer(), "Liste alle Dokumente für den APA", "Siehe Anlage.", klss.Dokumente());
                //System.Windows.Forms.Clipboard.SetText(excelzeilen.ToClipboard());
                Console.WriteLine("Tabelle ZulassungskonferenzBC in Zwischenablage geschrieben.");
                Console.ReadKey();
            }
            catch (IOException ex)
            {
                Console.WriteLine(ex.ToString());
                Console.ReadKey();
                Environment.Exit(0);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Heiliger Bimbam! Es ist etwas schiefgelaufen! Die Verarbeitung wird gestoppt.");
                Console.WriteLine("");
                Console.WriteLine(ex);
                Console.ReadKey();
                Environment.Exit(0);
            }
        }
    }
}
