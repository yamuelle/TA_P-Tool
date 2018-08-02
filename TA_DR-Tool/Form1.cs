using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Collections;
using excel = Microsoft.Office.Interop.Excel;
using System.Net;

namespace TA_DR_Tool
{
    public partial class Form1 : Form
    {
        String messageText = "Sehr geehrte/r {0} \n\nich möchte hiermit folgende Dienstreise anmelden :\n\n\t- Was ist der Anlass der Reise?\n\t  {1}\n\n\t- Wo und wann (Anfang- und Endzeiten) findet die Veranstaltung statt?\n\t  von {2} bis {3} in {4} \n\n\t- Welche Reisezeiten fallen an?\n\t (Reisezeiten können, bei entsprechenden Regelungen im Betrieb als Überstunden angerechnet werden)\n\t  Anreise : {5} \n\t  Rückreise : {6}\n\n\t- Wie viele Überstunden fallen an und wann werden diese abgebaut?\n\t  {7} Stunden werden abgebaut durch {8}\n\n\t- Was ist das Reisemittel, bzw. fahren Sie alleine oder mit Kollegen?\n\t  {9}, {10}\n\n Beste Grüße";
        String procText = "";
        Microsoft.Office.Interop.Outlook.Application app;
        Microsoft.Office.Interop.Outlook.MailItem mi;
        String conffp = "";
        
        string[] lines;

        ArrayList mails = new ArrayList();
        ArrayList names = new ArrayList();

        public Form1()
        {
            
            InitializeComponent();
            if (loadconfig() != 1)
            {
                MessageBox.Show("Ein Fehler ist aufgetreten. Bitte starten Sie das Programm neu und wählen Sie eine richtige Konfigurationsdatei aus.");
                File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\tadrtsbconf.csv");
                Environment.Exit(0);
            }
            app = new Microsoft.Office.Interop.Outlook.Application();
            mi = app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
            

        }

        private int loadconfig()
        {
            if (File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\tadrtsbconf.csv"))
            {
                string[] lines = File.ReadAllLines(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\tadrtsbconf.csv");

                string[] linesproc = new string[lines.Length - 1];

                for (int i = 0; i < lines.Length - 1; i++)
                {
                    linesproc[i] = lines[i];
                }
                foreach (string s in linesproc)
                {
                    string[] tmp = s.Split(';');
                    names.Add(tmp[0]);
                    mails.Add(tmp[1]);
                }
                comboBox1.Items.AddRange(names.ToArray());
               
            }
            else
            {
                MessageBox.Show("Es konnte keine Konfigurationsdatei gefunden werden. Bitte wählen sie im anschließenden Dialog eine Konfigurationsdatei aus.");
                OpenFileDialog of = new OpenFileDialog();
                if (of.ShowDialog() == DialogResult.OK)
                {
                    conffp = of.FileName;
                }
                File.Copy(conffp, Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\tadrtsbconf.csv");
                string[] lines = File.ReadAllLines(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\tadrtsbconf.csv");

                if (lines[lines.Length-1] != "tadrconfigfile;tadrconfigfile")
                {
                    Console.WriteLine(lines[lines.Length-1]);
                    return -1;
                }

                string[] linesproc = new string[lines.Length-1];

                for(int i = 0; i < lines.Length-1; i++)
                {
                    linesproc[i] = lines[i];
                }

                foreach (string s in linesproc)
                {
                    string[] tmp = s.Split(';');
                    names.Add(tmp[0]);
                    mails.Add(tmp[1]);
                }
                comboBox1.Items.AddRange(names.ToArray());
            }
            if (File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\avconf.tadrt"))
            {
                string [] avmail = File.ReadAllLines(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\avconf.tadrt");
                avMail.Text = avmail[0];
            }else
            {
                avMail.Show();
                MessageBox.Show("Es wurde noch keine AV-Mail eingetragen. Bitte tragen Sie die Mail im Reiter Datenbestand bearbeiten nach.");

            }
            return 1;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (avMail.Text != "" && sbMail.Text != "")
            {


                procText = String.Format(messageText, comboBox1.Text, adr.Text, tb.Text + " " + tp1.Text, te.Text + " " + tp2.Text, ort.Text, rb.Text + " " + tp3.Text, re.Text + " " + tp4.Text, ueberstunden.Text, ad.Text, rm.Text, cb1.Text);
                mi.To = sbMail.Text;
                mi.Subject = "Anmeldung einer Dienstreise";
                mi.CC = avMail.Text;

                mi.Body = procText;
                if (checkBox1.Checked == true) mi.Send();
                else mi.Display();
                

                MessageBox.Show("E-Mail erfolgreich versendet");
                if(checkBox2.Checked==true)toexcel();
                MessageBox.Show("Die Anwendung wird nun beendet.");
                System.Windows.Forms.Application.Exit();
            }
            else MessageBox.Show("Bitte alle Mail-Felder Ausfüllen (auch AV-Mail im Reiter Datenbestand bearbeiten)");
        }
        private void toexcel()
        {
            MessageBox.Show("Bitte das persönliche Template des Reisekostenantrages auswählen");
            string pathtotemplate = "";
            string pathtowork = "";
            OpenFileDialog of = new OpenFileDialog();
            of.Filter = "XLSM Dateien |*.xlsm";
            if (of.ShowDialog() == DialogResult.OK)
            {
                pathtotemplate = of.FileName;
                MessageBox.Show("Wo soll der ausgefüllte Reisekostenantrag abgespeichert werden ?");
                SaveFileDialog sd = new SaveFileDialog();
                sd.Filter = "XLSM Dateien |*.xlsm";
                if (sd.ShowDialog() == DialogResult.OK)
                {
                    pathtowork = sd.FileName;
                    File.Copy(pathtotemplate, pathtowork);
                }




            }
            Microsoft.Office.Interop.Excel.Application excelapp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wb = excelapp.Workbooks.Open(pathtowork);
            Microsoft.Office.Interop.Excel.Worksheet ws = wb.Sheets[1];
            excel.Range rbd = ws.Cells[24, 5] as excel.Range;
            rbd.Value2 = rb.Value.ToString("dd.MM.yyyy");
            excel.Range rbz = ws.Cells[24, 11] as excel.Range;
            rbz.Value2 = tp3.Text;
            excel.Range red = ws.Cells[24, 41] as excel.Range;
            red.Value2 = re.Value.ToString("dd.MM.yyyy");
            excel.Range rez = ws.Cells[24, 47] as excel.Range;
            rez.Value2 = tp4.Text;
            excel.Range tbd = ws.Cells[24, 17] as excel.Range;
            tbd.Value2 = tb.Value.ToString("dd.MM.yyyy");
            excel.Range ted = ws.Cells[24, 29] as excel.Range;
            ted.Value2 = te.Value.ToString("dd.MM.yyyy");
            excel.Range taz = ws.Cells[24, 23] as excel.Range;
            taz.Value2 = tp1.Text;
            excel.Range tez = ws.Cells[24, 35] as excel.Range;
            tez.Value2 = tp2.Text;
            excel.Range anlass = ws.Cells[27, 5] as excel.Range;
            anlass.Value2 = adr.Text;
            excel.Range o = ws.Cells[31, 5] as excel.Range;
            o.Value2 = ort.Text;
            wb.Save();
            excelapp.Quit();
            MessageBox.Show("Reisekostenantrag unter der angegebenen Adresse gespeichert. Programm wird beendet");
            System.Windows.Forms.Application.Exit();

        }










        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void re_ValueChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            // MessageBox.Show((string)mails[comboBox1.SelectedIndex]);
            sbMail.Text = (string)mails[comboBox1.SelectedIndex];
        }

        private void comboBox1_TextUpdate(object sender, EventArgs e)
        {
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            OpenFileDialog of = new OpenFileDialog();
            if (of.ShowDialog() == DialogResult.OK)
            {
                conffp = of.FileName;
                File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\tadrtsbconf.csv");
                File.Copy(conffp, Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\tadrtsbconf.csv");
                string[] lines = File.ReadAllLines(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\tadrtsbconf.csv");
            }
            

            System.Windows.Forms.Application.Restart();           
        }

        private void button3_Click(object sender, EventArgs e)
        {
            String[] m = {avMail.Text};
            if (File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\avconf.tadrt")){
                File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\avconf.tadrt"); }
            File.Create(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\avconf.tadrt").Close();
            File.WriteAllLines(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\avconf.tadrt", m);
            MessageBox.Show("AV-Mail erfolgreich geändert");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            toexcel();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://yam.telekom.de/people/Yannik_M%C3%BCller");
        }
    }
}
