using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System;
using System.Data;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using System.Configuration;
using System.Data.OleDb;

namespace program
{
    public partial class Form2 : Form
    {
        public string MainConnectionstring = ConfigurationManager.ConnectionStrings["program.Properties.Settings.hunter8ConnectionString"].ConnectionString;
        public DataTable kitolt(ComboBox comb)
        {
            DataTable D = new DataTable();

            string cmdtext = "SELECT nev FROM megbizott ";

            D = adatleker2(cmdtext);
            foreach (DataRow d in D.Rows)
            {
                try
                {
                    comb.Items.Add(d[0].ToString());


                }
                catch (Exception)
                {
                }

            }
            return D;
        }
        public Form2(string str, DataGridView grd, string lblstr)
        {
            InitializeComponent();
            textBox19.Text = str;
            comboBox1.Items.Clear();
            kitolt(comboBox1);
            label5.Text = lblstr;
            //Width = 915;


            DataTable d = new DataTable();
            d.Columns.AddRange(new DataColumn[8] {

                new DataColumn("Terület fekvése"),
                new DataColumn("Külvagybel"),
                new DataColumn("Hszám"),
                new DataColumn("Művelési ág"),
                new DataColumn("Öszz Terület"),
                new DataColumn("Tulajdoni hányad"),
                new DataColumn("Terület HA-ban"),
                new DataColumn("Rendelkezési jogcím")

            });

            try
            {
                foreach (DataGridViewRow row in grd.Rows)
                {
                    d.Rows.Add(row.Cells[0].Value.ToString(), row.Cells[1].Value.ToString(), row.Cells[2].Value.ToString(), row.Cells[3].Value.ToString(), row.Cells[4].Value.ToString(), row.Cells[5].Value.ToString(), row.Cells[6].Value.ToString(), row.Cells[7].Value.ToString());

                }

            }
            catch (Exception)
            {
                MessageBox.Show("Error");

            }
            dataGridView1.DataSource = d;





        }

        public DataTable adatleker2(string cmdstr)
        {

            DataTable Dtable = new DataTable();


            string connectionstr = MainConnectionstring;
                //@"Data Source = (LocalDB)\MSSQLLocalDB; AttachDbFilename = 'c:\users\ge70840\documents\visual studio 2015\Projects\Hunterfull8\program\hunter8.mdf'; Integrated Security = True";
            string commandstr = cmdstr;
            try
            {
                using (SqlConnection connect = new SqlConnection(connectionstr))
                {
                    using (SqlCommand command = new SqlCommand(commandstr, connect))
                    {
                        connect.Open();
                        SqlDataReader reader = command.ExecuteReader();

                        Dtable.Load(reader);
                        connect.Close();

                    }

                }
            }
            catch (Exception ex)
            {
                string lines = ex.ToString();
                //string path = @"C:\Users\ge70840\Documents\visual studio 2015\Projects\Hunter\Hunter\bin\Debug\error.log";

                //using (StreamWriter sw = File.AppendText(path))
                //{
                //    sw.WriteLine(DateTime.Now.ToString());
                //    sw.WriteLine(lines);
                //    sw.WriteLine();
                //}
                MessageBox.Show(ex.ToString());
            }


            return Dtable;

        }
        public DataTable kitolt3(string nev)
        {
            DataTable D = new DataTable();

            string cmdtext = "SELECT nev, cim, varos, szulhely, szulido, anyjaneve, szulnev, igazolvanytip, igazolvanyszam,id FROM megbizott WHERE(nev LIKE '" + nev + "' + '%')";
            D = adatleker2(cmdtext);
            foreach (DataRow d in D.Rows)
            {
                try
                {
                    textBox1.Text = d[0].ToString();
                    textBox2.Text = d[1].ToString();
                    textBox3.Text = d[2].ToString();
                    textBox4.Text = d[3].ToString();
                    textBox5.Text = d[4].ToString();
                    textBox6.Text = d[5].ToString();
                    textBox7.Text = d[6].ToString();
                    textBox8.Text = d[7].ToString();
                    textBox9.Text = d[8].ToString();
                    label7.Text = d[9].ToString();



                }
                catch (Exception)
                {
                }

            }
            return D;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            kitolt3(label6.Text);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            label6.Text = comboBox1.SelectedItem.ToString();
        }
        public void adatrogz(string nev, string cim, string varos, string szulhely, string szulido, string anyjaneve, string szulneve, string igazolvanytip, string igazolvanyszam)
        {


            string connectionstr = MainConnectionstring;
                //@"Data Source = (LocalDB)\MSSQLLocalDB; AttachDbFilename = 'c:\users\ge70840\documents\visual studio 2015\Projects\Hunterfull8\program\hunter8.mdf'; Integrated Security = True";

            SqlConnection con = new SqlConnection(connectionstr);
            try
            {
                string commandtext = "INSERT INTO megbizott (nev, cim, varos, szulhely, szulido, anyjaneve, szulnev, igazolvanytip, igazolvanyszam) VALUES('" + nev + "', '" + cim + "', '" + varos + "', '" + szulhely + "', '" + szulido + "', '" + anyjaneve + "', '" + szulneve + "', '" + igazolvanytip + "', '" + igazolvanyszam + "')";
                SqlCommand comm = new SqlCommand(commandtext, con);
                con.Open();
                comm.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                con.Close();
            }











        }

        private void button1_Click(object sender, EventArgs e)
        {
            adatrogz(textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text, textBox5.Text, textBox6.Text, textBox7.Text, textBox8.Text, textBox9.Text);
            MessageBox.Show(DialogResult.OK.ToString());
        }

        //private void button7_Click(object sender, EventArgs e)
        //{
        //    comboBox1.Items.Clear();
        //    string cmdtext = "SELECT [nev] FROM megbizott ";
        //    DataTable DD = new DataTable();
        //    DD = adatleker2(cmdtext);
        //    foreach (DataRow dd in DD.Rows)
        //    {
        //        if (dd[0].ToString() != String.Empty)
        //        {
        //            comboBox1.Items.Add(dd[0].ToString());
        //        }


        //    }
        //}
        public void torol(string id)
        {
            string ConnetionString = MainConnectionstring;
                //@"Data Source = (LocalDB)\MSSQLLocalDB; AttachDbFilename = 'c:\users\ge70840\documents\visual studio 2015\Projects\Hunterfull8\program\hunter8.mdf'; Integrated Security = True";
            try
            {

                using (SqlConnection con = new SqlConnection(ConnetionString))
                {
                    string commandtext = "DELETE FROM megbizott WHERE ID = '" + id + "'";


                    SqlCommand comm = new SqlCommand(commandtext, con);
                    con.Open();
                    comm.ExecuteNonQuery();

                }
            }
            catch (Exception)
            {

                MessageBox.Show("Error");
            }


        }
        private void button5_Click(object sender, EventArgs e)
        {
            torol(label7.Text);
            MessageBox.Show("Ok");
        }
        public void modosit(string id, string nev, string cim, string varos, string szulhely, string szulido, string anyaneve, string szulnev, string igazolvanytipus, string igazolvanyszam)
        {
            string ConnetionString = MainConnectionstring;
                //@"Data Source = (LocalDB)\MSSQLLocalDB; AttachDbFilename = 'c:\users\ge70840\documents\visual studio 2015\Projects\Hunterfull8\program\hunter8.mdf'; Integrated Security = True";
            
            try
            {
                using (SqlConnection con = new SqlConnection(ConnetionString))
                {
                    string commandtext = "UPDATE [dbo].[megbizott] SET [nev] = '" + nev + "', [cim] = '" + cim + "',[varos] = '" + varos + "',[szulhely] = '" + szulhely + "',[szulido] = '" + szulido + "',[anyjaneve] = '" + anyaneve + "',[szulnev] = '" + szulnev + "',[igazolvanytip] = '" + igazolvanytipus + "', [igazolvanyszam] = '" + igazolvanyszam + "' WHERE Id = " + id + "";
                    SqlCommand comm = new SqlCommand(commandtext, con);
                    con.Open();
                    comm.ExecuteNonQuery();

                }
            }
            catch (Exception)
            {

                MessageBox.Show("Error");
            }


        }

        private void button6_Click(object sender, EventArgs e)
        {
            modosit(label7.Text, textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text, textBox5.Text, textBox6.Text, textBox7.Text, textBox8.Text, textBox9.Text);
            MessageBox.Show("OK");
        }
        public void kitolt2()
        {
            DataTable D = new DataTable();
            string sss = textBox19.Text;
            int szulindex = sss.IndexOf("szül");
            int szulhossz = sss.IndexOf("szül") + "szül".Length;
            int anindex = sss.IndexOf("an:");
            int cimindex = sss.IndexOf("cím");
            int perindex = sss.IndexOf('/');
            string szulnev = null;
            string nev = null;
            string anyjanev = null;
            string cim = null;
            string[] varos = new string[3];

            try
            {
                if (sss.StartsWith("MAGYAR"))
                {
                    nev = sss.Substring(0, cimindex);
                }
                else
                {
                    nev = sss.Substring(0, anindex);
                }

                try
                {
                    if (szulindex > 0)
                    {
                        szulnev = sss.Substring(szulhossz + 1, (anindex) - (szulhossz + 1));
                    }
                    else
                    {
                        szulnev = "U.A.";
                    }

                }
                catch (Exception)
                {
                    MessageBox.Show("szulindexerror");

                }
                if (anindex > 0)
                {
                    anyjanev = sss.Substring(anindex + "an".Length + 1, (cimindex) - (anindex + "an".Length + 1));
                }
                else
                {
                    anyjanev = "Nincs Adat";
                }

                if (sss.StartsWith("MAGYAR"))
                {
                    cim = "Nincs Adat";
                }
                else
                {
                    cim = sss.Substring(cimindex + "cím".Length + 2, (perindex) - (cimindex + "cím".Length + 2));
                    varos = cim.Split(' ');
                    textBox11.Text = varos[2] + ' ' + varos[3] + ' ' + varos[4] + ' ' + varos[5];
                    textBox12.Text = varos[1];

                }
                textBox10.Text = nev;
                textBox15.Text = anyjanev;
                textBox16.Text = szulnev;

            }
            catch (Exception)
            {
                MessageBox.Show("nemntommierror");

            }




        }

        private void button2_Click(object sender, EventArgs e)
        {
            kitolt2();
        }
        private void CreateWordDocument(object fileName,
           object saveAs)
        {
            //Set Missing Value parameter - used to represent
            // a missing value when calling methods through
            // interop.
            object missing = System.Reflection.Missing.Value;

            //Setup the Word.Application class.
            Word.Application wordApp =
                new Word.ApplicationClass();

            //Setup our Word.Document class we'll use.
            Word.Document aDoc = null;

            // Check to see that file exists
            if (File.Exists((string)fileName))
            {
                DateTime today = DateTime.Now;

                object readOnly = false;
                object isVisible = false;

                //Set Word to be not visible.
                wordApp.Visible = false;

                //Open the word document
                aDoc = wordApp.Documents.Open(ref fileName, ref missing,
                    ref readOnly, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref isVisible, ref missing, ref missing,
                    ref missing, ref missing);

                // Activate the document
                aDoc.Activate();

                // Find Place Holders and Replace them with Values.
                //alulirott
                this.FindAndReplace(wordApp, "<name>", textBox10.Text);
                this.FindAndReplace(wordApp, "<address>", textBox11.Text);
                this.FindAndReplace(wordApp, "<city>", textBox12.Text);
                this.FindAndReplace(wordApp, "<bornplace>", textBox13.Text);
                this.FindAndReplace(wordApp, "<borndate>", textBox14.Text);
                this.FindAndReplace(wordApp, "<mothername>", textBox15.Text);
                this.FindAndReplace(wordApp, "<birthname>", textBox16.Text);
                this.FindAndReplace(wordApp, "<cardtype>", textBox17.Text);
                this.FindAndReplace(wordApp, "<cardnumber>", textBox18.Text);
                //meghatalmazom
                this.FindAndReplace(wordApp, "<name2>", textBox1.Text);
                this.FindAndReplace(wordApp, "<address2>", textBox2.Text);
                this.FindAndReplace(wordApp, "<city2>", textBox3.Text);
                this.FindAndReplace(wordApp, "<bornplace2>", textBox4.Text);
                this.FindAndReplace(wordApp, "<borndate2>", textBox5.Text);
                this.FindAndReplace(wordApp, "<mothername2>", textBox6.Text);
                this.FindAndReplace(wordApp, "<birthname2>", textBox7.Text);
                this.FindAndReplace(wordApp, "<cardtype2>", textBox8.Text);
                this.FindAndReplace(wordApp, "<cardnumber2>", textBox9.Text);
                //tablazat


                for (int i = 0; i < 9; i++)
                {
                    switch (i)
                    {
                        case 1: this.FindAndReplace(wordApp, "<col" + i + ">", dataGridView1.Rows[0].Cells[0].Value.ToString()); break;
                        case 2: this.FindAndReplace(wordApp, "<col" + i + ">", dataGridView1.Rows[0].Cells[1].Value.ToString()); break;
                        case 3: this.FindAndReplace(wordApp, "<col" + i + ">", dataGridView1.Rows[0].Cells[2].Value.ToString()); break;
                        case 4: this.FindAndReplace(wordApp, "<col" + i + ">", dataGridView1.Rows[0].Cells[3].Value.ToString()); break;
                        case 5: this.FindAndReplace(wordApp, "<col" + i + ">", dataGridView1.Rows[0].Cells[4].Value.ToString()); break;
                        case 6: this.FindAndReplace(wordApp, "<col" + i + ">", dataGridView1.Rows[0].Cells[5].Value.ToString()); break;
                        case 7: this.FindAndReplace(wordApp, "<col" + i + ">", dataGridView1.Rows[0].Cells[6].Value.ToString()); break;
                        case 8: this.FindAndReplace(wordApp, "<col" + i + ">", dataGridView1.Rows[0].Cells[7].Value.ToString()); break;

                        default:

                            break;
                    }


                }


                for (int k = 1; k < dataGridView1.RowCount; k++)
                {
                    for (int i = 1; i < 9; i++)
                    {
                        switch (i)
                        {
                            case 1: this.FindAndReplace(wordApp, "<col" + i + k + ">", dataGridView1.Rows[k].Cells[0].Value.ToString()); break;
                            case 2: this.FindAndReplace(wordApp, "<col" + i + k + ">", dataGridView1.Rows[k].Cells[1].Value.ToString()); break;
                            case 3: this.FindAndReplace(wordApp, "<col" + i + k + ">", dataGridView1.Rows[k].Cells[2].Value.ToString()); break;
                            case 4: this.FindAndReplace(wordApp, "<col" + i + k + ">", dataGridView1.Rows[k].Cells[3].Value.ToString()); break;
                            case 5: this.FindAndReplace(wordApp, "<col" + i + k + ">", dataGridView1.Rows[k].Cells[4].Value.ToString()); break;
                            case 6: this.FindAndReplace(wordApp, "<col" + i + k + ">", dataGridView1.Rows[k].Cells[5].Value.ToString()); break;
                            case 7: this.FindAndReplace(wordApp, "<col" + i + k + ">", dataGridView1.Rows[k].Cells[6].Value.ToString()); break;
                            case 8: this.FindAndReplace(wordApp, "<col" + i + k + ">", dataGridView1.Rows[k].Cells[7].Value.ToString()); break;

                            default:

                                break;
                        }





                    }
                }
                //ures cellák

                for (int k = dataGridView1.RowCount - 1; k < 22; k++)
                {
                    for (int i = 1; i < 9; i++)
                    {
                        switch (i)
                        {
                            case 1: this.FindAndReplace(wordApp, "<col" + i + k + ">", " "); break;
                            case 2: this.FindAndReplace(wordApp, "<col" + i + k + ">", " "); break;
                            case 3: this.FindAndReplace(wordApp, "<col" + i + k + ">", " "); break;
                            case 4: this.FindAndReplace(wordApp, "<col" + i + k + ">", " "); break;
                            case 5: this.FindAndReplace(wordApp, "<col" + i + k + ">", " "); break;
                            case 6: this.FindAndReplace(wordApp, "<col" + i + k + ">", " "); break;
                            case 7: this.FindAndReplace(wordApp, "<col" + i + k + ">", " "); break;
                            case 8: this.FindAndReplace(wordApp, "<col" + i + k + ">", " "); break;

                            default:

                                break;
                        }


                    }

                }


                //this.FindAndReplace(wordApp, "<hasum>", label5.Text);
                this.FindAndReplace(wordApp, "<date>", DateTime.Now.ToShortDateString());
                //tanu
                this.FindAndReplace(wordApp, "<tanu1name>", textBox20.Text);
                this.FindAndReplace(wordApp, "<tanu1address>", textBox21.Text);
                this.FindAndReplace(wordApp, "<tanu2name>", textBox22.Text);
                this.FindAndReplace(wordApp, "<tanu2address>", textBox23.Text);




                //Example of writing to the start of a document.
                //aDoc.Content.InsertBefore("This is at the beginning\r\n\r\n");

                //Example of writing to the end of a document.
                //aDoc.Content.InsertAfter("\r\n\r\nThis is at the end");
            }
            else
            {
                MessageBox.Show("File dose not exist.");
                return;
            }

            //Save the document as the correct file name.
            aDoc.SaveAs(ref saveAs, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing);

            //Close the document - you have to do this.
            aDoc.Close(ref missing, ref missing, ref missing);

            MessageBox.Show("Kész");
        }

        /// <summary>
        /// This is simply a helper method to find/replace 
        /// text.
        /// </summary>
        /// <param name="WordApp">Word Application to use</param>
        /// <param name="findText">Text to find</param>
        /// <param name="replaceWithText">Replacement text</param>
        private void FindAndReplace(Word.Application WordApp,
                                    object findText,
                                    object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object nmatchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            WordApp.Selection.Find.Execute(ref findText,
                ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike,
                ref nmatchAllWordForms, ref forward,
                ref wrap, ref format, ref replaceWithText,
                ref replace, ref matchKashida,
                ref matchDiacritics, ref matchAlefHamza,
                ref matchControl);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string path = Directory.GetCurrentDirectory();

            try
            {
                CreateWordDocument(path + @"\Eredeti.doc",
                                   path + @"\Meghatalmazás.doc");
                System.Diagnostics.Process.Start(path + @"\Meghatalmazás.doc");
            }
            catch (Exception)
            {
                MessageBox.Show("Hiba");
            }
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'hunter8DataSet.megbizott' table. You can move, or remove it, as needed.
            this.megbizottTableAdapter.Fill(this.hunter8DataSet.megbizott);

        }
    }
}
