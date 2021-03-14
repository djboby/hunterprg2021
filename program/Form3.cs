using System;
using System.Data;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using System.Configuration;
using Word = Microsoft.Office.Interop.Word;

namespace program
{
    public partial class Form3 : Form
    {
        public string MainConnectionstring = ConfigurationManager.ConnectionStrings["program.Properties.Settings.hunter8ConnectionString"].ConnectionString;

        public Form3()
        {
            InitializeComponent();
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'hunter8DataSet.Berletidij' table. You can move, or remove it, as needed.
            this.berletidijTableAdapter.Fill(this.hunter8DataSet.Berletidij);

        }

        private void button2_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                kitolt3(row.Cells[0].Value.ToString());
            }
            
        }
          public DataTable adatleker2(string cmdstr)
        {

            DataTable Dtable = new DataTable();


            string connectionstr = MainConnectionstring; 
                //@"Data Source = (LocalDB)\MSSQLLocalDB; AttachDbFilename = 'c:\users\ge70840\documents\visual studio 2015\Projects\Hunterfull8\program\hunter8.mdf'; Integrated Security = True";
           // MainConnectionstring;
                //
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
                        //connect.Close();

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
        public DataTable kitolt3(string id)
        {
            DataTable D = new DataTable();

            string cmdtext = "SELECT Id, nev, szulhely, szulido, anyjanev, adoazonosito, lakcim, evek, osszeg, osszegbetuvel, kelt, datum FROM Berletidij where id="+id+"";
            D = adatleker2(cmdtext);
            foreach (DataRow d in D.Rows)
            {
                try
                {
                    textBox1.Text = d[1].ToString();
                    textBox2.Text = d[2].ToString();
                    textBox3.Text = d[3].ToString();
                    textBox4.Text = d[4].ToString();
                    textBox5.Text = d[5].ToString();
                    textBox6.Text = d[6].ToString();
                    textBox7.Text = d[7].ToString();
                    textBox8.Text = d[8].ToString();
                    textBox9.Text = d[9].ToString();
                    textBox10.Text= d[0].ToString();



                }
                catch (Exception)
                {
                }

            }
            return D;
        }
        public void torol(string id)
        {
            string ConnetionString = MainConnectionstring;
            //@"Data Source = (LocalDB)\MSSQLLocalDB; AttachDbFilename = 'c:\users\ge70840\documents\visual studio 2015\Projects\Hunterfull8\program\hunter8.mdf'; Integrated Security = True";
            try
            {

                using (SqlConnection con = new SqlConnection(ConnetionString))
                {
                    string commandtext = "DELETE FROM [Berletidij] WHERE ID = '" + id + "'";


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

        private void button4_Click(object sender, EventArgs e)
        {

            torol(textBox10.Text);
            MessageBox.Show("Ok");
        }
        public void adatrogz(string nev, string szulhely,string szulido,string anyjanev,string adoazonosito,string lakcim,string evek,string osszeg,string oszzegbetuvel,string kelt,string datum)
        {


            string connectionstr = @"Data Source = (LocalDB)\MSSQLLocalDB; AttachDbFilename = 'c:\users\ge70840\documents\visual studio 2015\Projects\Hunterfull8\program\hunter8.mdf'; Integrated Security = True";
            //MainConnectionstring;
            //

            SqlConnection con = new SqlConnection(connectionstr);
            try
            {
                string commandtext = "INSERT INTO Berletidij (nev, szulhely,szulido,anyjanev,adoazonosito,lakcim,evek,osszeg,osszegbetuvel,kelt,datum)  VALUES('" + nev + "','"+szulhely+"','"+szulido+"','"+anyjanev+"','"+adoazonosito+"','"+lakcim+"','"+evek+"','"+osszeg+ "','" + oszzegbetuvel + "','" + kelt + "','" + datum + "')";
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
                //con.Close();
            }











        }

        private void button1_Click(object sender, EventArgs e)
        {
            adatrogz(textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text, textBox5.Text, textBox6.Text, textBox7.Text, textBox8.Text, textBox9.Text, dateTimePicker1.Text, dateTimePicker2.Text);
            MessageBox.Show(DialogResult.OK.ToString());
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }
  

       

        private void button6_Click(object sender, EventArgs e)
        {
            this.berletidijTableAdapter.Fill(this.hunter8DataSet.Berletidij);
            
        }

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
                this.FindAndReplace(wordApp, "<name>", textBox1.Text);
                this.FindAndReplace(wordApp, "<bornplace>", textBox2.Text);
                this.FindAndReplace(wordApp, "<borndate>", textBox3.Text);
                this.FindAndReplace(wordApp, "<mothername>", textBox4.Text);
                this.FindAndReplace(wordApp, "<taxnumber>", textBox5.Text);
                this.FindAndReplace(wordApp, "<address>", textBox6.Text);
                this.FindAndReplace(wordApp, "<years>", textBox7.Text);
                this.FindAndReplace(wordApp, "<price>", textBox8.Text);
                this.FindAndReplace(wordApp, "<pricealpabetical>", textBox9.Text);











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

        private void button5_Click(object sender, EventArgs e)
        {
            string path = Directory.GetCurrentDirectory();

            try
            {
                CreateWordDocument(path + @"\eredetikifizetes.doc",
                                   path + @"\kifizetes.doc");
                System.Diagnostics.Process.Start(path + @"\kifizetes.doc");
            }
            catch (Exception)
            {
                MessageBox.Show("Hiba");
            }

        }
    }
}
