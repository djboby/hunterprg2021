using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace program
{
     public partial class Form1 : Form
    {
        Class1 c = new Class1();
        public Form1()
        {
            InitializeComponent();
        }



        public string hanyad = null;



        private void button1_Click(object sender, EventArgs e)
        {
            try
            {

                dataGridView1.DataSource = c.kitolt2(textBox1.Text);
                dataGridView1.Columns[0].Width = 200;
                dataGridView1.Columns[3].Width = 200;
                dataGridView1.Columns[5].Width = 110;
            }
            catch (Exception)
            {
                MessageBox.Show("Error");

            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = c.kitoltneveshelyrajzi(textBox1.Text, textBox2.Text);
            dataGridView1.Columns[0].Width = 200;
            dataGridView1.Columns[5].Width = 200;
            try
            {
                label31.Text = dataGridView1.Rows[0].Cells[3].Value.ToString();
            }
            catch (Exception)
            {
                MessageBox.Show("Error");

            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = c.berletidijszamolas(textBox1.Text, textBox2.Text, textBox4.Text);
            dataGridView1.Columns[0].Width = 200;


        }

        private void button4_Click(object sender, EventArgs e)
        {
            Class1 c = new Class1();
            c.adatrogz4(textBox3.Text, textBox1.Text, dateTimePicker1.Text);
            MessageBox.Show("OK");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = c.kitoltnevesmeghatalmazott(string.Empty, textBox3.Text);
            dataGridView1.Columns[0].Width = 200;
            dataGridView1.Columns[5].Width = 200;
            try
            {
                label31.Text = dataGridView1.Rows[0].Cells[3].Value.ToString();
            }
            catch (Exception)
            {
                MessageBox.Show("Error");

            }

        }

        public void osszegszamol(DataTable d2, string cmdstr, DataGridView grd)
        {
            DataTable D = new DataTable();
            DataTable D2 = new DataTable();

            string cmdtext = cmdstr;
            double terulet = 0;
            D = c.adatleker2(cmdtext);

            foreach (DataRow drow in D.Rows)
            {
                try
                {
                    string[] Terültomb = drow[0].ToString().Split('/');
                    string terulett = null;
                    if (Terültomb.Length > 3)
                    {
                        terulett = Terültomb[2].ToString() + '/' + Terültomb[3].ToString();
                    }
                    else
                    {
                        terulett = Terültomb[2].ToString();
                    }
                    int aa = Convert.ToInt32(drow[3].ToString());
                    D2 = c.adatleker2("SELECT SUM(Terület) AS TerületÖsszesen FROM [Munka1$] WHERE([Helyrajzi szám] ='" + drow[0].ToString() + "')");
                    foreach (DataRow d in D2.Rows)
                    {
                        terulet = Convert.ToDouble(d[0].ToString());
                    }
                    int aaa = Convert.ToInt32(drow[2].ToString()); object a = (Decimal)aaa / (Decimal)aa;
                    object b = (Decimal)terulet * (Decimal)a; var a2 = (Decimal)a * 100;
                    a2 = Math.Round(a2, 5); b = Math.Round((Decimal)b, 5);
                    decimal szavazhato;
                    if (a2 > 50)
                    {
                        szavazhato = (decimal)terulet;
                    }
                    else
                    {
                        szavazhato = 0;
                    }
                    d2.Rows.Add(Terültomb[0].ToString(), Terültomb[1].ToString(), terulett, drow[5].ToString(), terulet, aaa, b, drow[1].ToString(), drow[6].ToString());

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());

                }
                grd.DataSource = d2;


            }

        }

        public string osszegez4(string nevstr, string helyrejzistr)
        {
            string osszegezstr = null;
            string helyrajzi = helyrejzistr;
            string meghat = null;
            string nev = nevstr;


            DataTable D2 = new DataTable();
            D2.Columns.AddRange(new DataColumn[9] {

                new DataColumn("Terület fekvése"),
                new DataColumn("Külvagybel"),
                new DataColumn("Hszám"),
                new DataColumn("Művelési ág"),
                new DataColumn("Öszz Terület"),
                new DataColumn("Tulajdoni hányad"),
                new DataColumn("Terület HA-ban"),
                new DataColumn("Rendelkezési jogcím"),
                new DataColumn("Név")


            });
            DataTable D33 = new DataTable();
            string cmdtext = "SELECT[Helyrajzi szám]  FROM [Munka1$] WHERE(Meghatalmazott LIKE '" + meghat + "%')  ORDER BY [Helyrajzi szám]";


            //D33 = c.adatleker2(cmdtext);

            osszegszamol(D2, "SELECT[Helyrajzi szám], [Érd#tipus],SUM(Részhányad) as reszhanyadossz,[teljes hányad],Meghatalmazott,[Műv#ág],Név FROM [Munka1$]  GROUP BY[Helyrajzi szám], [Érd#tipus],[teljes hányad],Meghatalmazott,[Műv#ág],Név HAVING(Név LIKE '" + nev + "%') and([Helyrajzi szám] like '" + helyrajzi + "%')  ORDER BY[Helyrajzi szám]", dataGridView1);



            decimal Total = 0;


            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                try
                {
                    Total += Convert.ToDecimal(dataGridView1.Rows[i].Cells[6].Value);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());

                }

            }




            D2.Rows.Add("TOTAL", "", "", "", "", "", Total.ToString());
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells[0].Value.ToString().StartsWith("T"))
                {
                    row.DefaultCellStyle.BackColor = System.Drawing.Color.Red;

                }
            }
            osszegezstr = Total.ToString();
            dataGridView1.DataSource = D2;
            dataGridView1.Columns[0].Width = 200;

            return osszegezstr;



        }


        private void button6_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = c.kitoltneveshelyrajzi(textBox1.Text, textBox2.Text);
            label31.Text = dataGridView1.Rows[0].Cells[3].Value.ToString();

            label32.Text = osszegez4(textBox1.Text, textBox2.Text);

        }


        private void button7_Click(object sender, EventArgs e)
        {

            //Creating iTextSharp Table from the DataTable data
            PdfPTable pdfTable = new PdfPTable(dataGridView1.ColumnCount);
            pdfTable.DefaultCell.Padding = 3;
            pdfTable.WidthPercentage = 100;
            pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
            pdfTable.DefaultCell.BorderWidth = 1;

            //Adding Header row
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                PdfPCell cell = new PdfPCell(new Phrase(column.HeaderText));
                cell.BackgroundColor = new iTextSharp.text.Color(240, 240, 240);
                pdfTable.AddCell(cell);
            }

            //Adding DataRow
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    pdfTable.AddCell(cell.Value.ToString());
                }
            }

            //Exporting to PDF
            string path = Directory.GetCurrentDirectory();

            string folderPath = path;
            //@"C:\Users\ge70840\Documents\Visual Studio 2015\Projects\HunterFullPrg\HunterFullPrg\PDF\";
            //@"C:\Users\ge70840\Documents\Visual Studio 2015\Projects\Hunter\";
            //if (!Directory.Exists(folderPath))
            //{
            //    Directory.CreateDirectory(folderPath);
            //}
            using (FileStream stream = new FileStream(folderPath + @"\Nyomtatvany.pdf", FileMode.Create))
            {
                Document pdfDoc = new Document(PageSize.A2, 10f, 10f, 10f, 0f);
                PdfWriter.GetInstance(pdfDoc, stream);
                pdfDoc.Open();
                pdfDoc.Add(pdfTable);
                pdfDoc.Close();
                stream.Close();
            }
            MessageBox.Show("kész");
            path = Directory.GetCurrentDirectory();
            string pathcut = path;
            //path.Substring(0, path.Length - 10);
            System.Diagnostics.Process.Start(pathcut + @"\Nyomtatvany.pdf");


        }



        private void button8_Click(object sender, EventArgs e)
        {
            //Class1 c = new Class1();
            //c.adatrogz5(textBox3.Text, textBox1.Text, textBox2.Text, dateTimePicker1.Text);
            ////public void adatrogz5(string meg, string nev, string hely, string meghatdatum)
            MessageBox.Show("ERROR");


        }

        private void button9_Click(object sender, EventArgs e)
        {
            Class1 c = new Class1();
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                //if (row.Cells[0].Value.ToString().StartsWith("T"))
                //{
                //    row.DefaultCellStyle.BackColor = System.Drawing.Color.Red;

                //}

                string helyrajzi = row.Cells[0].Value.ToString();
                string nev = row.Cells[3].Value.ToString();
                c.adatrogz5(string.Empty, nev, helyrajzi, string.Empty);
            }
            MessageBox.Show("ok");





        }

        private void button10_Click(object sender, EventArgs e)
        {
            Class1 c = new Class1();
            if (radioButton1.Checked)
            {
                foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                {
                    string helyrajzi = row.Cells[0].Value.ToString();
                    string nev = row.Cells[4].Value.ToString();
                    c.adatrogz5("tulajdonos", nev, helyrajzi, "tulajdonos");
                }

            }
            if (radioButton2.Checked)
            {

            
                foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                {
                    string helyrajzi = row.Cells[0].Value.ToString();
                    string nev = row.Cells[4].Value.ToString();
                    c.adatrogz5(textBox3.Text, nev, helyrajzi, dateTimePicker1.Text);
                }

            }




            MessageBox.Show("OK");


        }

        private void button11_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = c.kitoltnevesmeghatalmazott(textBox1.Text, textBox3.Text);
            dataGridView1.Columns[0].Width = 200;
            dataGridView1.Columns[5].Width = 200;
            try
            {
                label31.Text = dataGridView1.Rows[0].Cells[3].Value.ToString();
            }
            catch (Exception)
            {
                MessageBox.Show("Error");

            }
        }
        public void helyrajzieshanyad()
        {


            try
            {


                DataTable D = new DataTable();
                string helyrajzi = textBox2.Text;
                //string erdekelt = textBox2.Text;
                //textBox5.Text = textBox2.Text;
                string selecttext = "SELECT [Helyrajzi szám],Érd#tipus, Érdekelt, Hányad,SUBSTRING(Érdekelt, 0, CHARINDEX('cím', Érdekelt)) AS NÉV, Meghatalmazott FROM Munka1$ ";
                string filtertext = "WHERE(SUBSTRING(Érdekelt, 0, CHARINDEX('cím', Érdekelt)) IS NOT NULL) AND([Helyrajzi szám] LIKE '" + helyrajzi + "%')";
                string cmdtext = selecttext + filtertext;

                D = c.adatleker2(cmdtext);



                dataGridView1.DataSource = D;
                foreach (DataRow d in D.Rows)
                {
                    try
                    {
                        var pos2 = d[3].ToString().IndexOf('/');
                        hanyad = d[3].ToString().Substring(pos2 + 1);
                        label6.Text = hanyad;

                    }
                    catch (Exception)
                    {
                    }

                }
            }
            catch (Exception ext)
            {

                MessageBox.Show(ext.ToString());


            }



        }



        public void teruletszamolo()
        {
            string helyrajzi = textBox2.Text;

            DataTable D2 = new DataTable();

            D2 = c.adatleker2("SELECT SUM(Terület) AS TerületÖsszesen FROM Munka1$ WHERE ([Helyrajzi szám] ='" + helyrajzi + "')");
            double dddd = 0;
            foreach (DataRow d in D2.Rows)
            {
                try
                {
                    dddd = Convert.ToDouble(d[0]);
                }
                catch (Exception)
                {
                }

            }
            double terület = dddd;
            label7.Text = terület.ToString() + "HA";




        }

        public void szazalekeshektarszamolo()
        {

            string helyrajzi = textBox2.Text;
            //string erdekelt = textBox1.Text;
            double dddd = 0;
            DataTable D2 = new DataTable();

            D2 = c.adatleker2("SELECT SUM(Terület) AS TerületÖsszesen FROM Munka1$ WHERE ([Helyrajzi szám] ='" + helyrajzi + "')");

            foreach (DataRow d in D2.Rows)
            {
                try
                {
                    dddd = Convert.ToDouble(d[0]);
                }
                catch (Exception)
                {
                }

            }
            double terület = dddd;


            DataTable D3 = new DataTable();

            D3 = c.adatleker2("SELECT SUM(Részhányad) AS HÁNYADÖSSZ, Név,Meghatalmazott FROM Munka1$ GROUP BY Meghatalmazott,Név, [Helyrajzi szám] HAVING ([Helyrajzi szám] ='" + helyrajzi + "')");
            List<string> lista1 = new List<string>();
            List<string> lista2 = new List<string>();
            List<string> lista3 = new List<string>();
            List<string> lista4 = new List<string>();
            //List<string> lista5 = new List<string>();
            DataTable D4 = new DataTable();
            D4.Columns.AddRange(new DataColumn[4] {
                new DataColumn("Név"),
                new DataColumn("Hányadössz"),
                new DataColumn("Tulajdon%"),
                new DataColumn("TulajdonHA"),

            });

            foreach (DataRow drow in D3.Rows)
            {

                lista1.Add(drow[1].ToString());
                lista2.Add(drow[0].ToString());
                try
                {
                    double szam = Convert.ToDouble(drow[0].ToString());
                    double szam2 = (szam / Convert.ToInt32(hanyad)) * 100;
                    lista3.Add(szam2.ToString());
                    double s4 = (terület * szam2) / 100;
                    lista4.Add(s4.ToString());
                    D4.Rows.Add(drow[1].ToString(), drow[0].ToString(), szam2.ToString(), s4.ToString());
                }
                catch (Exception ex)
                {
                    //  MessageBox.Show(ex.ToString());


                }

            }

            //double a = 0;
            //double b = 0;
            //foreach (DataRow drow in D4.Rows)
            //{

            //    if (drow[4].ToString() == label5.Text)
            //    {
            //        a = a + Convert.ToDouble(drow[3]);
            //        b = b + Convert.ToDouble(drow[2]);


            //    }

            //}


            dataGridView1.DataSource = D4;
            //textBox4.Text = a.ToString();
            //textBox5.Text = b.ToString();
            //if (Convert.ToDouble(textBox8.Text) > 50)
            //{
            //    label5.Text = "Teljes Területtel szavazhat ";
            //}
            //else
            //{
            //    label5.Text = "Nem szavazhat teljes Területtel ";
            //}





        }

        private void button13_Click(object sender, EventArgs e)
        {
            helyrajzieshanyad();
            teruletszamolo();
            szazalekeshektarszamolo();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button12_Click(object sender, EventArgs e)
        {
            DataTable D4 = new DataTable();
            try
            {
                D4.Columns.AddRange(new DataColumn[8] {

                new DataColumn(dataGridView1.Columns[0].HeaderText),
                new DataColumn(dataGridView1.Columns[1].HeaderText),
                new DataColumn(dataGridView1.Columns[2].HeaderText),
                new DataColumn(dataGridView1.Columns[3].HeaderText),
                new DataColumn(dataGridView1.Columns[4].HeaderText),
                new DataColumn(dataGridView1.Columns[5].HeaderText),
                new DataColumn(dataGridView1.Columns[6].HeaderText),
                new DataColumn(dataGridView1.Columns[7].HeaderText),
                 });

            }
            catch (Exception)
            {
                MessageBox.Show("Error");

            }

            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                D4.Rows.Add(row.Cells[0].Value.ToString(), row.Cells[1].Value.ToString(), row.Cells[2].Value.ToString(), row.Cells[3].Value.ToString(), row.Cells[4].Value.ToString(), row.Cells[5].Value.ToString(), row.Cells[6].Value.ToString(), row.Cells[7].Value.ToString());
            }
            dataGridView1.DataSource = D4;

        }

        private void button14_Click(object sender, EventArgs e)
        {
            Form2 f = new Form2(label31.Text, dataGridView1, "LABEL5");
            f.Show();
            //MessageBox.Show("Később");
        }

        private void button15_Click(object sender, EventArgs e)
        {
            DataTable d = new DataTable();
            d.Columns.AddRange(new DataColumn[2] {
                    new DataColumn("Név"),
                    new DataColumn("Felhasználható Terület"),
                });

            string cmdtext = "SELECT Meghatalmazott FROM Munka1$ GROUP BY Meghatalmazott";
            DataTable DD = new DataTable();
            DD = c.adatleker2(cmdtext);
            foreach (DataRow dd in DD.Rows)
            {
                if (dd[0].ToString() != String.Empty)
                {
                    d.Rows.Add(dd[0].ToString(), c.osszegez2(dd[0].ToString(), dataGridView1));
                }


            }

            //d.Rows.Add("Nánai Andrea", osszegez2("Nánai"));
            dataGridView1.DataSource = d;

        }

        private void button16_Click(object sender, EventArgs e)
        {
            comboBox1.Items.Clear();
            string cmdtext = "SELECT DISTINCT Meghatalmazott FROM     Munka1$ WHERE(Meghatalmazott IS NOT NULL)";
            DataTable DD = new DataTable();
            DD = c.adatleker2(cmdtext);
            foreach (DataRow dd in DD.Rows)
            {
                if (dd[0].ToString() != String.Empty)
                {
                    comboBox1.Items.Add(dd[0].ToString());
                }


            }

        }


        private void button18_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Meghatalmazás Módosítása");
        }

        private void button17_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Meghatalmazás Törlése");
        }

        private void button20_Click(object sender, EventArgs e)
        {
            Form3 f = new Form3();
            f.Show();
        }

        private void button17_Click_1(object sender, EventArgs e)
        {
            dataGridView1.DataSource = c.kitoltneveshelyrajzi(String.Empty, textBox2.Text);
            dataGridView1.Columns[0].Width = 200;
            dataGridView1.Columns[5].Width = 200;
            try
            {
                label31.Text = dataGridView1.Rows[0].Cells[3].Value.ToString();
            }
            catch (Exception)
            {
                MessageBox.Show("Error");

            }
        }
    }
}
