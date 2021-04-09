using System;
using System.Configuration;
using System.Data;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace program
{

    class Class1
    {
        public string MainConnectionstring = ConfigurationManager.ConnectionStrings["program.Properties.Settings.hunterdataConnectionString"].ConnectionString;
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


        public DataTable kitolt2(string nevstr)
        {
            DataTable D = new DataTable();
            string nev = nevstr;

            string cmdtext = "SELECT[Helyrajzi szám], [Érd#tipus], SUM(Részhányad) as Részhanyadössz,  Név,Meghatalmazott,MeghatalmazasDatuma FROM[Munka1$] GROUP BY[Helyrajzi szám], Név, [Érd#tipus],Meghatalmazott,MeghatalmazasDatuma HAVING(Név LIKE '" + nev + "%')";

            //    "SELECT[Helyrajzi szám],SUM(Részhányad),NÉV, Meghatalmazott, MeghatalmazasDatuma FROM Munka1$   GROUP BY[Helyrajzi szám], NÉV, Meghatalmazott,MeghatalmazasDatuma HAVING(Név LIKE N'" + nev + "%')";
            //Érd#tipus,
            D = adatleker2(cmdtext);

            return D;
        }

        public DataTable meghatkitolt(string meghatnev)
        {
            DataTable D = new DataTable();
            string nev = meghatnev;

            string cmdtext = "SELECT[Munka1$].[Helyrajzi szám], [Munka1$].[Érd#tipus], [Munka1$].Részhányad, [Munka1$].Név, [Munka1$].Meghatalmazott FROM [Munka1$] WHERE((([Munka1$].Meghatalmazott)='" + meghatnev + "'))";





            D = adatleker2(cmdtext);

            return D;
        }
        public DataTable kitoltneveshelyrajzi(string nevstr, string helyrajzistr)
        {
            DataTable D = new DataTable();
            string helyrajzi = helyrajzistr;
            string nev = nevstr;

            //string selecttext = "SELECT [Helyrajzi szám],Érd#tipus, Érdekelt, Hányad,SUBSTRING(Érdekelt, 0, CHARINDEX('cím', Érdekelt)) AS NÉV, Meghatalmazott FROM Munka1$ ";
            //string filtertext = "WHERE(SUBSTRING(Érdekelt, 0, CHARINDEX('cím', Érdekelt)) IS NOT NULL) AND(Név LIKE N'" + nev + "%')";
            //string cmdtext = selecttext + filtertext;
            string cmdtext = "SELECT[Helyrajzi szám], [Érd#tipus], SUM(Részhányad) as Részhanyadössz,Érdekelt,Név,Meghatalmazott,MeghatalmazasDatuma FROM  [Munka1$]  GROUP BY[Helyrajzi szám], Név, [Érd#tipus],Meghatalmazott,MeghatalmazasDatuma,Érdekelt  HAVING ([Helyrajzi szám] LIKE  '" + helyrajzi + "%') AND  (Név LIKE '" + nevstr + "%') ";

            D = adatleker2(cmdtext);

            //dataGridView1.DataSource = D;
            foreach (DataRow d in D.Rows)
            {
                try
                {
                    var pos2 = d[3].ToString().IndexOf('/');
                    //textBox3.Text = d[3].ToString().Substring(pos2 + 1);

                }
                catch (Exception)
                {
                }

            }
            return D;
        }
        public DataTable kitoltnevesmeghatalmazott(string nevstr, string meghatstr)
        {
            DataTable D = new DataTable();
            string meghat = meghatstr;
            string nev = nevstr;

            //string selecttext = "SELECT [Helyrajzi szám],Érd#tipus, Érdekelt, Hányad,SUBSTRING(Érdekelt, 0, CHARINDEX('cím', Érdekelt)) AS NÉV, Meghatalmazott FROM Munka1$ ";
            //string filtertext = "WHERE(SUBSTRING(Érdekelt, 0, CHARINDEX('cím', Érdekelt)) IS NOT NULL) AND(Név LIKE N'" + nev + "%')";
            //string cmdtext = selecttext + filtertext;
            string cmdtext = "SELECT[Helyrajzi szám], [Érd#tipus], SUM(Részhányad) as Részhanyadössz,Érdekelt,Név,Meghatalmazott,MeghatalmazasDatuma FROM  [Munka1$]  GROUP BY[Helyrajzi szám], Név, [Érd#tipus],Meghatalmazott,MeghatalmazasDatuma,Érdekelt  HAVING ([Meghatalmazott] LIKE  '" + meghat + "%') AND  (Név LIKE '" + nevstr + "%') ";

            D = adatleker2(cmdtext);

            //dataGridView1.DataSource = D;
            foreach (DataRow d in D.Rows)
            {
                try
                {
                    var pos2 = d[3].ToString().IndexOf('/');
                    //textBox3.Text = d[3].ToString().Substring(pos2 + 1);

                }
                catch (Exception)
                {
                }

            }
            return D;
        }
        public DataTable berletidijszamolas(string nevstr, string helyrajzistr, string berletihakent)
        {
            DataTable D = new DataTable();
            DataTable D2 = new DataTable();
            string nev = nevstr;
            string helyrajzi = helyrajzistr;

            string cmdtext = "SELECT[Helyrajzi szám], SUM(Részhányad) AS Részhanyadössz, [teljes hányad], Név FROM[Munka1$] GROUP BY[Helyrajzi szám], [teljes hányad], Név HAVING(Név LIKE '" + nev + "%') AND ([Helyrajzi szám] LIKE '" + helyrajzi + "%')";

            D = adatleker2(cmdtext);

            string terulet = string.Empty;
            double berletidij = 0;
            DataTable D4 = new DataTable();
            D4.Columns.AddRange(new DataColumn[8] {
                new DataColumn("Helyrajzi"),
                 new DataColumn("részterület"),
                new DataColumn("részhanyadössz"),
                new DataColumn("teljeshanyad"),
                 new DataColumn("terület"),
                new DataColumn("bérleti díj/HA"),
                new DataColumn("bérleti díj"),
                new DataColumn("Név"),




            });
            MessageBox.Show("Elkezdtem");
            foreach (DataRow drow in D.Rows)
            {
                D2 = adatleker2("SELECT SUM(Terület) AS TerületÖsszesen FROM [Munka1$] WHERE([Helyrajzi szám] ='" + drow[0].ToString() + "')");
                foreach (DataRow d in D2.Rows)
                {
                    terulet = d[0].ToString();


                }
                double resz = Convert.ToDouble(drow[1].ToString());
                double teljes = Convert.ToDouble(drow[2].ToString());
                double reszterulet = (resz / teljes) * Convert.ToDouble(terulet);

                berletidij = reszterulet * Convert.ToDouble(berletihakent);
                int berletidjegesz = Convert.ToInt32(berletidij);
                D4.Rows.Add(drow[0].ToString(), reszterulet, drow[1].ToString(), drow[2].ToString(), terulet, berletihakent + "FT/HA", berletidjegesz, drow[3].ToString());
            }
            MessageBox.Show("vége");

            Double Total = 0;
            foreach (DataRow drow in D4.Rows)
            {
                Total += Convert.ToDouble(drow[6].ToString());
            }
            D4.Rows.Add("TOTAL", "", "", "", "", "", Total.ToString());
            return D4;
        }
         public void adatrogz5(string meg, string nev, string hely, string meghatdatum)
        {
            string ConnetionString = MainConnectionstring;
                //@"Data Source = (LocalDB)\MSSQLLocalDB; AttachDbFilename = 'c:\users\ge70840\documents\visual studio 2015\Projects\Hunterfull8\program\hunter8.mdf'; Integrated Security = True";


            using (SqlConnection sqlcon = new SqlConnection(ConnetionString))
            {
                string sqlcommandtext = "UPDATE [dbo].[Munka1$]  SET [Meghatalmazott]=   '" + meg + "',[MeghatalmazasDatuma]='" + meghatdatum + "'   WHERE [Helyrajzi szám]  LIKE'" + hely + "' and [Név] LIKE '" + nev + "%'";


                SqlCommand sqlcomm = new SqlCommand(sqlcommandtext, sqlcon);
                sqlcon.Open();
                sqlcomm.ExecuteNonQuery();
                sqlcon.Close();

            }




        }
        public void adatrogz4(string meg, string nev, string meghatdatum)
        {
            string ConnetionString = MainConnectionstring;
                //@"Data Source = (LocalDB)\MSSQLLocalDB; AttachDbFilename = 'c:\users\ge70840\documents\visual studio 2015\Projects\Hunterfull8\program\hunter8.mdf'; Integrated Security = True";

            using (SqlConnection sqlcon = new SqlConnection(ConnetionString))
            {
                string sqlcommandtext = "UPDATE [dbo].[Munka1$]  SET [Meghatalmazott]=   '" + meg + "',[MeghatalmazasDatuma]='" + meghatdatum + "'   WHERE [Név] LIKE '" + nev + "%'";


                SqlCommand sqlcomm = new SqlCommand(sqlcommandtext, sqlcon);
                sqlcon.Open();
                sqlcomm.ExecuteNonQuery();

            }


        }
        public ComboBox combotölt()
        {
            ComboBox comboBox1 = new ComboBox();
            string cmdtext = "SELECT Meghatalmazott FROM Munka1$ GROUP BY Meghatalmazott";
            DataTable DD = new DataTable();
            DD = adatleker2(cmdtext);
            foreach (DataRow dd in DD.Rows)
            {
                if (dd[0].ToString() != String.Empty)
                {
                    comboBox1.Items.Add(dd[0].ToString());
                }


            }
            return comboBox1;
        }

        public void meghatalmazottrakeres()
        {

        }

        public double teruletszamolo(string helyrajzi)
        {
            double ter = 0;
            DataTable D2 = new DataTable();
            D2 = adatleker2("SELECT SUM(Terület) AS TerületÖsszesen FROM Munka1$ WHERE ([Helyrajzi szám] ='" + helyrajzi + "')");


            foreach (DataRow d in D2.Rows)
            {
                try
                {
                    ter = Convert.ToDouble(d[0]);
                }
                catch (Exception)
                {
                }

            }
            return ter;

        }
        public string osszegez2(string meghatstr, DataGridView grd)
        {
            string osszegezstr = null;
            //string helyrajzi = null;/// "SZÉCSÉNY/K/269";// "HUGYAG/K/100/27";// "CSITÁR /K/173"; 
            string meghat = meghatstr;
            DataTable D = new DataTable();
            string cmdtext = "SELECT[Helyrajzi szám], Érd#tipus,SUM(Részhányad),[teljes hányad],Meghatalmazott   FROM Munka1$ GROUP BY[Helyrajzi szám], Érd#tipus,[teljes hányad],Meghatalmazott HAVING(Meghatalmazott LIKE N'" + meghat + "%')  ORDER BY [Helyrajzi szám]";
            //and[Helyrajzi szám] = 'ORHALOM/K/166/39'
            //Név,Név,

            D = adatleker2(cmdtext);


            DataTable D2 = new DataTable();
            D2.Columns.AddRange(new DataColumn[7] {

                new DataColumn(D.Columns[0].ColumnName),
                new DataColumn(D.Columns[4].ColumnName),
                new DataColumn(D.Columns[2].ColumnName),
                new DataColumn("Terület %-ban"),
                new DataColumn("Terület HA-ban"),
                new DataColumn("Terület"),
                new DataColumn("Felhasználható Terület"),
            });

            foreach (DataRow drow in D.Rows)
            {
                try
                {
                    int aa = Convert.ToInt32(drow[3].ToString());
                    double terulet = teruletszamolo(drow[0].ToString());
                    int aaa = Convert.ToInt32(drow[2].ToString());
                    object a = (Decimal)aaa / (Decimal)aa;
                    object b = (Decimal)terulet * (Decimal)a;
                    var a2 = (Decimal)a * 100;
                    decimal szavazhato;
                    if (a2 > 50)
                    {
                        szavazhato = (decimal)terulet;
                    }
                    else
                    {
                        szavazhato = 0;
                    }
                    D2.Rows.Add(drow[0].ToString(), drow[4].ToString(), drow[2].ToString(), a2, b, terulet, szavazhato);



                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                grd.DataSource = D2;
                grd.Columns[0].Width = 200;





            }

            decimal Total = 0;


            for (int i = 0; i < grd.Rows.Count; i++)
            {
                try
                {
                    Total += Convert.ToDecimal(grd.Rows[i].Cells[6].Value);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());

                }

            }

            D2.Rows.Add("TOTAL", "", "", "", "", "", Total.ToString());
            //foreach (DataGridViewRow row in dataGridView4.Rows)
            //{
            //    if (row.Cells[0].Value.ToString().StartsWith("T"))
            //    {
            //        //row.DefaultCellStyle.BackColor = System.Drawing.Color.Red;

            //    }
            //}
            osszegezstr = Total.ToString();

            return osszegezstr;



        }




    }
}
