using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using Excel = Microsoft.Office.Interop.Excel;
namespace NewVersionDip
{

    class WorkHelper
    {
        SqlConnection connection;
        public double[] g = new double[13];
        public double[] g1 = new double[13];
        public double[] g2 = new double[13];
        public double[] g3 = new double[13];
        List<DPNag> DP = new List<DPNag>();
        List<ZNag> Z = new List<ZNag>();
        bool flag = false;
        bool flagZ = false;
        public WorkHelper()
        {
            connection = new SqlConnection(Properties.Settings.Default.connString);
        }

        public void Clearer()
        {
            Z.Clear();
            DP.Clear();
            flag = false;
        }
        public Dictionary<int, string> Mounth = new Dictionary<int, string>
        {
            {1,"Января"},
            {2,"Февраля"},
            {3,"Марта"},
            {4,"Апреля"},
            {5,"Мая"},
            {6,"Июня"},
            {7,"Июля"},
            {8,"Августа"},
            {9,"Сентября"},
            {10,"Октября"},
            {11,"Ноября"},
            {12,"Декабря"}
        };
        public Dictionary<int, string> MounthZ = new Dictionary<int, string>
        {
            {1,"Январь"},
            {2,"Февраль"},
            {3,"Март"},
            {4,"Апрель"},
            {5,"Май"},
            {6,"Июнь"},
            {7,"Июль"},
            {8,"Август"},
            {9,"Сентябрь"},
            {10,"Октябрь"},
            {11,"Ноябрь"},
            {12,"Декабрь"}
        };
        public Dictionary<int, string> MounthE = new Dictionary<int, string>
        {
            {1,"Январе"},
            {2,"Феврале"},
            {3,"Марте"},
            {4,"Апреле"},
            {5,"Мае"},
            {6,"Июне"},
            {7,"Июле"},
            {8,"Августе"},
            {9,"Сентябре"},
            {10,"Октябре"},
            {11,"Ноябре"},
            {12,"Декабре"}
        };
        public Dictionary<int, string> Day = new Dictionary<int, string>
        {
            {1,"31"},
            {2,"28"},
            {3,"31"},
            {4,"30"},
            {5,"31"},
            {6,"30"},
            {7,"31"},
            {8,"31"},
            {9,"30"},
            {10,"31"},
            {11,"30"},
            {12,"31"}
        };
        public void comboBoxFuller()
        {
            //try
            //{
                string queryNaDob = @"SELECT T_Family, T_FirstName,T_SecondNamme  FROM Teacher";
                string queryNaDob1 = @"SELECT G_Family, G_FirstName,G_SecondNamme  FROM GEK";
            connection.Open();
                SqlDataAdapter sqlQuery = new SqlDataAdapter(queryNaDob, connection);
                System.Data.DataTable dtblDob = new System.Data.DataTable();
                sqlQuery.Fill(dtblDob);
                string[] strSave = new string[dtblDob.Rows.Count];
                for (int i = 0; i < dtblDob.Rows.Count; i++)
                {
                    for (int j = 0; j < dtblDob.Columns.Count; j++)
                    {
                        strSave[i] += dtblDob.Rows[i][j].ToString();
                        if (j != dtblDob.Columns.Count - 1)
                        {
                            strSave[i] += " ";
                        }

                    }
                }
            for (int i = 0; i < dtblDob.Rows.Count; i++)
            {
                Program.cent.FIOcombo = strSave[i];
            }
            SqlDataAdapter sqlQuery2 = new SqlDataAdapter(queryNaDob1, connection);
                System.Data.DataTable dtblDob1 = new System.Data.DataTable();
                sqlQuery2.Fill(dtblDob1);
                string[] strSave1 = new string[dtblDob1.Rows.Count];
                for (int i = 0; i < dtblDob1.Rows.Count; i++)
                {
                    for (int j = 0; j < dtblDob1.Columns.Count; j++)
                    {
                        strSave1[i] += dtblDob1.Rows[i][j].ToString();
                        if (j != dtblDob1.Columns.Count - 1)
                        {
                            strSave1[i] += " ";
                        }

                    }
                }
                for (int i = 0; i < dtblDob1.Rows.Count; i++)
                {
                    Program.cent.GEKcombo = strSave1[i];
                }
                Program.cent.Faccom1 = "КИТП";
                Program.cent.Faccom1 = "ИПФМИ";
                

                string query = @"SELECT name FROM Cafedra";

                SqlDataAdapter sqlQuery1 = new SqlDataAdapter(query, connection);
                System.Data.DataTable dtbl = new System.Data.DataTable();
                sqlQuery1.Fill(dtbl);

                for (int i = 0; i < dtbl.Rows.Count; i++)
                {
                    Program.cent.Cafcombo = dtbl.Rows[i][0].ToString();
                }
            connection.Close();
            //}
            //catch (Exception ex) { MessageBox.Show(ex.Message); }
            //finally { connection.Close(); }
        }
        
        public void WordFuller(string nameFile)//ОБЩИЙ СЛУЧАЙ изменить под заявления
        {
            try
            {

                var helper = new WordHelper(nameFile);
                string str;
                string[] splitTeacher;
                var helper1 = new ExcelHelper();
                StringBuilder build = new StringBuilder();
                if (Program.cent.Combo3 == 0 || Program.cent.Combo3 == 1 || Program.cent.Combo4 == 0 || Program.cent.Combo4 == 1 || Program.cent.Combo5 == 0 || Program.cent.Combo5 == 1 || Program.cent.Combo6 == 0 || Program.cent.Combo6 == 1)
                {
                    str = Program.cent.GEKcombo;
                    splitTeacher = str.Split();
                    build.Append("SELECT * FROM GEK WHERE G_Family = '" + splitTeacher[0] + "'" + "AND G_FirstName=" + "'" + splitTeacher[1] + "'" + "AND G_SecondNamme = " + "'" + splitTeacher[2] + "'");
                }
                else
                {
                    str = Program.cent.FIOcombo;
                    splitTeacher = str.Split();
                    build.Append("SELECT * FROM Teacher WHERE T_Family = '" + splitTeacher[0] + "'" + "AND T_FirstName=" + "'" + splitTeacher[1] + "'" + "AND T_SecondNamme = " + "'" + splitTeacher[2] + "'");
                }

                string query = Convert.ToString(build);
                build.Clear();
                build.Append("SELECT Family, apellation, Sec_Name FROM Cafedra WHERE name = '" + Program.cent.Cafcombo + "'"); //build.Append(Program.cent.Cafcombo); build.Append("'");
                string query1 = Convert.ToString(build);
                build.Clear();
                connection.Open();
                SqlDataAdapter sql = new SqlDataAdapter(query, connection);
                System.Data.DataTable dtbl = new System.Data.DataTable();
                sql.Fill(dtbl);
                SqlDataAdapter sql1 = new SqlDataAdapter(query1, connection);
                System.Data.DataTable dtbl1 = new System.Data.DataTable();
                sql1.Fill(dtbl1);
                StringBuilder ZavCaf = new StringBuilder();
                for (int i = 0; i < dtbl1.Columns.Count; i++)
                {
                    ZavCaf.Append(dtbl1.Rows[0][i] + " ");
                }
                string Zav = ZavCaf.ToString();
                string[] heaher = new string[dtbl.Columns.Count];
                for (int i = 0; i < dtbl.Columns.Count; i++)
                {
                    heaher[i] = dtbl.Rows[0][i].ToString();
                }
                string passport = heaher[8];
                StringBuilder dat = new StringBuilder();
                for (int i = passport.Length - 10; i < passport.Length; i++)
                {
                    dat.Append(heaher[8][i]);
                }
                string[] archpass = new string[3];

                archpass = PassCreator(passport);

                string[] fdr = heaher[5].Split();
                heaher[5] = fdr[0];
                string h = Convert.ToString(dat);
                archpass[2] = archpass[2].Replace(h, "");
                var items = new Dictionary<string, string> { };
                int itemSaver = Convert.ToInt32(Program.cent.Faccom);
                string[] date = heaher[5].ToString().Split(' ');
                if (Program.cent.Combo7 == 0 || Program.cent.Combo7 == 1 || Program.cent.Combo9 == 0 || Program.cent.Combo9 == 1)
                {
                    NagFReader();
                }
                if (Program.cent.Check == true)
                {
                    FullStackDoc(helper, items, heaher, archpass, splitTeacher, h, Zav);

                }
                else
                {
                    string TegQuery = "";//
                    SqlDataAdapter sqlTeg;
                    string[] NewBase;
                    System.Data.DataTable dtblTeg = new System.Data.DataTable();

                    if (itemSaver == 0 || Program.cent.Combo7 == 0 || Program.cent.Combo9 == 0)
                    {
                        TegQuery = "SELECT * FROM TEG";
                        sqlTeg = new SqlDataAdapter(TegQuery, connection);

                        sqlTeg.Fill(dtblTeg);
                        NewBase = new string[dtblTeg.Rows.Count];
                        for (int i = 0; i < dtblTeg.Rows.Count; i++)
                        {
                            dtblTeg.Rows[i][0] = dtblTeg.Rows[i][0].ToString().Replace(" ", "");
                        }
                        NewBase = NewDataSource(heaher, dtblTeg, archpass, splitTeacher, h, Zav);

                        for (int i = 0; i < dtblTeg.Rows.Count; i++)
                        {
                            items.Add(Convert.ToString(dtblTeg.Rows[i][0]), NewBase[i]);
                        }
                    }
                    else if (itemSaver == 1 || Program.cent.Combo7 == 1 || Program.cent.Combo9 == 1)
                    {
                        TegQuery = "SELECT * FROM TEG1";
                        sqlTeg = new SqlDataAdapter(TegQuery, connection);
                        sqlTeg.Fill(dtblTeg);
                        NewBase = new string[dtblTeg.Rows.Count];
                        for (int i = 0; i < dtblTeg.Rows.Count; i++)
                        {
                            dtblTeg.Rows[i][0] = dtblTeg.Rows[i][0].ToString().Replace(" ", "");
                        }
                        NewBase = NewDataSource(heaher, dtblTeg, archpass, splitTeacher, h, Zav);

                        for (int i = 0; i < dtblTeg.Rows.Count; i++)
                        {
                            items.Add(Convert.ToString(dtblTeg.Rows[i][0]), NewBase[i]);
                        }
                    }


                    helper.Process(items);
                    string std = helper.ReOpen();
                    Process.Start(std);
                    

                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            finally { connection.Close(); }

        }
        public void WordZFuller(string nameFile)//ОБЩИЙ СЛУЧАЙ изменить под заявления
        {
            try
            {
                var helper = new WordHelper(nameFile);
                string str;

                string[] splitTeacher;
                var helper1 = new ExcelHelper();
                StringBuilder build = new StringBuilder();

                str = Program.cent.FIOcombo;
                splitTeacher = str.Split();
                build.Append("SELECT * FROM Teacher WHERE T_Family = '" + splitTeacher[0] + "'" + "AND T_FirstName=" + "'" + splitTeacher[1] + "'" + "AND T_SecondNamme = " + "'" + splitTeacher[2] + "'");


                string query = Convert.ToString(build);
                build.Clear();
                build.Append("SELECT Family, apellation, Sec_Name FROM Cafedra WHERE name = '" + Program.cent.Cafcombo + "'"); //build.Append(Program.cent.Cafcombo); build.Append("'");
                string query1 = Convert.ToString(build);
                build.Clear();
                connection.Open();
                SqlDataAdapter sql = new SqlDataAdapter(query, connection);
                System.Data.DataTable dtbl = new System.Data.DataTable();
                sql.Fill(dtbl);
                SqlDataAdapter sql1 = new SqlDataAdapter(query1, connection);
                System.Data.DataTable dtbl1 = new System.Data.DataTable();
                sql1.Fill(dtbl1);
                StringBuilder ZavCaf = new StringBuilder();
                for (int i = 0; i < dtbl1.Columns.Count; i++)
                {
                    ZavCaf.Append(dtbl1.Rows[0][i] + " ");
                }
                string Zav = ZavCaf.ToString();
                string[] heaher = new string[dtbl.Columns.Count];
                for (int i = 0; i < dtbl.Columns.Count; i++)
                {
                    heaher[i] = dtbl.Rows[0][i].ToString();
                }
                string passport = heaher[8];
                StringBuilder dat = new StringBuilder();
                for (int i = passport.Length - 10; i < passport.Length; i++)
                {
                    dat.Append(heaher[8][i]);
                }
                string[] archpass = new string[3];

                archpass = PassCreator(passport);

                string[] fdr = heaher[5].Split();
                heaher[5] = fdr[0];
                string h = Convert.ToString(dat);
                archpass[2] = archpass[2].Replace(h, "");
                var items = new Dictionary<string, string> { };
                int itemSaver = Convert.ToInt32(Program.cent.Faccom);
                string[] date = heaher[5].ToString().Split(' ');

                NagZFReader();
                if (flagZ == true)
                {
                    MessageBox.Show("Рабочая нагрузка на два полугодия превышает 100 часов!");
                    return;
                }
                items = DataSourceZ(items);
                string TegQuery = "";//
                SqlDataAdapter sqlTeg;
                string[] NewBase;
                System.Data.DataTable dtblTeg = new System.Data.DataTable();

                if (itemSaver == 0)
                {
                    TegQuery = "SELECT * FROM TEG";
                    sqlTeg = new SqlDataAdapter(TegQuery, connection);

                    sqlTeg.Fill(dtblTeg);
                    NewBase = new string[dtblTeg.Rows.Count];
                    for (int i = 0; i < dtblTeg.Rows.Count; i++)
                    {
                        dtblTeg.Rows[i][0] = dtblTeg.Rows[i][0].ToString().Replace(" ", "");
                    }
                    NewBase = NewDataSource(heaher, dtblTeg, archpass, splitTeacher, h, Zav);

                    for (int i = 0; i < dtblTeg.Rows.Count; i++)
                    {
                        items.Add(Convert.ToString(dtblTeg.Rows[i][0]), NewBase[i]);
                    }
                }

                else if (itemSaver == 1)
                {
                    TegQuery = "SELECT * FROM TEG1";
                    sqlTeg = new SqlDataAdapter(TegQuery, connection);
                    sqlTeg.Fill(dtblTeg);
                    NewBase = new string[dtblTeg.Rows.Count];
                    for (int i = 0; i < dtblTeg.Rows.Count; i++)
                    {
                        dtblTeg.Rows[i][0] = dtblTeg.Rows[i][0].ToString().Replace(" ", "");
                    }
                    NewBase = NewDataSource(heaher, dtblTeg, archpass, splitTeacher, h, Zav);

                    for (int i = 0; i < dtblTeg.Rows.Count; i++)
                    {
                        items.Add(Convert.ToString(dtblTeg.Rows[i][0]), NewBase[i]);
                    }
                }


                helper.Process(items);
                string std = helper.ReOpen();
                Process.Start(std);
               
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            finally { connection.Close(); }
            

        }
        public struct ZNagSaver 
        {
           
            public string disp;
            public double hour;
        };
        public Dictionary<string, string> DataSourceZ(Dictionary<string, string> items)
        {
            double count = 0;
            int summary = 0;
            int chet = 1;
            int i = 0;
            ZNagSaver Box = new ZNagSaver();
            foreach (var nag in Z)
            {
                count += (nag.Lek + nag.Lab + nag.TecCon + nag.Zach + nag.Exc + nag.DPRuc + nag.DPPredZ + nag.KonKExc);
                summary++;
            }
            //MessageBox.Show(summary.ToString());
            items.Add("<NAG1>", count.ToString());
            items.Add("<NAG>", count.ToString());
            items.Add("<1>", MouthZ(Convert.ToInt32(DateTime.Now.ToString("MM"))));
            foreach (var nag in Z)
            {

                Box.disp = "Лекции";
                Box.hour = nag.Lek;
                if(Box.hour > 0 && Box.disp.Equals("Лекции"))
                {
                    chet += 1;
                    items.Add("<" + chet.ToString() + ">", Box.disp);
                    chet += 1;
                    items.Add("<" + (chet).ToString() + ">", Box.hour.ToString());
                }
                
                Box.disp = "Лабораторнвые работы";
                Box.hour = nag.Lab;
                if (Box.hour > 0 && Box.disp.Equals("Лабораторнвые работы"))
                {
                    items.Add("<" + (chet + 1).ToString() + ">", Box.disp);
                    chet += 1;
                    items.Add("<" + (chet + 1).ToString() + ">", Box.hour.ToString());
                    chet += 1;
                }
                
                Box.disp = "текущ.конс.";
                Box.hour = nag.TecCon;
                if (Box.hour > 0 && Box.disp.Equals("текущ.конс."))
                {
                    items.Add("<" + (chet + 1).ToString() + ">", Box.disp);
                    chet += 1;
                    items.Add("<" + (chet + 1).ToString() + ">", Box.hour.ToString());
                    chet += 1;
                }
                
                Box.disp = "Зачёты";
                Box.hour = nag.Zach;
                if (Box.hour > 0 && Box.disp.Equals("Зачёты"))
                {
                    items.Add("<" + (chet + 1).ToString() + ">", Box.disp);
                    chet += 1;
                    items.Add("<" + (chet + 1).ToString() + ">", Box.hour.ToString());
                    chet += 1;
                }
                
                Box.disp = "Экзамены";
                Box.hour = nag.Exc;
                if (Box.hour > 0 && Box.disp.Equals("Экзамены"))
                {
                    items.Add("<" + (chet + 1).ToString() + ">", Box.disp);
                    chet += 1;
                    items.Add("<" + (chet + 1).ToString() + ">", Box.hour.ToString());
                    chet += 1;
                }
                
                Box.disp = "Консультации к экзаменам";
                Box.hour = nag.KonKExc;
                if (Box.hour > 0 && Box.disp.Equals("Консультации к экзаменам"))
                {
                    items.Add("<" + (chet + 1).ToString() + ">", Box.disp);
                    chet += 1;
                    items.Add("<" + (chet + 1).ToString() + ">", Box.hour.ToString());
                    chet += 1;
                }
                
                Box.disp = "ДП руководство";
                Box.hour = nag.DPRuc;
                if (Box.hour > 0 && Box.disp.Equals("ДП руководство"))
                {
                    items.Add("<" + (chet + 1).ToString() + ">", Box.disp);
                    chet += 1;
                    items.Add("<" + (chet + 1).ToString() + ">", Box.hour.ToString());
                    chet += 1;
                }
                
                Box.disp = "ДП предзащита";
                Box.hour = nag.DPPredZ;
                if (Box.hour > 0 && Box.disp.Equals("ДП предзащита"))
                {
                    items.Add("<" + (chet + 1).ToString() + ">", Box.disp);
                    chet += 1;
                    items.Add("<" + (chet + 1).ToString() + ">", Box.hour.ToString());
                    chet += 1;
                }
                


            }

            //MessageBox.Show(i.ToString());
            for (int j = chet+1; j < 28; j++)
            {
                //MessageBox.Show(i.ToString());
                items.Add("<" + j.ToString() + ">", "");
            }


            return items;
        }
        public void WordZNPOFuller(string nameFile)
        {
            try
            {
                var helper = new WordHelper(nameFile);
                string str;
                string[] splitTeacher;
                var helper1 = new ExcelHelper();
                StringBuilder build = new StringBuilder();
                if (Program.cent.Combo3 == 0 || Program.cent.Combo3 == 1 || Program.cent.Combo4 == 0 || Program.cent.Combo4 == 1 || Program.cent.Combo5 == 0 || Program.cent.Combo5 == 1 || Program.cent.Combo6 == 0 || Program.cent.Combo6 == 1)
                {
                    str = Program.cent.GEKcombo;
                    splitTeacher = str.Split();
                    build.Append("SELECT * FROM GEK WHERE G_Family = '" + splitTeacher[0] + "'" + "AND G_FirstName=" + "'" + splitTeacher[1] + "'" + "AND G_SecondNamme = " + "'" + splitTeacher[2] + "'");
                }
                else
                {
                    str = Program.cent.FIOcombo;
                    splitTeacher = str.Split();
                    build.Append("SELECT * FROM Teacher WHERE T_Family = '" + splitTeacher[0] + "'" + "AND T_FirstName=" + "'" + splitTeacher[1] + "'" + "AND T_SecondNamme = " + "'" + splitTeacher[2] + "'");
                }

                string query = Convert.ToString(build);
                build.Clear();
                build.Append("SELECT Family, apellation, Sec_Name FROM Cafedra WHERE name = '" + Program.cent.Cafcombo + "'"); //build.Append(Program.cent.Cafcombo); build.Append("'");
                string query1 = Convert.ToString(build);
                NagFReader();

                if (flag == true)
                {
                    MessageBox.Show("Рабочая нагрузка на два полугодия превышает 300 часов!");
                    return;
                }

                build.Clear();
                connection.Open();
                SqlDataAdapter sql = new SqlDataAdapter(query, connection);
                System.Data.DataTable dtbl = new System.Data.DataTable();
                sql.Fill(dtbl);
                SqlDataAdapter sql1 = new SqlDataAdapter(query1, connection);
                System.Data.DataTable dtbl1 = new System.Data.DataTable();
                sql1.Fill(dtbl1);
                StringBuilder ZavCaf = new StringBuilder();
                for (int i = 0; i < dtbl1.Columns.Count; i++)
                {
                    ZavCaf.Append(dtbl1.Rows[0][i] + " ");
                }
                string Zav = ZavCaf.ToString();
                string[] heaher = new string[dtbl.Columns.Count];
                for (int i = 0; i < dtbl.Columns.Count; i++)
                {
                    heaher[i] = dtbl.Rows[0][i].ToString();
                }
                string passport = heaher[8];
                StringBuilder dat = new StringBuilder();
                for (int i = passport.Length - 10; i < passport.Length; i++)
                {
                    dat.Append(heaher[8][i]);
                }
                string[] archpass = new string[3];

                archpass = PassCreator(passport);

                string[] fdr = heaher[5].Split();
                heaher[5] = fdr[0];
                string h = Convert.ToString(dat);
                archpass[2] = archpass[2].Replace(h, "");
                var items = new Dictionary<string, string> { };
                int itemSaver = Convert.ToInt32(Program.cent.Faccom);
                string[] date = heaher[5].ToString().Split(' ');

                if (Program.cent.Check == true)
                {
                    FullStackDoc(helper, items, heaher, archpass, splitTeacher, h, Zav);

                }
                else
                {
                    string TegQuery = "";//
                    SqlDataAdapter sqlTeg;
                    string[] NewBase;
                    System.Data.DataTable dtblTeg = new System.Data.DataTable();

                    if (Program.cent.Combo7 == 0 || Program.cent.Combo9 == 0)
                    {
                        TegQuery = "SELECT * FROM TEG";
                        sqlTeg = new SqlDataAdapter(TegQuery, connection);

                        sqlTeg.Fill(dtblTeg);
                        NewBase = new string[dtblTeg.Rows.Count];
                        for (int i = 0; i < dtblTeg.Rows.Count; i++)
                        {
                            dtblTeg.Rows[i][0] = dtblTeg.Rows[i][0].ToString().Replace(" ", "");
                        }
                        NewBase = NewDataSource(heaher, dtblTeg, archpass, splitTeacher, h, Zav);

                        for (int i = 0; i < dtblTeg.Rows.Count; i++)
                        {
                            items.Add(Convert.ToString(dtblTeg.Rows[i][0]), NewBase[i]);
                        }
                    }
                    else if (Program.cent.Combo7 == 1 || Program.cent.Combo9 == 1)
                    {
                        TegQuery = "SELECT * FROM TEG1";
                        sqlTeg = new SqlDataAdapter(TegQuery, connection);
                        sqlTeg.Fill(dtblTeg);
                        NewBase = new string[dtblTeg.Rows.Count];
                        for (int i = 0; i < dtblTeg.Rows.Count; i++)
                        {
                            dtblTeg.Rows[i][0] = dtblTeg.Rows[i][0].ToString().Replace(" ", "");
                        }
                        NewBase = NewDataSource(heaher, dtblTeg, archpass, splitTeacher, h, Zav);

                        for (int i = 0; i < dtblTeg.Rows.Count; i++)
                        {
                            items.Add(Convert.ToString(dtblTeg.Rows[i][0]), NewBase[i]);
                        }
                    }
                    int count = 0;
                    if (Program.cent.Combo9 == 1)
                    {
                        for (int i = 0; i < 26; i++)
                        {
                            if (i < 26 / 2)
                            {

                                items.Add("<" + i.ToString() + i.ToString() + i.ToString() + ">", g1[i].ToString());
                            }

                            else
                            {
                                items.Add("<" + i.ToString() + i.ToString() + i.ToString() + ">", g3[count].ToString());
                                count++;//MessageBox.Show(g3[count].ToString());
                            }
                        }
                        items.Add("<" + "NAG1" + ">", Convert.ToString(g1[12] + g3[12]));

                    }
                    else if (Program.cent.Combo9 == 0)
                    {
                        for (int i = 0; i < 26; i++)
                        {
                            if (i < 26 / 2)
                                items.Add("<" + i.ToString() + ">", g[i].ToString());
                            else
                            {
                                items.Add("<" + i.ToString() + ">", g2[count].ToString());
                                count++;
                            }
                        }
                        items.Add("<" + "NAG" + ">", Convert.ToString(g[12] + g2[12]));
                    }
                    else if (Program.cent.Combo7 == 0)
                    {
                        for (int i = 0; i < 26; i++)
                        {
                            if (i < 26 / 2)
                                items.Add("<" + i.ToString() + ">", g[i].ToString());
                            else
                            {
                                items.Add("<" + i.ToString() + ">", g2[count].ToString());
                                count++;
                            }
                        }
                        items.Add("<" + "NAG" + ">", Convert.ToString(g[12] + g2[12]));
                    }
                    else if (Program.cent.Combo7 == 1)
                    {
                        for (int i = 0; i < 26; i++)
                        {
                            if (i < 26 / 2)
                            {

                                items.Add("<" + i.ToString() + i.ToString() + i.ToString() + ">", g1[i].ToString());
                            }

                            else
                            {

                                items.Add("<" + i.ToString() + i.ToString() + i.ToString() + ">", g3[count].ToString());
                                count++;//MessageBox.Show(g3[count].ToString());
                            }
                        }
                        items.Add("<" + "NAG1" + ">", Convert.ToString(g1[12] + g3[12]));
                    }

                    helper.Process(items);
                    string std = helper.ReOpen();
                    Process.Start(std);
                }
            }

            catch (Exception ex) { MessageBox.Show(ex.Message); }
            finally { connection.Close(); }
                
            
        }
        public void WordDPFuller(string nameFile)
        {
            try
            {
                var helper = new WordHelper(nameFile);
                string str;
                string[] splitTeacher;
                var helper1 = new ExcelHelper();
                StringBuilder build = new StringBuilder();
                str = Program.cent.GEKcombo;
                splitTeacher = str.Split();
                build.Append("SELECT * FROM GEK WHERE G_Family = '" + splitTeacher[0] + "'" + "AND G_FirstName=" + "'" + splitTeacher[1] + "'" + "AND G_SecondNamme = " + "'" + splitTeacher[2] + "'");
                string query = Convert.ToString(build);
                build.Clear();
                build.Append("SELECT Family, apellation, Sec_Name FROM Cafedra WHERE name = '" + Program.cent.Cafcombo + "'"); //build.Append(Program.cent.Cafcombo); build.Append("'");
                string query1 = Convert.ToString(build);
                build.Clear();
                connection.Open();
                SqlDataAdapter sql = new SqlDataAdapter(query, connection);
                System.Data.DataTable dtbl = new System.Data.DataTable();
                sql.Fill(dtbl);
                SqlDataAdapter sql1 = new SqlDataAdapter(query1, connection);
                System.Data.DataTable dtbl1 = new System.Data.DataTable();
                sql1.Fill(dtbl1);
                StringBuilder ZavCaf = new StringBuilder();
                for (int i = 0; i < dtbl1.Columns.Count; i++)
                {
                    ZavCaf.Append(dtbl1.Rows[0][i] + " ");
                }
                string Zav = ZavCaf.ToString();
                string[] heaher = new string[dtbl.Columns.Count];
                for (int i = 0; i < dtbl.Columns.Count; i++)
                {
                    heaher[i] = dtbl.Rows[0][i].ToString();
                }
                string passport = heaher[8];
                StringBuilder dat = new StringBuilder();
                for (int i = passport.Length - 10; i < passport.Length; i++)
                {
                    dat.Append(heaher[8][i]);
                }
                string[] archpass = new string[3];

                archpass = PassCreator(passport);

                string[] fdr = heaher[5].Split();
                heaher[5] = fdr[0];
                string h = Convert.ToString(dat);
                archpass[2] = archpass[2].Replace(h, "");
                var items = new Dictionary<string, string> { };
                int itemSaver = Convert.ToInt32(Program.cent.Faccom);
                string[] date = heaher[5].ToString().Split(' ');
                string TegQuery = "";//
                SqlDataAdapter sqlTeg;
                string[] NewBase;
                System.Data.DataTable dtblTeg = new System.Data.DataTable();
                TegQuery = "SELECT * FROM TEG1";
                sqlTeg = new SqlDataAdapter(TegQuery, connection);
                sqlTeg.Fill(dtblTeg);
                NewBase = new string[dtblTeg.Rows.Count];
                for (int i = 0; i < dtblTeg.Rows.Count; i++)
                {
                    dtblTeg.Rows[i][0] = dtblTeg.Rows[i][0].ToString().Replace(" ", "");
                }
                NewBase = NewDataSource(heaher, dtblTeg, archpass, splitTeacher, h, Zav);
                NagDPFReader();
                items = DataSource(items, heaher);
                for (int i = 0; i < dtblTeg.Rows.Count; i++)
                {
                    items.Add(Convert.ToString(dtblTeg.Rows[i][0]), NewBase[i]);
                }
                double sum = 0;
                foreach (var nag in DP)
                {
                    sum += nag.GAK;
                }
                items.Add("<NAG1>", sum.ToString());
                helper.Process(items);

                string std = helper.ReOpen();
                Process.Start(std);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            finally { connection.Close(); }
        }
        public string MouthZ(int mou)
        {
            foreach (var m in MounthZ)
            {
                if (m.Key == mou)
                {
                    return m.Value;
                }
            }
            return "0";
        }
        public string MouthE(int mou)
        {
            foreach (var m in MounthE)
            {
                if (m.Key == mou)
                {
                    return m.Value;
                }
            }
            return "0";
        }
        public string Mouth(int mou)
        {
            foreach(var m in Mounth)
            {
                if(m.Key == mou)
                {
                    return m.Value;
                }
            }
            return "0";
        }
        public string Days(int mou)
        {
            foreach (var m in Day)
            {
                if (m.Key == mou)
                {
                    return m.Value;
                }
            }
            return "0";
        }
        public string[] NewDataSource(string[] heaher, System.Data.DataTable dtblTeg, string[] archpass, string[] splitTeacher, string dat, string ZavCaf)//отсортировывается набор данных из БД
        {

            string[] NewBase = new string[dtblTeg.Rows.Count];
            string[] dipData = heaher[16].Split(' ');
            string[] fdr = heaher[5].Split();
            NewBase[0] = splitTeacher[0] + " " + splitTeacher[1] + " " + splitTeacher[2]; //
            NewBase[1] = heaher[2];// Фамилия
            NewBase[2] = heaher[3];// Имя
            NewBase[3] = heaher[4];// Отчество
            NewBase[4] = fdr[0];// День рождения
            NewBase[5] = heaher[6];// Учёнове звание
            NewBase[6] = heaher[7];// Учёная степень
            NewBase[7] = archpass[0];// Серия паспорта
            NewBase[8] = archpass[1];// Номер паспорта
            NewBase[9] = archpass[2];// Остальное
            NewBase[10] = heaher[9];// SNILS
            NewBase[11] = heaher[13] + " " + "кв" + " " + heaher[10] + " " + heaher[11] + " " + heaher[12];//Address(11, heaher);
            NewBase[12] = heaher[14];
            NewBase[13] = dipData[0] + " " + heaher[20];
            NewBase[14] = heaher[17];
            NewBase[15] = heaher[18];
            NewBase[16] = heaher[19];
            NewBase[17] = DateTime.Now.ToString("yyyy");
            NewBase[18] = Mouth(Convert.ToInt32(DateTime.Now.ToString("MM")));
            NewBase[19] = Program.cent.Cafcombo;
            NewBase[20] = Program.cent.Faccom1;
            NewBase[21] = Days(Convert.ToInt32(DateTime.Now.ToString("MM")));
            NewBase[22] = heaher[21];
            NewBase[23] = splitTeacher[0] + " " + splitTeacher[1][0] + " " + splitTeacher[2][0];
            NewBase[24] = dat;
            NewBase[25] = ZavCaf;
            NewBase[26] = (Convert.ToInt32(DateTime.Now.ToString("yyyy")) + 1).ToString();
            return NewBase;


        }
        
        public Dictionary<string, string> DataSource(Dictionary<string, string> items, string[] heaher)
        {
            foreach (var nag in DP)
            {
                nag.Disp = heaher[18];
            }
            double count = 0;
            int summary = 0;
            int i = 0;
            foreach(var nag in DP)
            {
                count += nag.Itogo;
                summary++;
            }
            //MessageBox.Show(summary.ToString());
            items.Add("<NAG>", count.ToString());
            
            foreach(var nag in DP)
            {
                items.Add("<" + (i).ToString() + ">", nag.Spec.ToString());
                items.Add("<" + (i+1).ToString() + ">", nag.Group.ToString());
                items.Add("<" + (i+2).ToString() + ">", nag.NumOfStud.ToString());
                items.Add("<" + (i+3).ToString() + ">", nag.Disp.ToString());
                items.Add("<" + (i + 4).ToString() + ">", nag.GAK.ToString());
                items.Add("<" + (i + 5).ToString() + ">", nag.Itogo.ToString());
                
                
                i += 6;
            }
            //MessageBox.Show(i.ToString());
            for(int j = i;j< 50; j++)
            {
                //MessageBox.Show(i.ToString());
                items.Add("<" + j.ToString() + ">", "");
            }
            
            
            return items;
        }
        
        public string Address(int i, string[] heaher)//формируется адресс для набора данных
        {
            StringBuilder address = new StringBuilder();
            
            address.Append(heaher[i+2]+ " " + "кв" + " " + heaher[i-1] + " " + heaher[i] + " " + heaher[i+1]);
            return Convert.ToString(address);
        }
        
        public void FullStackDoc(WordHelper helper, Dictionary<string, string> items, string[] heaher, string[] archpass, string[] splitTeacher, string h, string Zav)//полное заполнение заявлений на почасовую оплату
        {
            string TegQuery = "";//
            SqlDataAdapter sqlTeg;
            string[] NewBase;
            System.Data.DataTable dtblTeg = new System.Data.DataTable();
            TegQuery = "SELECT * FROM TEG";
            sqlTeg = new SqlDataAdapter(TegQuery, connection);
            //connection.Open();
            sqlTeg.Fill(dtblTeg);
            NewBase = new string[dtblTeg.Rows.Count];
            for (int i = 0; i < dtblTeg.Rows.Count; i++)
            {
                dtblTeg.Rows[i][0] = dtblTeg.Rows[i][0].ToString().Replace(" ", "");
            }
            NewBase = NewDataSource(heaher, dtblTeg, archpass, splitTeacher, h, Zav);

            for (int i = 0; i < dtblTeg.Rows.Count; i++)
            {
                
                items.Add(Convert.ToString(dtblTeg.Rows[i][0]), NewBase[i]);
            }
            dtblTeg.Clear();
            TegQuery = "SELECT * FROM TEG1";
            sqlTeg = new SqlDataAdapter(TegQuery, connection);
            sqlTeg.Fill(dtblTeg);
            NewBase = new string[dtblTeg.Rows.Count];
            for (int i = 0; i < dtblTeg.Rows.Count; i++)
            {
                dtblTeg.Rows[i][0] = dtblTeg.Rows[i][0].ToString().Replace(" ", "");
            }
            NewBase = NewDataSource(heaher, dtblTeg, archpass, splitTeacher, h, Zav);

            for (int i = 0; i < dtblTeg.Rows.Count; i++)
            {
                //MessageBox.Show(dtblTeg.Rows[i][0].ToString());
                items.Add(Convert.ToString(dtblTeg.Rows[i][0]), NewBase[i]);
            }
            int count = 0;
            //int count1 = 0;

            if (Program.cent.Combo9 == 0 || Program.cent.Combo9 == 1 && Program.cent.Check == true)
            {
                for (int i = 0; i < 26; i++)
                {
                    if (i < 26 / 2)
                        items.Add("<" + i.ToString() + ">", g[i].ToString());
                    else
                    {
                        items.Add("<" + i.ToString() + ">", g2[count].ToString());
                        count++;
                    }
                }
                items.Add("<" + "NAG" + ">", Convert.ToString(g[12] + g2[12]));
            }
            count = 0;
            if (Program.cent.Combo9 == 1 || Program.cent.Combo9 == 0 && Program.cent.Check == true)
            {
                for (int i = 0; i < 26; i++)
                {
                    if (i < 26 / 2)
                        //try { 
                        items.Add("<" + i.ToString() + i.ToString() + i.ToString() + ">", g1[i].ToString()); //} catch (Exception ex) { MessageBox.Show(i.ToString()); }
                    else
                    {
                        items.Add("<" + i.ToString() + i.ToString() + i.ToString() + ">", g3[count].ToString());
                        //MessageBox.Show(count.ToString());
                        count++;
                    }
                }
                items.Add("<" + "NAG1" + ">", Convert.ToString(g1[12] + g3[12]));
            }
            count = 0;
            if (Program.cent.Combo7 == 0 || Program.cent.Combo7 == 1 && Program.cent.Check == true)
            {
                for (int i = 0; i < 26; i++)
                {
                    if (i < 26 / 2)
                        items.Add("<" + i.ToString() + ">", g[i].ToString());
                    else
                    {
                        items.Add("<" + i.ToString() + ">", g2[count].ToString());
                        count++;
                    }
                }
                items.Add("<" + "NAG" + ">", Convert.ToString(g[12] + g2[12]));
            }
            count = 0;
            if (Program.cent.Combo7 == 1 || Program.cent.Combo7 == 0 && Program.cent.Check == true)
            {
                for (int i = 0; i < 26; i++)
                {
                    if (i < 26 / 2)
                        //try { 
                            items.Add("<" + i.ToString() + i.ToString() + i.ToString() + ">", g1[i].ToString()); //} catch (Exception ex) { MessageBox.Show(i.ToString()); }
                    else
                    {
                        items.Add("<" + i.ToString() + i.ToString() + i.ToString() + ">", g3[count].ToString());
                        count++;
                    }
                }
                items.Add("<" + "NAG1" + ">", Convert.ToString(g1[12] + g3[12]));
            }
            helper.Process(items);
            string std = helper.ReOpen();
            Process.Start(std);
            connection.Close();
        }
       
        public String[] PassCreator(string passport)//разбивает паспортные данные
        {
            string[] archpass = new string[3];
            string slovoBuf = "";
            for (int i = 0; i < passport.Length; i++)
            {

                if (i == 4)
                {
                    archpass[0] = slovoBuf;
                    slovoBuf = "";
                }
                else if (i < 4)
                {
                    slovoBuf += passport[i];
                }
                if (i > 4 && i < 11)
                {
                    slovoBuf += passport[i];
                }
                else if (i == 11)
                {
                    archpass[1] = slovoBuf;
                    slovoBuf = "";
                }
                if (i > 11)
                {
                    slovoBuf += passport[i];
                }
            }// разбиение паспортных данных
            archpass[2] = slovoBuf;
            return archpass;
        }

        public void choiceDockA4Excel()
        {
            int sdf = Program.cent.Faccom;
            string exePath;
            
            string pathExc;
            //try
            //{

            switch (Program.cent.Combo1)
            {
                case 0:
                    exePath = AppDomain.CurrentDomain.BaseDirectory;
                    string Query = @"SELECT * FROM ZBV";
                    pathExc = Path.Combine(exePath, "res\\zajavlenie_bjudzhet_vo.xlsx");
                    //MessageBox.Show(pathExc);
                    if (sdf == 1)
                    {

                        ExcelFuller(pathExc, Query);
                    }
                    else
                    {
                        break;
                    }
                    break;
                case 1:
                    exePath = AppDomain.CurrentDomain.BaseDirectory;
                    Query = @"SELECT * FROM ZBV";
                    pathExc = Path.Combine(exePath, "res\\zajavlenie_vnebjudzhet_vo.xlsx");

                    if (sdf == 1)
                    {
                        ExcelFuller(pathExc, Query);
                    }
                    else
                    {
                        break;
                    }
                    break;
            }


            switch (Program.cent.Combo2)
            {
                case 0:

                    exePath = AppDomain.CurrentDomain.BaseDirectory;
                    string Query = @"SELECT * FROM ZBV";
                    pathExc = Path.Combine(exePath, "res\\zajavlenie_bjudzhet_spo.xlsx");
                    if (sdf == 0)
                    {
                        ExcelFuller(pathExc, Query);
                    }
                    else
                    {
                        break;
                    }
                    break;
                case 1:

                    exePath = AppDomain.CurrentDomain.BaseDirectory;
                    Query = @"SELECT * FROM ZBV";
                    pathExc = Path.Combine(exePath, "res\\zajavlenie_vnebjudzhet_spo.xlsx");
                    if (sdf == 0)
                    {
                        ExcelFuller(pathExc, Query);
                    }
                    else
                    {
                        break;
                    }
                    break;
            }
            //}
            //catch (Exception ex) { MessageBox.Show(ex.Message); }
        }
        public void choiceDockA5Word()
        {
            int sdf = Program.cent.Faccom;
            
            string exePath;
            string path;
            
            //try
            //{

            switch (Program.cent.Combo5)
            {
                case 0:

                    exePath = AppDomain.CurrentDomain.BaseDirectory;
                    path = Path.Combine(exePath, "res\\akt_budget_vo.doc");
                    if (sdf == 1)
                    {
                        WordDPFuller(path);
                    }
                    else
                    {
                        break;
                    }
                    break;
                case 1:

                    exePath = AppDomain.CurrentDomain.BaseDirectory;
                    path = Path.Combine(exePath, "res\\akt_vnebudget_vo.doc");
                    if (sdf == 1)
                        WordDPFuller(path); 
                    break;

            }


            switch (Program.cent.Combo6)
            {
                case 0:

                    exePath = AppDomain.CurrentDomain.BaseDirectory;
                    path = Path.Combine(exePath, "res\\akt_budget_spo.doc");
                    if (sdf == 0)
                        WordDPFuller(path); 
                    break;
                case 1:

                    exePath = AppDomain.CurrentDomain.BaseDirectory;
                    path = Path.Combine(exePath, "res\\akt_vnebudget_spo.doc"); 
                    if (sdf == 0)
                        WordDPFuller(path); 
                    break;
            }


            switch (Program.cent.Combo7)
            {
                case 0:

                    exePath = AppDomain.CurrentDomain.BaseDirectory;
                    path = Path.Combine(exePath, "res\\zayavlenie_na_pochasovuyu_oplatu_byudzhet.doc"); 
                    if (sdf == 1)
                        WordZNPOFuller(path); 
                    break;
                case 1:

                    exePath = AppDomain.CurrentDomain.BaseDirectory;
                    path = Path.Combine(exePath, "res\\zayavlenie_na_pochasovuyu_oplatu_byudzhet.doc"); 
                    if (sdf == 1)
                        WordZNPOFuller(path);
                    break;

            }
            switch (Program.cent.Combo9)
            {
                case 0:

                    exePath = AppDomain.CurrentDomain.BaseDirectory;
                    path = Path.Combine(exePath, "res\\zayavlenie_na_pochasovuyu_oplatu_vnebyudzhet.doc");
                    if (sdf == 0)
                        WordZNPOFuller(path);
                    break;
                case 1:

                    exePath = AppDomain.CurrentDomain.BaseDirectory;
                    path = Path.Combine(exePath, "res\\zayavlenie_na_pochasovuyu_oplatu_vnebyudzhet.doc");
                    if (sdf == 0)
                        WordZNPOFuller(path);
                    break;

            }
            switch (Program.cent.Combo3)
            {
                case 0:

                    exePath = AppDomain.CurrentDomain.BaseDirectory;
                    path = Path.Combine(exePath, "res\\dogovor_podryada_budget_vo_iii.doc");
                    if (sdf == 1)
                        WordDPFuller(path);
                    break;
                case 1:

                    exePath = AppDomain.CurrentDomain.BaseDirectory;
                    path = Path.Combine(exePath, "res\\dogovor_podryada_vnebudget_vo_iii.doc");
                    if (sdf == 1)
                        WordDPFuller(path);
                    break;

            }
            switch (Program.cent.Combo4)
            {
                case 0:

                    exePath = AppDomain.CurrentDomain.BaseDirectory;
                    path = Path.Combine(exePath, "res\\dogovor_podryada_budget_spo_iii.doc");
                    if (sdf == 0)
                        WordDPFuller(path);
                    break;
                case 1:

                    exePath = AppDomain.CurrentDomain.BaseDirectory;
                    path = Path.Combine(exePath, "res\\dogovor_podryada_vnebudget_spo_iii.doc");
                    if (sdf == 0)
                        WordDPFuller(path);
                    break;

            }

            //}
            //catch(Exception ex) { MessageBox.Show(ex.Message); }
        }
        public void choiceDockA5Excel()
        {
            int sdf = Program.cent.Faccom;
            string exePath;
            
            string pathExc;
            //try
            //{

            switch (Program.cent.Combo5)
            {
                case 0:

                    exePath = AppDomain.CurrentDomain.BaseDirectory;
                    pathExc = Path.Combine(exePath, "res\\akt_budget_vo.xlsx");
                    string Query = @"SELECT * FROM aktBV";
                    if (sdf == 1)
                    {
                        ExcelFuller(pathExc, Query);
                    }
                    else
                    {
                        break;
                    }
                    break;
                case 1:

                    exePath = AppDomain.CurrentDomain.BaseDirectory;
                    pathExc = Path.Combine(exePath, "res\\akt_vnebudget_vo.xlsx");
                    Query = @"SELECT * FROM aktBV";
                    if (sdf == 1)
                        ExcelFuller(pathExc, Query);
                    break;

            }


            switch (Program.cent.Combo6)
            {
                case 0:

                    exePath = AppDomain.CurrentDomain.BaseDirectory;
                     pathExc = Path.Combine(exePath, "res\\akt_budget_spo.xlsx");
                    string Query = @"SELECT * FROM aktBV";
                    if (sdf == 0)
                        ExcelFuller(pathExc, Query);
                    break;
                case 1:

                    exePath = AppDomain.CurrentDomain.BaseDirectory;
                    pathExc = Path.Combine(exePath, "res\\akt_vnebudget_spo.xlsx");
                    Query = @"SELECT * FROM aktBV";
                    if (sdf == 0)
                        ExcelFuller(pathExc, Query);
                    break;
            }

            switch (Program.cent.Combo7)
            {
                case 0:

                    exePath = AppDomain.CurrentDomain.BaseDirectory;
                    pathExc = Path.Combine(exePath, "res\\zayavlenie_na_pochasovuyu_oplatu_byudzhet.xlsx");
                    if (sdf == 1)
                        ExeclDocFullerPOV(pathExc);
                    break;
                case 1:

                    exePath = AppDomain.CurrentDomain.BaseDirectory;
                    pathExc = Path.Combine(exePath, "res\\zayavlenie_na_pochasovuyu_oplatu_byudzhet.xlsx");
                    if (sdf == 1)
                        ExeclDocFullerPOV(pathExc);
                    break;

            }
            switch (Program.cent.Combo9)
            {
                case 0:

                    exePath = AppDomain.CurrentDomain.BaseDirectory;
                    pathExc = Path.Combine(exePath, "res\\zayavlenie_na_pochasovuyu_oplatu_vnebyudzhet.xlsx");
                    if (sdf == 0)
                        ExeclDocFullerPOV(pathExc);
                    break;
                case 1:

                    exePath = AppDomain.CurrentDomain.BaseDirectory;
                    pathExc = Path.Combine(exePath, "res\\zayavlenie_na_pochasovuyu_oplatu_vnebyudzhet.xlsx");
                    if (sdf == 0)
                        ExeclDocFullerPOV(pathExc);
                    break;

            }
            switch (Program.cent.Combo3)
            {
                case 0:

                    exePath = AppDomain.CurrentDomain.BaseDirectory;
                    pathExc = Path.Combine(exePath, "res\\dag_podryada_bjuVO.xlsx");
                    string Query = "select * from DP";
                    if (sdf == 1)
                        ExcelFuller(pathExc, Query);
                    break;
                case 1:

                    exePath = AppDomain.CurrentDomain.BaseDirectory;
                    Query = "select * from DP";
                    pathExc = Path.Combine(exePath, "res\\dag_podryada_vneVO.xlsx");
                    if (sdf == 1)
                        ExcelFuller(pathExc, Query);
                    break;

            }
            switch (Program.cent.Combo4)
            {
                case 0:

                    exePath = AppDomain.CurrentDomain.BaseDirectory;
                    pathExc = Path.Combine(exePath, "res\\dag_podryada_bjuSPO.xlsx");
                    string Query = "select * from DP";
                    if (sdf == 0)
                        ExcelFuller(pathExc, Query);
                    break;
                case 1:

                    exePath = AppDomain.CurrentDomain.BaseDirectory;
                    pathExc = Path.Combine(exePath, "res\\dag_podryada_vneSPO.xlsx");
                    Query = "select * from DP";
                    if (sdf == 0)
                        ExcelFuller(pathExc, Query);
                    break;

            }

            //}
            //catch(Exception ex) { MessageBox.Show(ex.Message); }
        }
        public void choiceDockA4Word()
        {
            int sdf = Program.cent.Faccom;
            string exePath;
            string path;
            
            //try
            //{
           
                switch (Program.cent.Combo1)
                {
                    case 0:
                        exePath = AppDomain.CurrentDomain.BaseDirectory;
                        path = Path.Combine(exePath, "res\\Zajavlenie_bjudzhet_VO.doc");
                        
                        if (sdf == 1)
                        {
                        WordZFuller(path);
                            
                        }
                        else
                        {
                            break;
                        }
                        break;
                    case 1:
                        exePath = AppDomain.CurrentDomain.BaseDirectory;
                        path = Path.Combine(exePath, "res\\zajavlenie_vnebjudzhet_vo.doc");
                        

                        if (sdf == 1)
                        {
                        WordZFuller(path);
                            
                        }
                        else
                        {
                            break;
                        }
                        break;
                }
            
            
                switch (Program.cent.Combo2)
                {
                    case 0:

                        exePath = AppDomain.CurrentDomain.BaseDirectory;
                        path = Path.Combine(exePath, "res\\zajavlenie_bjudzhet_spo.doc");
                        
                        if (sdf == 0)
                        {
                        WordZFuller(path);
                        }
                        else
                        {
                            break;
                        }
                        break;
                    case 1:

                        exePath = AppDomain.CurrentDomain.BaseDirectory;
                        path = Path.Combine(exePath, "res\\zajavlenie_vnebjudzhet_spo.doc");
                        
                        if (sdf == 0)
                        {
                        WordZFuller(path);
                        }
                        else
                        {
                            break;
                        }
                        break;
                }
            



            //}
            //catch (Exception ex) { MessageBox.Show(ex.Message); }

        }
        public void ExcelFuller(string path, string Query)
        {
            using (ExcelHelper helper = new ExcelHelper())
            {
                try
                {
                    if (helper.Open(filePath: path))
                    {

                        string str;
                        string[] splitTeacher;
                        var helper1 = new ExcelHelper();
                        StringBuilder build = new StringBuilder();
                        if (Program.cent.Combo3 == 0 || Program.cent.Combo3 == 1 || Program.cent.Combo4 == 0 || Program.cent.Combo4 == 1 || Program.cent.Combo5 == 0 || Program.cent.Combo5 == 1 || Program.cent.Combo6 == 0 || Program.cent.Combo6 == 1)
                        {
                            str = Program.cent.GEKcombo;
                            splitTeacher = str.Split();
                            build.Append("SELECT * FROM GEK WHERE G_Family = '" + splitTeacher[0] + "'" + "AND G_FirstName=" + "'" + splitTeacher[1] + "'" + "AND G_SecondNamme = " + "'" + splitTeacher[2] + "'");
                        }
                        else
                        {
                            str = Program.cent.FIOcombo;
                            splitTeacher = str.Split();
                            build.Append("SELECT * FROM Teacher WHERE T_Family = '" + splitTeacher[0] + "'" + "AND T_FirstName=" + "'" + splitTeacher[1] + "'" + "AND T_SecondNamme = " + "'" + splitTeacher[2] + "'");
                        }

                        string query = Convert.ToString(build);
                        build.Clear();
                        build.Append("SELECT * FROM Cafedra WHERE name = '"); build.Append(Program.cent.Cafcombo); build.Append("'");
                        string query1 = Convert.ToString(build);
                        build.Clear();
                        connection.Open();
                        SqlDataAdapter sql = new SqlDataAdapter(query, connection);
                        System.Data.DataTable dtbl = new System.Data.DataTable();
                        System.Data.DataTable dtbl1 = new System.Data.DataTable(); ;
                        sql.Fill(dtbl);

                        build.Append("SELECT Family, apellation, Sec_Name FROM Cafedra WHERE name = '" + Program.cent.Cafcombo + "'"); //build.Append(Program.cent.Cafcombo); build.Append("'");
                        string query3 = Convert.ToString(build);
                        SqlDataAdapter sql3 = new SqlDataAdapter(query3, connection);
                        System.Data.DataTable dtbl2 = new System.Data.DataTable(); ;
                        sql3.Fill(dtbl2);
                        StringBuilder ZavCaf = new StringBuilder();
                        for (int i = 0; i < dtbl2.Columns.Count; i++)
                        {
                            ZavCaf.Append(dtbl2.Rows[0][i] + " ");
                        }
                        string Zav = ZavCaf.ToString();

                        string[] heaher = new string[dtbl.Columns.Count];
                        for (int i = 0; i < dtbl.Columns.Count; i++)
                        {
                            heaher[i] = dtbl.Rows[0][i].ToString();

                        }

                        string[] archpass = new string[3];
                        archpass = PassCreator(heaher[8]);
                        string[] fdr = heaher[5].Split();

                        SqlDataAdapter sqlQuery = new SqlDataAdapter(Query, connection);
                        System.Data.DataTable dtblCod = new System.Data.DataTable();
                        sqlQuery.Fill(dtblCod);
                        for (int i = 0; i < dtblCod.Rows.Count; i++)
                        {

                            dtblCod.Rows[i][0] = dtblCod.Rows[i][0].ToString().Replace(" ", "");
                        }
                        string[] hardSave = new string[dtblCod.Rows.Count];
                        if (Program.cent.Combo1 == 1 || Program.cent.Combo1 == 0 || Program.cent.Combo2 == 1 || Program.cent.Combo2 == 0)
                        {
                            
                            NagZFReader();
                            if (flagZ == true)
                            {
                                MessageBox.Show("Рабочая нагрузка на два полугодия превышает 100 часов!");
                                return;
                            }
                            hardSave = NabDataZ(heaher, archpass, hardSave, Zav);
                            ZNAGCreator(helper);
                            
                            
                        }
                        else if (Program.cent.Combo3 == 1 || Program.cent.Combo3 == 0 || Program.cent.Combo4 == 1 || Program.cent.Combo4 == 0)
                        {
                            NagDPFReader();
                            double sum = NagWorkerExcelDP(helper, heaher);
                            hardSave = NabDataDP(heaher, archpass, hardSave, sum);


                        }
                        else if (Program.cent.Combo5 == 1 || Program.cent.Combo5 == 0 || Program.cent.Combo6 == 1 || Program.cent.Combo6 == 0)
                        {
                            NagDPFReader();
                            double sum = NagWorkerExcelAkt(helper, heaher);

                            hardSave = NabDataAkt(heaher, archpass, hardSave, sum);
                        }

                        for (int i = 0; i < dtblCod.Rows.Count; i++)
                        {
                            //MessageBox.Show((dtblCod.Rows[i][0]+ dtblCod.Rows[i][1].ToString() + " " + hardSave[i]));
                            helper.Set(row: Convert.ToInt32(dtblCod.Rows[i][1]), data: hardSave[i], column: Convert.ToString(dtblCod.Rows[i][0]));

                        }
                    }
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }

                FileInfo info = new FileInfo(path);
                string direct = info.Directory.ToString();
                string name = info.Name.ToString();
                Object newFileName = Path.Combine(info.Directory.ToString(), DateTime.Now.ToString("yyyyMMdd") + info.Name);


                //MessageBox.Show("1");
                helper.Save(newFileName.ToString());
                    
                //MessageBox.Show("2");
                
                    Process.Start(newFileName.ToString());
                    connection.Close();
                

            }

        }
        public double NagWorkerExcelAkt(ExcelHelper helper, string[] heaher)
        {
            double sum = 0;
            try
            {
                foreach (var nag in DP)
                {
                    nag.Disp = heaher[18];
                }
                string Query = "SELECT * FROM AKTNAG";
                SqlDataAdapter sqlQuery = new SqlDataAdapter(Query, connection);
                System.Data.DataTable dtblCod = new System.Data.DataTable();
                sqlQuery.Fill(dtblCod);
                for (int i = 0; i < dtblCod.Rows.Count; i++)
                {

                    dtblCod.Rows[i][0] = dtblCod.Rows[i][0].ToString().Replace(" ", "");
                }
                int count = 0, i1 = 0;
                
                foreach (var nag in DP)
                {

                    helper.Set(row: Convert.ToInt32(dtblCod.Rows[count][1]), data: nag.Spec, column: Convert.ToString(dtblCod.Rows[count][0]));
                    helper.Set(row: Convert.ToInt32(dtblCod.Rows[count + 1][1]), data: nag.Group, column: Convert.ToString(dtblCod.Rows[count + 1][0]));
                    helper.Set(row: Convert.ToInt32(dtblCod.Rows[count + 2][1]), data: nag.NumOfStud, column: Convert.ToString(dtblCod.Rows[count + 2][0]));
                    helper.Set(row: Convert.ToInt32(dtblCod.Rows[count + 3][1]), data: nag.Disp, column: Convert.ToString(dtblCod.Rows[count + 3][0]));
                    helper.Set(row: Convert.ToInt32(dtblCod.Rows[count + 4][1]), data: nag.GAK, column: Convert.ToString(dtblCod.Rows[count + 4][0]));
                    helper.Set(row: Convert.ToInt32(dtblCod.Rows[count + 5][1]), data: nag.Itogo, column: Convert.ToString(dtblCod.Rows[count + 5][0]));

                    count = count + 6;
                    sum += nag.Itogo;
                }
                return sum;
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            return 0;
        }
        public double NagWorkerExcelDP(ExcelHelper helper, string[] heaher)
        {
            try
            {

                foreach(var nag in DP)
                {
                    nag.Disp = heaher[18];
                }
                string Query = "SELECT * FROM DPNAG";
                SqlDataAdapter sqlQuery = new SqlDataAdapter(Query, connection);
                System.Data.DataTable dtblCod = new System.Data.DataTable();
                sqlQuery.Fill(dtblCod);
                for (int i = 0; i < dtblCod.Rows.Count; i++)
                {

                    dtblCod.Rows[i][0] = dtblCod.Rows[i][0].ToString().Replace(" ", "");
                }
                int count = 0, i1 = 0;
                double sum = 0;
                foreach (var nag in DP)
                {

                    helper.Set(row: Convert.ToInt32(dtblCod.Rows[count][1]), data: nag.Spec, column: Convert.ToString(dtblCod.Rows[count][0]));
                    helper.Set(row: Convert.ToInt32(dtblCod.Rows[count + 1][1]), data: nag.Group, column: Convert.ToString(dtblCod.Rows[count + 1][0]));
                    helper.Set(row: Convert.ToInt32(dtblCod.Rows[count + 2][1]), data: nag.NumOfStud, column: Convert.ToString(dtblCod.Rows[count + 2][0]));
                    helper.Set(row: Convert.ToInt32(dtblCod.Rows[count + 3][1]), data: nag.Disp, column: Convert.ToString(dtblCod.Rows[count + 3][0]));
                    helper.Set(row: Convert.ToInt32(dtblCod.Rows[count + 4][1]), data: nag.GAK, column: Convert.ToString(dtblCod.Rows[count + 4][0]));
                    helper.Set(row: Convert.ToInt32(dtblCod.Rows[count + 5][1]), data: nag.Itogo, column: Convert.ToString(dtblCod.Rows[count + 5][0]));

                    count = count + 6;
                    sum += nag.Itogo;
                }

                return sum;
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            return 0;
        }
        public string[] NabDataDP(string[] heaher, string[] archpass, string[] hardSave, double sum)
        {
            try
            {


                string passport = heaher[8];
                string dat = "";
                for (int i = passport.Length - 10; i < passport.Length; i++)
                {
                    dat += heaher[8][i];
                }
                string[] fdr = heaher[5].Split(' ');
                archpass[2] = archpass[2].Replace(dat, "");
                string[] splitTeacher = new string[3];
                splitTeacher[0] = heaher[2];
                splitTeacher[1] = heaher[3];
                splitTeacher[2] = heaher[4];
                hardSave[0] = (heaher[2] + " " + heaher[3] + " " + heaher[4]);
                hardSave[1] = archpass[0] + " " + archpass[1];
                hardSave[2] = archpass[2] + " " + dat;

                hardSave[3] = hardSave[0];
                hardSave[4] = heaher[17];

                hardSave[5] = heaher[18];
                hardSave[6] = heaher[6] + " " + heaher[7];
                hardSave[7] = fdr[0];
                hardSave[8] = archpass[0];
                hardSave[9] = archpass[1];
                hardSave[10] = archpass[2];
                hardSave[11] = dat;

                hardSave[12] = heaher[13] + " " + heaher[10] + " " + heaher[11] + " " + heaher[12] + " " + heaher[14];
                hardSave[13] = heaher[9];
                hardSave[14] = heaher[18];
                hardSave[15] = Days(Convert.ToInt32(DateTime.Now.ToString("MM")));
                hardSave[16] = Mouth(Convert.ToInt32(DateTime.Now.ToString("MM")));
                hardSave[17] = DateTime.Now.ToString("yyyy");
                hardSave[18] = sum.ToString();
                hardSave[19] = splitTeacher[1][0] + "."+splitTeacher[2][0]+"."+splitTeacher[0];
                
                return hardSave;
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            return hardSave;
        }
        public string[] NabDataAkt(string[] heaher, string[] archpass, string[] hardSave, double sum)
        {
            string passport = heaher[8];
            string dat = "";
            try
            {
                for (int i = passport.Length - 10; i < passport.Length; i++)
                {
                    dat += heaher[8][i];
                }
                string[] splitTeacher = new string[3];
                splitTeacher[0] = heaher[2];
                splitTeacher[1] = heaher[3];
                splitTeacher[2] = heaher[4];
                archpass[2] = archpass[2].Replace(dat, "");
                string[] fdr = heaher[5].Split();
                hardSave[0] = (heaher[2] + " " + heaher[3] + " " + heaher[4]);
                hardSave[1] = fdr[0];
                hardSave[2] = heaher[7] + " " + heaher[6];
                hardSave[3] = archpass[0];
                hardSave[4] = archpass[1];
                hardSave[5] = archpass[2];
                hardSave[6] = dat;
                hardSave[7] = heaher[13] + " " + heaher[10] + " " + heaher[11] + " " + heaher[12] + " " + heaher[14];
                hardSave[8] = heaher[9];
                hardSave[9] = heaher[19];
                hardSave[10] = Days(Convert.ToInt32(DateTime.Now.ToString("dd")));
                hardSave[11] = Mouth(Convert.ToInt32(DateTime.Now.ToString("MM")));
                hardSave[12] = hardSave[0];
                hardSave[13] = archpass[0] + " " + archpass[1];
                hardSave[14] = archpass[2] + " " + dat;
                hardSave[15] = heaher[17];
                hardSave[16] = heaher[18];
                hardSave[17] = DateTime.Now.ToString("yyyy");
                hardSave[18] = sum.ToString();
                hardSave[19] = sum.ToString();
                hardSave[20] = Program.cent.Cafcombo;
                hardSave[21] = splitTeacher[1][0] + "." + splitTeacher[2][0] + "." + splitTeacher[3];
                return hardSave;
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            return hardSave;
        }
        public void ZNAGCreator(ExcelHelper helper)
        {
            Dictionary<string, string> Nag = new Dictionary<string, string> { };
            try
            {
                StringBuilder build = new StringBuilder();
                build.Append("SELECT * FROM ZNAG ORDER BY NUMB");
                string query = Convert.ToString(build);
                build.Clear();

                foreach (var nag in Z)
                {
                    Nag.Add("Лекции", nag.Lek.ToString());
                    Nag.Add("Лабораторнвые работы", nag.Lab.ToString());
                    Nag.Add("текущ.конс.", nag.TecCon.ToString());
                    Nag.Add("Зачёты", nag.Zach.ToString());
                    Nag.Add("Экзамены", nag.Exc.ToString());
                    Nag.Add("Консультации к экзаменам", nag.KonKExc.ToString());
                    Nag.Add("ДП руководство", nag.DPRuc.ToString());
                    Nag.Add("ДП предзащита", nag.DPPredZ.ToString());
                }

                SqlDataAdapter sql = new SqlDataAdapter(query, connection);
                System.Data.DataTable dtbl = new System.Data.DataTable();
                System.Data.DataTable dtbl1 = new System.Data.DataTable(); ;
                sql.Fill(dtbl);
                int NotZero = 0;
                foreach (var nag in Nag)
                {
                    if (Convert.ToDouble(nag.Value) > 0)
                    {
                        NotZero++;
                    }
                }
                //MessageBox.Show(NotZero.ToString());
                string[] sd = new string[NotZero];
                double[] gg = new double[NotZero];
                int schet = 0;
                foreach (var nag in Nag)
                {
                    if (Convert.ToDouble(nag.Value) > 0)
                    {
                        sd[schet] = nag.Key;
                        gg[schet] = Convert.ToDouble(nag.Value);
                        schet++;
                    }
                }

                int j = 0;
                for (int i = 0; i < NotZero; i++)
                {
                    helper.Set(row: Convert.ToInt32(dtbl.Rows[j][1]), data: sd[i], column: Convert.ToString(dtbl.Rows[j][0]));
                    helper.Set(row: Convert.ToInt32(dtbl.Rows[j + 1][1]), data: gg[i], column: Convert.ToString(dtbl.Rows[j + 1][0]));
                    j += 2;

                }

            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            finally 
            {
                Nag.Clear();
            }
        }
        
        public string[] NabDataZ(string[] heaher, string[] archpass, string[] hardSave, string Zav)
        {
            try
            {
                string passport = heaher[8];
                string dat = "";
                for (int i = passport.Length - 10; i < passport.Length; i++)
                {
                    dat += heaher[8][i];
                }
                double sum = 0;
                foreach (var nag in Z)
                {
                    sum += nag.Zach + nag.TecCon + nag.Lek + nag.Lab + nag.KonKExc + nag.Exc + nag.DPRuc + nag.DPPredZ;
                }
                string[] birthDay = heaher[5].Split(' ');
                archpass[2] = archpass[2].Replace(dat, "");
                hardSave[0] = (heaher[2] + " " + heaher[3] + " " + heaher[4]);
                hardSave[1] = heaher[18];
                hardSave[2] = birthDay[0];
                hardSave[3] = heaher[7];
                hardSave[4] = heaher[6];
                hardSave[5] = heaher[17];
                hardSave[6] = heaher[13] + " " + heaher[9] + " " + heaher[10] + " " + heaher[11] + " " + heaher[12] + " " + heaher[14];
                hardSave[7] = archpass[0];
                hardSave[8] = archpass[1];
                hardSave[9] = archpass[2];
                hardSave[10] = heaher[9];
                hardSave[11] = MouthE(Convert.ToInt32(DateTime.Now.ToString("MM")));
                hardSave[12] = Program.cent.Cafcombo;
                hardSave[13] = Program.cent.Faccom1;
                hardSave[14] = DateTime.Now.ToString("yyyy") + " г.";
                hardSave[15] = Days(Convert.ToInt32(DateTime.Now.ToString("MM")));
                hardSave[16] = Mouth(Convert.ToInt32(DateTime.Now.ToString("MM")));

                hardSave[17] = MouthZ(Convert.ToInt32(DateTime.Now.ToString("MM")));
                hardSave[18] = sum.ToString();
                
                return hardSave;
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            return hardSave;
        }
        
        public void ExeclDocFullerPOV(string path)
        {
            //try
            //{
            using (ExcelHelper helper = new ExcelHelper())
            {
                connection.Open();
                if (helper.Open(filePath: path))
                {
                    //try
                    //{
                    if (Program.cent.Check == true && (Program.cent.Combo7 == 1 || Program.cent.Combo9 == 1 || Program.cent.Combo7 == 0 || Program.cent.Combo9 == 0))
                    {
                        NagFReader();
                        ZeroZNPO(helper);
                        OneZNPO(helper);
                    }
                    else if (Program.cent.Combo7 == 0 || Program.cent.Combo9 == 0)
                    {
                        NagFReader();
                        ZeroZNPO(helper);
                    }
                    else if (Program.cent.Combo7 == 1 || Program.cent.Combo9 == 1)
                    {
                        NagFReader();
                        OneZNPO(helper);
                    }


                    FileInfo info = new FileInfo(path);
                    string direct = info.Directory.ToString();
                    string name = info.Name.ToString();
                    Object newFileName = Path.Combine(info.Directory.ToString(), DateTime.Now.ToString("yyyyMMdd") + info.Name);


                    //MessageBox.Show("1");
                    helper.Save(newFileName.ToString());
                    connection.Close();
                    Process.Start(newFileName.ToString());
                    
//}
                    //catch (Exception ex) {  MessageBox.Show(ex.Message); }
                }
            }
            
            //}
            //catch(Exception ex) { MessageBox.Show(ex.Message); }
            //finally {  }
        }
        
        public void ZeroZNPO(ExcelHelper helper)
        {
            string str;
            string[] splitTeacher;
            string Query = @"SELECT * FROM ZNPOV";
            StringBuilder build = new StringBuilder();
            str = Program.cent.FIOcombo;
            splitTeacher = str.Split();
            build.Append("SELECT * FROM Teacher WHERE T_Family = '" + splitTeacher[0] + "'" + "AND T_FirstName=" + "'" + splitTeacher[1] + "'" + "AND T_SecondNamme = " + "'" + splitTeacher[2] + "'");
            string query = Convert.ToString(build);
            build.Clear();
            build.Append("SELECT * FROM Cafedra WHERE name = '"); build.Append(Program.cent.Cafcombo); build.Append("'");
            string query1 = Convert.ToString(build);
            build.Clear();
            //try
            //{
                //connection.Open();
                SqlDataAdapter sql = new SqlDataAdapter(query, connection);
                System.Data.DataTable dtbl = new System.Data.DataTable();
                System.Data.DataTable dtbl1 = new System.Data.DataTable(); ;
                sql.Fill(dtbl);

                build.Append("SELECT Family, apellation, Sec_Name FROM Cafedra WHERE name = '" + Program.cent.Cafcombo + "'"); //build.Append(Program.cent.Cafcombo); build.Append("'");
                string query3 = Convert.ToString(build);
                SqlDataAdapter sql3 = new SqlDataAdapter(query3, connection);
                System.Data.DataTable dtbl2 = new System.Data.DataTable(); ;
                sql3.Fill(dtbl2);
                StringBuilder ZavCaf = new StringBuilder();
                for (int i = 0; i < dtbl2.Columns.Count; i++)
                {
                    ZavCaf.Append(dtbl2.Rows[0][i] + " ");
                }
                string Zav = ZavCaf.ToString();

                string[] heaher = new string[dtbl.Columns.Count];
                for (int i = 0; i < dtbl.Columns.Count; i++)
                {
                    heaher[i] = dtbl.Rows[0][i].ToString();

                }
                string[] archpass = new string[3];
                archpass = PassCreator(heaher[8]);
                string[] fdr = heaher[5].Split();
                SqlDataAdapter sqlQuery = new SqlDataAdapter(Query, connection);
                System.Data.DataTable dtblCod = new System.Data.DataTable();
                sqlQuery.Fill(dtblCod);
                for (int i = 0; i < dtblCod.Rows.Count; i++)
                {

                    dtblCod.Rows[i][0] = dtblCod.Rows[i][0].ToString().Replace(" ", "");
                }
                string[] hardSave = new string[dtblCod.Rows.Count];
                string[] dipData = heaher[16].Split(' ');
                hardSave[0] = splitTeacher[0] + " " + splitTeacher[1][0] + " " + splitTeacher[2][0];
                hardSave[1] = heaher[17];
                hardSave[2] = (Program.cent.Cafcombo) + " " + heaher[18];
                hardSave[3] = (Program.cent.Cafcombo);
                hardSave[4] = Days(Convert.ToInt32(DateTime.Now.ToString("MM")));
                hardSave[5] = Mouth(Convert.ToInt32(DateTime.Now.ToString("MM")));
                hardSave[6] = DateTime.Now.ToString("yyyy") + " г";
                hardSave[7] = heaher[2];
                hardSave[8] = heaher[3];
                hardSave[9] = heaher[4];
                hardSave[10] = fdr[0];
                hardSave[11] = heaher[21];
                hardSave[12] = dipData[0] + " " + heaher[20];
                hardSave[13] = splitTeacher[0] + " " + splitTeacher[1][0] + " " + splitTeacher[2][0];//FIO
                hardSave[14] = heaher[6];
                hardSave[15] = heaher[7];
                hardSave[16] = archpass[0] + " " + archpass[1];
                hardSave[17] = archpass[2];
                hardSave[18] = heaher[9];
                hardSave[19] = heaher[13] + " " + heaher[10] + " " + heaher[11] + " " + heaher[12] + " " + heaher[14];
                hardSave[20] = Days(Convert.ToInt32(DateTime.Now.ToString("MM")));
                hardSave[21] = Mouth(Convert.ToInt32(DateTime.Now.ToString("MM")));
                hardSave[22] = DateTime.Now.ToString("yyyy") + " г";
                hardSave[23] = hardSave[0];
                hardSave[24] = hardSave[3];
                hardSave[25] = (g[12] + g2[12]).ToString();//NAGRUZKA
                hardSave[26] = hardSave[25];
                hardSave[27] = Zav;
                hardSave[28] = hardSave[25];
                hardSave[29] = hardSave[25];

            for (int i = 0; i < dtblCod.Rows.Count; i++)
                {
                    //MessageBox.Show((dtblCod.Rows[i][0]+ dtblCod.Rows[i][1].ToString() + " " + hardSave[i] + " "+ i));
                    helper.Set(row: Convert.ToInt32(dtblCod.Rows[i][1]), data: hardSave[i], column: Convert.ToString(dtblCod.Rows[i][0]));

                }
                ZNPOSB(helper);
            //}
            //catch (Exception ex) { MessageBox.Show(ex.Message); }
            //finally {
                connection.Close(); //}
            
            
        }
        public void OneZNPO(ExcelHelper helper)
        {
            string str;
            string[] splitTeacher;
            string Query = @"SELECT * FROM ZNPOS";
            StringBuilder build = new StringBuilder();
            str = Program.cent.FIOcombo;
            splitTeacher = str.Split();
            build.Append("SELECT * FROM Teacher WHERE T_Family = '" + splitTeacher[0] + "'" + "AND T_FirstName=" + "'" + splitTeacher[1] + "'" + "AND T_SecondNamme = " + "'" + splitTeacher[2] + "'");
            string query = Convert.ToString(build);
            build.Clear();
            build.Append("SELECT * FROM Cafedra WHERE name = '"); build.Append(Program.cent.Cafcombo); build.Append("'");
            string query1 = Convert.ToString(build);
            build.Clear();
            //connection.Open();
            //try
            //{
                SqlDataAdapter sql = new SqlDataAdapter(query, connection);
                System.Data.DataTable dtbl = new System.Data.DataTable();
                System.Data.DataTable dtbl1 = new System.Data.DataTable(); ;
                sql.Fill(dtbl);

                build.Append("SELECT Family, apellation, Sec_Name FROM Cafedra WHERE name = '" + Program.cent.Cafcombo + "'"); //build.Append(Program.cent.Cafcombo); build.Append("'");
                string query3 = Convert.ToString(build);
                SqlDataAdapter sql3 = new SqlDataAdapter(query3, connection);
                System.Data.DataTable dtbl2 = new System.Data.DataTable(); ;
                sql3.Fill(dtbl2);
                StringBuilder ZavCaf = new StringBuilder();
                for (int i = 0; i < dtbl2.Columns.Count; i++)
                {
                    ZavCaf.Append(dtbl2.Rows[0][i] + " ");
                }
                string Zav = ZavCaf.ToString();

                string[] heaher = new string[dtbl.Columns.Count];
                for (int i = 0; i < dtbl.Columns.Count; i++)
                {
                    heaher[i] = dtbl.Rows[0][i].ToString();

                }
                string[] archpass = new string[3];
                archpass = PassCreator(heaher[8]);
                string[] fdr = heaher[5].Split();
                SqlDataAdapter sqlQuery = new SqlDataAdapter(Query, connection);
                System.Data.DataTable dtblCod = new System.Data.DataTable();
                sqlQuery.Fill(dtblCod);
                for (int i = 0; i < dtblCod.Rows.Count; i++)
                {

                    dtblCod.Rows[i][0] = dtblCod.Rows[i][0].ToString().Replace(" ", "");
                }
                string[] hardSave = new string[dtblCod.Rows.Count];
                string[] dipData = heaher[16].Split(' ');
                hardSave[0] = splitTeacher[0] + " " + splitTeacher[1][0] + " " + splitTeacher[2][0];
                hardSave[1] = heaher[17];
                hardSave[2] = (Program.cent.Cafcombo) + " " + heaher[18];
                hardSave[3] = (Program.cent.Cafcombo);
                hardSave[4] = Days(Convert.ToInt32(DateTime.Now.ToString("MM")));
                hardSave[5] = Mouth(Convert.ToInt32(DateTime.Now.ToString("MM")));
                hardSave[6] = DateTime.Now.ToString("yyyy") + " г";
                hardSave[7] = heaher[2];
                hardSave[8] = heaher[3];
                hardSave[9] = heaher[4];
                hardSave[10] = fdr[0];
                hardSave[11] = heaher[21];
                hardSave[12] = dipData[0] + " " + heaher[20];
                hardSave[13] = splitTeacher[0] + " " + splitTeacher[1][0] + " " + splitTeacher[2][0];//FIO
                hardSave[14] = heaher[6];
                hardSave[15] = heaher[7];
                hardSave[16] = archpass[0] + " " + archpass[1];
                hardSave[17] = archpass[2];
                hardSave[18] = heaher[9];
                hardSave[19] = heaher[13] + " " + heaher[10] + " " + heaher[11] + " " + heaher[12] + " " + heaher[14];
                hardSave[20] = Days(Convert.ToInt32(DateTime.Now.ToString("MM")));
                hardSave[21] = Mouth(Convert.ToInt32(DateTime.Now.ToString("MM")));
                hardSave[22] = DateTime.Now.ToString("yyyy") + " г";
                hardSave[23] = hardSave[0];
                hardSave[24] = hardSave[3];
                hardSave[25] = (g1[12] + g3[12]).ToString();//NAGRUZKA
                hardSave[26] = hardSave[25];
                hardSave[27] = Zav;
                hardSave[28] = hardSave[25];
                hardSave[29] = hardSave[25];
                for (int i = 0; i < dtblCod.Rows.Count; i++)
                {
                    //MessageBox.Show((dtblCod.Rows[i][0]+ dtblCod.Rows[i][1].ToString() + " " + hardSave[i]));
                    helper.Set(row: Convert.ToInt32(dtblCod.Rows[i][1]), data: hardSave[i], column: Convert.ToString(dtblCod.Rows[i][0]));

                }
                ZNPOSV(helper);
                ZNPOSVD(helper);
            //}
            //catch (Exception ex) { MessageBox.Show(ex.Message); }
            //inally {
            connection.Close(); //}
            
        }
        public void ZNPOSB(ExcelHelper helper)
        {
            try
            {
                string Query = "select * from VOB";
                SqlDataAdapter sqlQuery = new SqlDataAdapter(Query, connection);
                System.Data.DataTable dtblCod = new System.Data.DataTable();
                sqlQuery.Fill(dtblCod);
                for (int i = 0; i < dtblCod.Rows.Count; i++)
                {
                    dtblCod.Rows[i][0] = dtblCod.Rows[i][0].ToString().Replace(" ", "");
                }
                for (int i = 0; i < dtblCod.Rows.Count; i++)
                {
                    helper.Set(row: Convert.ToInt32(dtblCod.Rows[i][1]), data: g[i], column: Convert.ToString(dtblCod.Rows[i][0]));

                }
                string Query1 = "select * from VSN";
                SqlDataAdapter sqlQuery1 = new SqlDataAdapter(Query1, connection);
                System.Data.DataTable dtblCod1 = new System.Data.DataTable();
                sqlQuery1.Fill(dtblCod1);
                for (int i = 0; i < dtblCod1.Rows.Count; i++)
                {
                    dtblCod1.Rows[i][0] = dtblCod1.Rows[i][0].ToString().Replace(" ", "");
                }
                for (int i = 0; i < dtblCod1.Rows.Count; i++)
                {
                    helper.Set(row: Convert.ToInt32(dtblCod1.Rows[i][1]), data: g2[i], column: Convert.ToString(dtblCod1.Rows[i][0]));


                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }
        public void ZNPOSVD(ExcelHelper helper)
        {
            //try
            //{
            string Query1 = "select * from VOV";

            SqlDataAdapter sqlQuery1 = new SqlDataAdapter(Query1, connection);

            System.Data.DataTable dtblCod1 = new System.Data.DataTable();

            sqlQuery1.Fill(dtblCod1);

            for (int i = 0; i < dtblCod1.Rows.Count; i++)
            {
                dtblCod1.Rows[i][0] = dtblCod1.Rows[i][0].ToString().Replace(" ", "");
            }
            
            for (int i = dtblCod1.Rows.Count; i < dtblCod1.Rows.Count; i++)
            {
                
                helper.Set(row: Convert.ToInt32(dtblCod1.Rows[i][1]), data: g1[i], column: Convert.ToString(dtblCod1.Rows[i][0]));

            }
            //}
            //catch (Exception ex) { MessageBox.Show(ex.Message); }
        }
        public void ZNPOSV(ExcelHelper helper)
        {
            //try
            //{
            string Query1 = "select * from VOV";
            string Query = "select * from VSNV";
                SqlDataAdapter sqlQuery = new SqlDataAdapter(Query, connection);
                System.Data.DataTable dtblCod = new System.Data.DataTable();
                sqlQuery.Fill(dtblCod);
            SqlDataAdapter sqlQuery1 = new SqlDataAdapter(Query1, connection);

            System.Data.DataTable dtblCod1 = new System.Data.DataTable();

            sqlQuery1.Fill(dtblCod1);

            for (int i = 0; i < dtblCod.Rows.Count; i++)
                {
                    dtblCod.Rows[i][0] = dtblCod.Rows[i][0].ToString().Replace(" ", "");
                dtblCod1.Rows[i][0] = dtblCod1.Rows[i][0].ToString().Replace(" ", "");
            }
                for (int i = 0; i < dtblCod.Rows.Count; i++)
                {
                    helper.Set(row: Convert.ToInt32(dtblCod.Rows[i][1]), data: g3[i], column: Convert.ToString(dtblCod.Rows[i][0]));
                helper.Set(row: Convert.ToInt32(dtblCod1.Rows[i][1]), data: g1[i], column: Convert.ToString(dtblCod1.Rows[i][0]));
            }
            //}
            //catch (Exception ex) { MessageBox.Show(ex.Message); }
        }
        public void NagFReader()
        {

            for( int i = 0; i< 13; i++)
            {
                g[i] = 0;
                g1[i] = 0;
                g2[i] = 0;
                g3[i] = 0;
            }
            var dialog = new OpenFileDialog
            {
                FileName = "Document", // Default file name
                DefaultExt = ".xlsx", // Default file extension
                Filter = "Excel (.xlsx)|*.xlsx" // Filter files by extension
            };

            // Show open file dialog box
            bool? result = Convert.ToBoolean(dialog.ShowDialog());
            string path = "";
            // Process open file dialog box results
            
            if (result == true)
            {
                
                // Open document
                path = dialog.FileName;
                var excelfile = new Application();
                Workbook workbook = excelfile.Workbooks.Open(path);
                Worksheet worksheet = workbook.Worksheets[1];

                
                

                using (ExcelHelper helper = new ExcelHelper())
                {

                    if (helper.Open(filePath: path))
                    {


                        Excel.Range Rng; //диапазон ячеек
                        string textToFind = "СПО, ФиПМ, Почасовая"; //текст для поиска
                        Rng = worksheet.Cells.Find(textToFind, Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart); //осуществляем поиск на листе
                        string h = Rng.Address;
                        string stolb = "A";
                        int rangeEnd = 0;
                        int yach = Convert.ToInt32(Rng.Row.ToString());
                        bool flg = true;

                        //int bjudget = 0, vnebjudget = 0, osn = 0, vsn = 0, osnB = 0, vsnB = 0;
                        List<Nag> NagSaver = new List<Nag>();
                        
                        for (int i = yach; flg; i++)
                        {
                            h = Convert.ToString(helper.Get(stolb, i)).Trim();
                            if (textToFind.Equals(h) == false)
                            {
                                flg = false;
                                rangeEnd = i;

                            }
                        }
                        
                        if (helper.Open(filePath: path))
                        {
                                
                                    NagSaver = new List<Nag>();
                                    for (int i = yach + 1; i <= rangeEnd; i++)
                                    {
                                        NagSaver.Add(new Nag
                                        {
                                            Deyat = Convert.ToString(helper.Get(column: "A", row: i)),
                                            Disp = Convert.ToString(helper.Get(column: "C", row: i)),
                                            Sem = Convert.ToString(helper.Get(column: "D", row: i)),
                                            Categ = Convert.ToString(helper.Get(column: "E", row: i)),
                                            Itogo = Convert.ToDouble(helper.Get(column: "H", row: i)),
                                            NumOfStud = Convert.ToInt32(helper.Get(column: "J", row: i)),

                                            Lek = Convert.ToDouble(helper.Get(column: "M", row: i)),
                                            Lab = Convert.ToDouble(helper.Get(column: "N", row: i)),
                                            TecCons = Convert.ToDouble(helper.Get(column: "O", row: i)),
                                            Zach = Convert.ToDouble(helper.Get(column: "P", row: i)),
                                            Exz = Convert.ToDouble(helper.Get(column: "Q", row: i)),
                                            KonKEkz = Convert.ToDouble(helper.Get(column: "R", row: i)),
                                            RukDP = Convert.ToDouble(helper.Get(column: "S", row: i)),
                                            DPPredZach = Convert.ToDouble(helper.Get(column: "T", row: i)),


                                            //PraktZn = Convert.ToDouble(helper.Get(column: "U", row: i)),
                                            //KursPr = Convert.ToDouble(helper.Get(column: "V", row: i)),
                                            //PrPrak = Convert.ToDouble(helper.Get(column: "W", row: i)),
                                            //ExzPoMod = Convert.ToDouble(helper.Get(column: "X", row: i)),
                                            //Gek = Convert.ToDouble(helper.Get(column: "Y", row: i)),

                                        });
                                    }
                                
                                helper.Save();
                            
                        }
                            
                                List<Nag> OsnB = new List<Nag>();
                                List<Nag> OsnV = new List<Nag>();
                                List<Nag> VsnB = new List<Nag>();
                                List<Nag> VsnV = new List<Nag>();
                                foreach (var d in NagSaver)
                                {
                                    if (d.Sem.Equals("Осн") && d.Categ.Equals("Б"))
                                    {
                                        OsnB.Add(d);
                                    }
                                    else if (d.Sem.Equals("Всн") && d.Categ.Equals("Б"))
                                    {
                                        VsnB.Add(d);
                                    }
                                    else if (d.Sem.Equals("Осн") && d.Categ.Equals("ВНБ"))
                                    {
                                        OsnV.Add(d);
                                    }
                                    else if (d.Sem.Equals("Всн") && d.Categ.Equals("ВНБ"))
                                    {
                                        VsnV.Add(d);
                                    }
                                }

                                foreach (var d in OsnB)
                                {
                                    g[0] = d.Lek + g[0];
                                    g[1] = d.PraktZn + g[1];
                                    g[2] = g[2] + d.Lab;
                                    g[3] = g[3] + d.KursPr;//
                                    g[4] = g[4] + d.Zach;
                                    g[5] = g[5] + d.Exz;
                                    g[6] = g[6] + d.KonKEkz;
                                    g[7] = g[7] + d.Gek;
                                    g[8] = g[8] + d.RukDP;
                                    g[9] = g[9] + d.DPPredZach;
                                    g[10] = g[10] + d.PrPrak;
                                    g[11] = g[11] + d.ExzPoMod + d.TecCons;
                                }

                                foreach (var d in OsnV)
                                {
                                    g1[0] = d.Lek + g1[0];
                                    g1[1] = d.PraktZn + g1[1];
                                    g1[2] = g1[2] + d.Lab;
                                    g1[3] = g1[3] + d.KursPr;//
                                    g1[4] = g1[4] + d.Zach;
                                    g1[5] = g[5] + d.Exz;
                                    g1[6] = g1[6] + d.KonKEkz;
                                    g1[7] = g1[7] + d.Gek;
                                    g1[8] = g1[8] + d.RukDP;
                                    g1[9] = g1[9] + d.DPPredZach;
                                    g1[10] = g1[10] + d.PrPrak;
                                    g1[11] = g1[11] + d.ExzPoMod + d.TecCons;
                                }
                                foreach (var d in VsnB)
                                {
                                    g2[0] = d.Lek + g2[0];
                                    g2[1] = d.PraktZn + g2[1];
                                    g2[2] = g2[2] + d.Lab;
                                    g2[3] = g2[3] + d.KursPr;//
                                    g2[4] = g2[4] + d.Zach;
                                    g2[5] = g2[5] + d.Exz;
                                    g2[6] = g2[6] + d.KonKEkz;
                                    g2[7] = g2[7] + d.Gek;
                                    g2[8] = g2[8] + d.RukDP;
                                    g2[9] = g2[9] + d.DPPredZach;
                                    g2[10] = g2[10] + d.PrPrak;
                                    g2[11] = g2[11] + d.ExzPoMod + d.TecCons;
                                }
                                foreach (var d in VsnV)
                                {
                                    g3[0] = d.Lek + g3[0];
                                    g3[1] = d.PraktZn + g3[1];
                                    g3[2] = g3[2] + d.Lab;
                                    g3[3] = g3[3] + d.KursPr;//
                                    g3[4] = g3[4] + d.Zach;
                                    g3[5] = g3[5] + d.Exz;
                                    g3[6] = g3[6] + d.KonKEkz;
                                    g3[7] = g3[7] + d.Gek;
                                    g3[8] = g3[8] + d.RukDP;
                                    g3[9] = g3[9] + d.DPPredZach;
                                    g3[10] = g3[10] + d.PrPrak;
                                    g3[11] = g3[11] + d.ExzPoMod + d.TecCons;
                                }

                                for (int i = 0; i < g.Length - 1; i++)
                                {

                                    g[12] += g[i];//MessageBox.Show(g[12].ToString());

                                    g1[12] += g1[i];//MessageBox.Show(g1[12].ToString());

                                    g2[12] += g2[i];//MessageBox.Show(g2[12].ToString());

                                    g3[12] += g3[i];//MessageBox.Show(g3[12].ToString());

                                }
                            }
                            
                            helper.Save();
                        
                        if(g[g.Length-1]+g2[g.Length - 1] > 300 || g1[g.Length - 1] + g3[g.Length - 1] >300)
                        {
                            flag = true;
                        }
                    }
                
                Console.Read();  
            }
            
            
        }
        public void NagDPFReader()
        {
            var dialog = new OpenFileDialog
            {
                FileName = "Document", // Default file name
                DefaultExt = ".xlsx", // Default file extension
                Filter = "Excel (.xlsx)|*.xlsx" // Filter files by extension
            };

            // Show open file dialog box
            bool? result = Convert.ToBoolean(dialog.ShowDialog());
            string path = "";
            // Process open file dialog box results
            if (result == true)
            {
                //try
                //{
                // Open document
                path = dialog.FileName;
                var excelfile = new Application();
                Workbook workbook = excelfile.Workbooks.Open(path);
                Worksheet worksheet = workbook.Worksheets[1];

                //}
                //catch (Exception ex) { MessageBox.Show(ex.Message); }
                //try
                //{

                using (ExcelHelper helper = new ExcelHelper())
                {

                    if (helper.Open(filePath: path))
                    {


                        Excel.Range Rng; //диапазон ячеек
                        string textToFind = "СПО, ФиПМ, Почасовая"; //текст для поиска
                        Rng = worksheet.Cells.Find(textToFind, Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart); //осуществляем поиск на листе
                        string h = Rng.Address;
                        string stolb = "A";
                        int rangeEnd = 0;
                        int yach = Convert.ToInt32(Rng.Row.ToString());
                        bool flg = true;

                        //int bjudget = 0, vnebjudget = 0, osn = 0, vsn = 0, osnB = 0, vsnB = 0;
                        List<DPNag> NagSaver = new List<DPNag>();

                        for (int i = yach; flg; i++)
                        {
                            h = Convert.ToString(helper.Get(stolb, i)).Trim();
                            if (textToFind.Equals(h) == false)
                            {
                                flg = false;
                                rangeEnd = i;

                            }
                        }
                        //MessageBox.Show("1");
                        Excel.Range Rng1; //диапазон ячеек
                        string textToFind1 = "ФиПМ, Почасовая"; //текст для поиска
                        Rng1 = worksheet.Cells.Find(textToFind1, Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart); //осуществляем поиск на листе
                        string h1 = Rng1.Address;
                        
                        int rangeEnd1 = 0;
                        int yach1 = Convert.ToInt32(Rng1.Row.ToString());
                        bool flg1 = true;

                        //int bjudget = 0, vnebjudget = 0, osn = 0, vsn = 0, osnB = 0, vsnB = 0;
                        //MessageBox.Show("2");

                        for (int i = rangeEnd; flg1; i++)
                        {
                            h1 = Convert.ToString(helper.Get(stolb, i)).Trim();
                            if (textToFind1.Equals(h1) == false)
                            {
                                flg1 = false;
                                rangeEnd1 = i;

                            }
                        }
                        //MessageBox.Show(rangeEnd1.ToString() + " " + yach1);
                        //MessageBox.Show("3");

                        if (helper.Open(filePath: path))
                        {
                                try
                                {
                                    NagSaver = new List<DPNag>();
                                    if(Program.cent.Combo4 == 0 || Program.cent.Combo4 == 1 || Program.cent.Combo6 == 0 || Program.cent.Combo6 == 1)
                                    {
                                        for (int i = yach; i <= rangeEnd; i++)
                                        {
                                            NagSaver.Add(new DPNag
                                            {

                                                Disp = Convert.ToString(helper.Get(column: "B", row: i)),
                                                Sem = Convert.ToString(helper.Get(column: "C", row: i)),
                                                Categ = Convert.ToString(helper.Get(column: "D", row: i)),
                                                Itogo = Convert.ToDouble(helper.Get(column: "H", row: i)),
                                                NumOfStud = Convert.ToInt32(helper.Get(column: "I", row: i)),
                                                Group = Convert.ToString(helper.Get(column: "F", row: i)),
                                                GAK = Convert.ToDouble(helper.Get(column: "K", row: i)),
                                                Spec = Convert.ToString(helper.Get(column: "G", row: i)),

                                            }); ;
                                        }
                                    }
                                    if (Program.cent.Combo5 == 0 || Program.cent.Combo5 == 1 || Program.cent.Combo3 == 0 || Program.cent.Combo3 == 1)
                                    {
                                        for (int i = rangeEnd; i <= rangeEnd1; i++)
                                        {
                                            NagSaver.Add(new DPNag
                                            {

                                                Disp = Convert.ToString(helper.Get(column: "B", row: i)),
                                                Sem = Convert.ToString(helper.Get(column: "C", row: i)),
                                                Categ = Convert.ToString(helper.Get(column: "D", row: i)),
                                                Itogo = Convert.ToDouble(helper.Get(column: "H", row: i)),
                                                NumOfStud = Convert.ToInt32(helper.Get(column: "I", row: i)),
                                                Group = Convert.ToString(helper.Get(column: "F", row: i)),
                                                GAK = Convert.ToDouble(helper.Get(column: "K", row: i)),
                                                Spec = Convert.ToString(helper.Get(column: "G", row: i)),

                                            }); ;
                                        }
                                    }
                                }
                                catch (Exception ex) { MessageBox.Show(ex.Message); }
                                finally { helper.Save(); }
                            
                        }
                        string[] group;
                        foreach(var nag in NagSaver)
                            {
                                group = nag.Group.Split(',');
                                if(group.Length > 1)
                                {
                                    DP.Add(new DPNag
                                    {

                                        Disp = nag.Disp,
                                        Sem = nag.Sem,
                                        Categ = nag.Categ,
                                        Itogo = nag.Itogo,
                                        NumOfStud = nag.NumOfStud,
                                        Group = group[0],
                                        GAK = nag.GAK,
                                        Spec = nag.Spec,

                                    });
                                    DP.Add(new DPNag
                                    {

                                        Disp = nag.Disp,
                                        Sem = nag.Sem,
                                        Categ = nag.Categ,
                                        Itogo = nag.Itogo,
                                        NumOfStud = nag.NumOfStud,
                                        Group = group[1],
                                        GAK = nag.GAK,
                                        Spec = nag.Spec,

                                    });
                                    //NagSaver.Remove(nag);
                                }
                            }
                        
                        foreach (var nag in NagSaver)
                        {
                            //MessageBox.Show(nag.Disp + " " + nag.NumOfStud + " " + nag.GAK + " " + nag.Itogo);
                            //if (nag.Categ.Equals('Б') || nag.Categ.Equals("ВНБ"))
                            group = nag.Group.Split(',');
                            
                            if (nag.NumOfStud>0 && group.Length < 2)
                            {
                                DP.Add(nag);
                            }
                        }
                        
                        helper.Save();
                    }
                }
                Console.Read();
                //}
                //catch (Exception ex) { MessageBox.Show(ex.Message); }
            }
            //try
            //{

        }
        public void NagZFReader()
        {
            var dialog = new OpenFileDialog
            {
                FileName = "Document", // Default file name
                DefaultExt = ".xlsx", // Default file extension
                Filter = "Excel (.xlsx)|*.xlsx" // Filter files by extension
            };

            // Show open file dialog box
            bool? result = Convert.ToBoolean(dialog.ShowDialog());
            string path = "";
            // Process open file dialog box results
            if (result == true)
            {
                //try
                //{
                // Open document
                path = dialog.FileName;
                var excelfile = new Application();
                Workbook workbook = excelfile.Workbooks.Open(path);
                Worksheet worksheet = workbook.Worksheets[1];

                //}
                //catch (Exception ex) { MessageBox.Show(ex.Message); }
                try
                {

                using (ExcelHelper helper = new ExcelHelper())
                {

                    if (helper.Open(filePath: path))
                    {


                        Excel.Range Rng; //диапазон ячеек

                        string textToFind = MouthZ(Convert.ToInt32(DateTime.Now.ToString("MM"))); //текст для поиска
                        
                        Rng = worksheet.Cells.Find(textToFind, Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart); //осуществляем поиск на листе
                        string h = Rng.Address;
                        string stolb = "W";
                        int rangeEnd = 0;
                        int yach = Convert.ToInt32(Rng.Row.ToString());
                        bool flg = true;

                        //int bjudget = 0, vnebjudget = 0, osn = 0, vsn = 0, osnB = 0, vsnB = 0;
                        List<ZNag> NagSaver = new List<ZNag>();

                        for (int i = yach; flg; i++)
                        {
                            h = Convert.ToString(helper.Get(stolb, i)).Trim();
                            if (textToFind.Equals(h) == false)
                            {
                                flg = false;
                                rangeEnd = i;

                            }
                        }
                            try
                            {
                                if (helper.Open(filePath: path))
                                {
                                    NagSaver = new List<ZNag>();

                                    NagSaver.Add(new ZNag
                                    {

                                        Lek = Convert.ToDouble(helper.Get(column: "X", row: yach)),
                                        Lab = Convert.ToDouble(helper.Get(column: "Y", row: yach)),
                                        TecCon = Convert.ToDouble(helper.Get(column: "Z", row: yach)),
                                        Zach = Convert.ToDouble(helper.Get(column: "AA", row: yach)),
                                        Exc = Convert.ToDouble(helper.Get(column: "AB", row: yach)),

                                        KonKExc = Convert.ToDouble(helper.Get(column: "AC", row: yach)),
                                        DPRuc = Convert.ToDouble(helper.Get(column: "AD", row: yach)),
                                        DPPredZ = Convert.ToDouble(helper.Get(column: "AE", row: yach)),

                                    });

                                    helper.Save();
                                }
                                foreach (var nag in NagSaver)
                                {
                                    Z.Add(nag);

                                }
                            }
                            catch (Exception ex) { MessageBox.Show(ex.Message); }
                            finally {helper.Save(); }
                            double sum = 0;
                        foreach(var nag in NagSaver)
                            {
                                sum = nag.DPPredZ + nag.DPRuc + nag.Exc + nag.KonKExc + nag.Lab + nag.Lek + nag.TecCon + nag.Zach;
                            }
                            if (sum > 100)
                            {
                                flagZ = true;
                            }
                    }
                }
                Console.Read();
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
            }
            //try
            //{

        }
        public teacherForm teacherForm
        {
            get => default;
            set
            {
            }
        }
        public void AddData()
        {
            throw new System.NotImplementedException();
        }

        public void DelData()
        {
            throw new System.NotImplementedException();
        }

        public void ChangeData()
        {
            throw new System.NotImplementedException();
        }

        public Central Form1
        {
            get => default;
            set
            {
            }
        }
    }
    
}
