using System;
using System.Collections.Generic;

using Application = Microsoft.Office.Interop.Excel.Application;

using Path = System.IO.Path;
using System.Data.SqlClient;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using Point = System.Drawing.Point;
using System.Runtime.InteropServices;
using System.IO;
using Microsoft.Office.Interop.Word;
using System.Text;

namespace NewVersionDip
{
    public partial class Central : Form
    {
        bool flag = false;
        WorkHelper work = new WorkHelper();
        private Point MouseHook;
        int lastUsedRow = 0;
        int lastUsedColumn = 0;
        bool flagAut = true;
        List<Dict> teachers;
       
        public static SqlConnection connection = new SqlConnection(Properties.Settings.Default.connString);
        OpenFileDialog openFileDialogRes = new OpenFileDialog();

        SaveFileDialog saveFileDialogBack = new SaveFileDialog();
        public static bool SQLStat = true;
        public Central()
        {
            Program.cent = this;
            InitializeComponent();
            this.FormBorderStyle = FormBorderStyle.None;
            work.comboBoxFuller();
            CafcomboBox.SelectedIndex = 0;
            CafcomboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            FIOcomboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            faccomboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox5.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox6.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox7.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox8.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox9.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            faccomboBox.SelectedItem = 0;
            FIOcomboBox.SelectedItem = 0;
            CafcomboBox.SelectedItem = 0;

        }
        public string Faccom1
        {
            get { return faccomboBox.SelectedItem.ToString(); }
            set { faccomboBox.Items.Add(value); }
        }
        public int Faccom
        {
            get { return faccomboBox.SelectedIndex; }
            set { faccomboBox.Items.Add(value); }
        }
        public string FIOcombo
        {
            get { return FIOcomboBox.SelectedItem.ToString(); }
            set { FIOcomboBox.Items.Add(Convert.ToString(value)); }
        }
        public string Cafcombo
        {
            get { return CafcomboBox.SelectedItem.ToString(); }
            set { CafcomboBox.Items.Add(value); }
        }

        public int Combo7 
        {
            get { return comboBox7.SelectedIndex; }
            set { } 
        }
        public int Combo1
        {
            get { return comboBox1.SelectedIndex; }
            set {  }
        }
        public int Combo2
        {
            get { return comboBox2.SelectedIndex; }
            set {  }
        }
        public int Combo5
        {
            get { return comboBox5.SelectedIndex; }
            set { }
        }
        public int Combo6
        {
            get { return comboBox6.SelectedIndex; }
            set {  }
        }
        public bool Check
        {
            get { return checkBox1.Checked; }
            set {  }
        }

        public int Combo4 
        {
            get { return comboBox4.SelectedIndex; }
            set {  }
        }
        public int Combo3 {
            get { return comboBox3.SelectedIndex; }
            set {  }
        }
        public int Combo9
        {
            get { return comboBox9.SelectedIndex; }
            set { }
        }
        public string GEKcombo 
        {
            get { return comboBox8.SelectedItem.ToString(); }
            set { comboBox8.Items.Add(Convert.ToString(value)); }
        }

        

        private void DocButtonFuller_Click(object sender, EventArgs e)
        {
            //Autentif aut = new Autentif();

            
            //if (aut.flag == true)
            //{
                work.choiceDockA4Word();
                comboBox1.SelectedItem = null;
                comboBox2.SelectedItem = null;
                comboBox3.SelectedItem = null;
                comboBox4.SelectedItem = null;
                comboBox5.SelectedItem = null;
                comboBox6.SelectedItem = null;
                comboBox7.SelectedItem = null;
                comboBox8.SelectedItem = null;
                comboBox9.SelectedItem = null;
                faccomboBox.SelectedItem = null;
                FIOcomboBox.SelectedItem = null;
                CafcomboBox.SelectedItem = null;
            work.Clearer();
            //}


        }

        private void Form1_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) MouseHook = e.Location;
            Location = new Point((Size)Location - (Size)MouseHook + (Size)e.Location);
            
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            Environment.Exit(0);

        }

        private void panel6_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) MouseHook = e.Location;
            Location = new Point((Size)Location - (Size)MouseHook + (Size)e.Location);
        }

        private void panel5_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) MouseHook = e.Location;
            Location = new Point((Size)Location - (Size)MouseHook + (Size)e.Location);
        }

        private void panel3_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) MouseHook = e.Location;
            Location = new Point((Size)Location - (Size)MouseHook + (Size)e.Location);
        }
        
        

        private void SVbutton_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void A5Word_Click(object sender, EventArgs e)
        {
            
            /*if(aut.flag == false)
            {
                aut.ShowDialog();
            }
            
            if (aut.flag == true)
            {*/
                work.choiceDockA5Word();
                comboBox1.SelectedItem = null;
                comboBox2.SelectedItem = null;
                comboBox3.SelectedItem = null;
                comboBox4.SelectedItem = null;
                comboBox5.SelectedItem = null;
                comboBox6.SelectedItem = null;
                comboBox7.SelectedItem = null;
                comboBox8.SelectedItem = null;
                comboBox9.SelectedItem = null;
                faccomboBox.SelectedItem = null;
                FIOcomboBox.SelectedItem = null;
                CafcomboBox.SelectedItem = null;
            work.Clearer();
            //}

        }

        private void A5Excel_Click(object sender, EventArgs e)
        {
            //Autentif aut = new Autentif();

            /*if (aut.flag == false)
            {
                aut.ShowDialog();
            }
            if (aut.flag == true)
            {*/
                work.choiceDockA5Excel();
                comboBox1.SelectedItem = null;
                comboBox2.SelectedItem = null;
                comboBox3.SelectedItem = null;
                comboBox4.SelectedItem = null;
                comboBox5.SelectedItem = null;
                comboBox6.SelectedItem = null;
                comboBox7.SelectedItem = null;
                comboBox8.SelectedItem = null;
                comboBox9.SelectedItem = null;
                faccomboBox.SelectedItem = null;
                FIOcomboBox.SelectedItem = null;
                CafcomboBox.SelectedItem = null;
            work.Clearer();
            //}

        }

        private void A4Excel_Click(object sender, EventArgs e)
        {
            //Autentif aut = new Autentif();

            /*if (flag == false)
            {
                aut.ShowDialog();
            }
            if (aut.flag == true)
            {
                flag = aut.flag;*/
                work.choiceDockA4Excel();
                comboBox1.SelectedItem = null;
                comboBox2.SelectedItem = null;
                comboBox3.SelectedItem = null;
                comboBox4.SelectedItem = null;
                comboBox5.SelectedItem = null;
                comboBox6.SelectedItem = null;
                comboBox7.SelectedItem = null;
                comboBox8.SelectedItem = null;
                comboBox9.SelectedItem = null;
                faccomboBox.SelectedItem = null;
                FIOcomboBox.SelectedItem = null;
                CafcomboBox.SelectedItem = null;
            work.Clearer();
            //}

        }

        

        private void изменитьToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            teacherForm teacher = new teacherForm();
            teacher.ShowDialog();
        }

        private void изменитьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            teacherForm teacher = new teacherForm();
            flag = true;
            teacher.ShowDialog();
        }

        private void добавитьToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            var dialog = new System.Windows.Forms.OpenFileDialog();
            dialog.FileName = "Document"; // Default file name
            dialog.DefaultExt = ".xlsx"; // Default file extension
            dialog.Filter = "Excel (.xlsx)|*.xlsx"; // Filter files by extension

            // Show open file dialog box
            bool? result = Convert.ToBoolean(dialog.ShowDialog());
            string path = "";
            // Process open file dialog box results
            if (result == true)
            {
                try
                {
                    // Open document
                    path = dialog.FileName;
                    var excelfile = new Application();
                    Workbook workbook = excelfile.Workbooks.Open(path);
                    Worksheet worksheet = workbook.Worksheets[1];
                    lastUsedRow = worksheet.Cells.Find("*", System.Reflection.Missing.Value,
                                       System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                       Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious,
                                       false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                    // Find the last real column
                    lastUsedColumn = worksheet.Cells.Find("*", System.Reflection.Missing.Value,
                                                   System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                                   Microsoft.Office.Interop.Excel.XlSearchOrder.xlByColumns, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious,
                                                   false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
                try
                {

                    using (ExcelHelper helper = new ExcelHelper())
                    {

                        if (helper.Open(filePath: path))
                        {
                            teachers = new List<Dict>();
                            for (int i = 1; i <= lastUsedRow; i++)
                            {
                                teachers.Add(new Dict
                                {
                                    Caf = Convert.ToString(helper.Get(column: "A", row: i)),
                                    Family = Convert.ToString(helper.Get(column: "B", row: i)),
                                    Name = Convert.ToString(helper.Get(column: "C", row: i)),
                                    SecName = Convert.ToString(helper.Get(column: "D", row: i)),
                                    Birth = DateTime.FromOADate(Convert.ToDouble(helper.Get(column: "E", row: i))),

                                    ScienceStep = Convert.ToString(helper.Get(column: "F", row: i)),
                                    ScienceZV = Convert.ToString(helper.Get(column: "G", row: i)),
                                    PassDate = Convert.ToString(helper.Get(column: "H", row: i)),
                                    SNLS = Convert.ToString(helper.Get(column: "I", row: i)),
                                    Toun = Convert.ToString(helper.Get(column: "J", row: i)),
                                    Strit = Convert.ToString(helper.Get(column: "K", row: i)),
                                    House = Convert.ToString(helper.Get(column: "L", row: i)),
                                    Flow = Convert.ToString(helper.Get(column: "M", row: i)),
                                    Index = Convert.ToString(helper.Get(column: "N", row: i)),
                                    Telephone = Convert.ToString(helper.Get(column: "O", row: i)),
                                    Diplom = DateTime.FromOADate(Convert.ToDouble(helper.Get(column: "P", row: i))),
                                    PlaceWorking = Convert.ToString(helper.Get(column: "Q", row: i)),
                                    dolgnost = Convert.ToString(helper.Get(column: "R", row: i)),
                                    INN = Convert.ToString(helper.Get(column: "S", row: i)),
                                    dipSerNum = Convert.ToString(helper.Get(column: "T", row: i)),
                                    PLB = Convert.ToString(helper.Get(column: "U", row: i))
                                });
                            }
                            //helper.Save();
                        }
                    }
                    Console.Read();
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
            }
            try
            {
                foreach (Dict teacher in teachers)
                {

                    using (SqlConnection conn = new SqlConnection(Properties.Settings.Default.connString))
                    {
                        string queryNaDob = @"SELECT * FROM Teacher WHERE T_Family =" + "'" + teacher.Family + "'" + " and " +
                            "T_FirstName =" + "'" + teacher.Name + "'" + " and " + "T_SecondNamme =" + "'" + teacher.SecName + "'";
                        conn.Open();
                        SqlDataAdapter sqlQuery = new SqlDataAdapter(queryNaDob, connection);
                        System.Data.DataTable dtblDob = new System.Data.DataTable();
                        sqlQuery.Fill(dtblDob);
                        if (!(dtblDob.Rows.Count == 1))
                        {
                            string query = "WITH SRC AS ( SELECT TOP(1) T_ID, T_Family FROM Teacher ORDER BY T_ID DESC ) SELECT* FROM SRC ORDER BY T_ID";
                            SqlDataAdapter sql = new SqlDataAdapter(query, connection);
                            System.Data.DataTable dtbl = new System.Data.DataTable();
                            sql.Fill(dtbl);

                            string[] heaher = new string[dtbl.Columns.Count];
                            for (int i = 0; i < dtbl.Columns.Count; i++)
                            {
                                heaher[i] = dtbl.Rows[0][i].ToString();
                            }
                            using (SqlCommand cmd = new SqlCommand("INSERT INTO Teacher(Caf_ID, T_Family, T_FirstName, T_SecondNamme, dateOfBirth, ScienceStepen, ScienceZvanie, PassportDate, SNILS, toun, strit, house, flow, ind, telephone, diplomDate, placeWorking, Dolgnost, INN, dipSerNum, PlaseBirth)" +
                            "VALUES (@Caf_ID, @T_Family, @T_FirstName, @T_SecondNamme, @dateOfBirth, @ScienceStepen, @ScienceZvanie, @PassportDate, @SNILS, @toun, @strit, @house, @flow, @ind, @telephone, @diplomDate,@placeWorking,@Dolgnost, @INN, @dipSerNum, @PLB)", conn))
                            {
                                //cmd.Parameters.AddWithValue("@T_ID", Convert.ToInt16(heaher[0]) + 1);@T_ID T_ID, 
                                cmd.Parameters.AddWithValue("@Caf_ID", teacher.Caf);
                                cmd.Parameters.AddWithValue("@T_Family", teacher.Family);
                                cmd.Parameters.AddWithValue("@T_FirstName", teacher.Name);
                                cmd.Parameters.AddWithValue("@T_SecondNamme", teacher.SecName);
                                cmd.Parameters.AddWithValue("@dateOfBirth", teacher.Birth);
                                cmd.Parameters.AddWithValue("@ScienceStepen", teacher.ScienceStep);
                                cmd.Parameters.AddWithValue("@ScienceZvanie", teacher.ScienceZV);
                                cmd.Parameters.AddWithValue("@PassportDate", teacher.PassDate);
                                cmd.Parameters.AddWithValue("@SNILS", teacher.SNLS);
                                cmd.Parameters.AddWithValue("@toun", teacher.Toun);
                                cmd.Parameters.AddWithValue("@strit", teacher.Strit);
                                cmd.Parameters.AddWithValue("@house", teacher.House);
                                cmd.Parameters.AddWithValue("@flow", teacher.Flow);
                                cmd.Parameters.AddWithValue("@ind", teacher.Index);
                                cmd.Parameters.AddWithValue("@telephone", teacher.Telephone);
                                cmd.Parameters.AddWithValue("@diplomDate", teacher.Diplom);
                                cmd.Parameters.AddWithValue("@placeWorking", teacher.PlaceWorking);
                                cmd.Parameters.AddWithValue("@Dolgnost", teacher.dolgnost);
                                cmd.Parameters.AddWithValue("@INN", teacher.INN);
                                cmd.Parameters.AddWithValue("@dipSerNum", teacher.dipSerNum);
                                cmd.Parameters.AddWithValue("@PLB", teacher.PLB);

                                //add whatever parameters you required to update here
                                int rows = cmd.ExecuteNonQuery();
                            }
                        }
                        
                        conn.Close();

                    }
                }
                teachers.Clear();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void добавитьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            var dialog = new System.Windows.Forms.OpenFileDialog();
            dialog.FileName = "Document"; // Default file name
            dialog.DefaultExt = ".xlsx"; // Default file extension
            dialog.Filter = "Excel (.xlsx)|*.xlsx"; // Filter files by extension

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
                lastUsedRow = worksheet.Cells.Find("*", System.Reflection.Missing.Value,
                                   System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                   Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious,
                                   false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                // Find the last real column
                lastUsedColumn = worksheet.Cells.Find("*", System.Reflection.Missing.Value,
                                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                               Microsoft.Office.Interop.Excel.XlSearchOrder.xlByColumns, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious,
                                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
                //}
                //catch (Exception ex) { MessageBox.Show(ex.Message); }
                //try
                //{

                using (ExcelHelper helper = new ExcelHelper())
                {

                    if (helper.Open(filePath: path))
                    {
                        teachers = new List<Dict>();
                        for (int i = 1; i <= lastUsedRow; i++)
                        {
                            teachers.Add(new Dict
                            {
                                Caf = Convert.ToString(helper.Get(column: "A", row: i)),
                                Family = Convert.ToString(helper.Get(column: "B", row: i)),
                                Name = Convert.ToString(helper.Get(column: "C", row: i)),
                                SecName = Convert.ToString(helper.Get(column: "D", row: i)),
                                Birth = DateTime.FromOADate(Convert.ToDouble(helper.Get(column: "E", row: i))),

                                ScienceStep = Convert.ToString(helper.Get(column: "F", row: i)),
                                ScienceZV = Convert.ToString(helper.Get(column: "G", row: i)),
                                PassDate = Convert.ToString(helper.Get(column: "H", row: i)),
                                SNLS = Convert.ToString(helper.Get(column: "I", row: i)),
                                Toun = Convert.ToString(helper.Get(column: "J", row: i)),
                                Strit = Convert.ToString(helper.Get(column: "K", row: i)),
                                House = Convert.ToString(helper.Get(column: "L", row: i)),
                                Flow = Convert.ToString(helper.Get(column: "M", row: i)),
                                Index = Convert.ToString(helper.Get(column: "N", row: i)),
                                Telephone = Convert.ToString(helper.Get(column: "O", row: i)),
                                Diplom = DateTime.FromOADate(Convert.ToDouble(helper.Get(column: "P", row: i))),
                                PlaceWorking = Convert.ToString(helper.Get(column: "Q", row: i)),
                                dolgnost = Convert.ToString(helper.Get(column: "R", row: i)),
                                INN = Convert.ToString(helper.Get(column: "T", row: i)),
                                dipSerNum = Convert.ToString(helper.Get(column: "S", row: i)),
                                PLB = Convert.ToString(helper.Get(column: "U", row: i))
                            });
                        }
                        //helper.Save();
                    }
                }
                Console.Read();
                //}
                //catch (Exception ex) { MessageBox.Show(ex.Message); }
            }
            //try
            //{
            foreach (Dict teacher in teachers)
            {

                using (SqlConnection conn = new SqlConnection(Properties.Settings.Default.connString))
                {
                    StringBuilder queryNaDob = new StringBuilder();
                    queryNaDob.Append(@"SELECT * FROM GEK WHERE G_Family =");
                    queryNaDob.Append("'");
                    queryNaDob.Append(teacher.Family + "'" + " and " + "G_FirstName =" + "'" + teacher.Name + "'" + " and " + "G_SecondNamme =" + "'" + teacher.SecName + "'");
                    conn.Open();
                    SqlDataAdapter sqlQuery = new SqlDataAdapter(Convert.ToString(queryNaDob), connection);
                    System.Data.DataTable dtblDob = new System.Data.DataTable();
                    sqlQuery.Fill(dtblDob);
                    if (dtblDob.Rows.Count == 0)
                    {
                        string query = "WITH SRC AS ( SELECT TOP(1) G_ID, G_Family FROM GEK ORDER BY G_ID DESC ) SELECT* FROM SRC ORDER BY G_ID";
                        SqlDataAdapter sql = new SqlDataAdapter(query, connection);
                        System.Data.DataTable dtbl = new System.Data.DataTable();
                        sql.Fill(dtbl);

                        
                        using (SqlCommand cmd = new SqlCommand("INSERT INTO GEK(Caf_ID, G_Family, G_FirstName, G_SecondNamme, dateOfBirth, ScienceStepen, ScienceZvanie, PassportDate, SNILS, toun, strit, house, flow, ind, telephone, diplomDate, placeWorking, Dolgnost, INN, dipSerNum, PlaceBirth)" +
                            "VALUES (@Caf_ID, @T_Family, @T_FirstName, @T_SecondNamme, @dateOfBirth, @ScienceStepen, @ScienceZvanie, @PassportDate, @SNILS, @toun, @strit, @house, @flow, @ind, @telephone, @diplomDate,@placeWorking,@Dolgnost, @INN, @dipSerNum, @PLB)", conn))
                        {
                            //cmd.Parameters.AddWithValue("@T_ID", Convert.ToInt16(heaher[0]) + 1);@T_ID T_ID, 
                            cmd.Parameters.AddWithValue("@Caf_ID", teacher.Caf);
                            cmd.Parameters.AddWithValue("@T_Family", teacher.Family);
                            cmd.Parameters.AddWithValue("@T_FirstName", teacher.Name);
                            cmd.Parameters.AddWithValue("@T_SecondNamme", teacher.SecName);
                            cmd.Parameters.AddWithValue("@dateOfBirth", teacher.Birth);
                            cmd.Parameters.AddWithValue("@ScienceStepen", teacher.ScienceStep);
                            cmd.Parameters.AddWithValue("@ScienceZvanie", teacher.ScienceZV);
                            cmd.Parameters.AddWithValue("@PassportDate", teacher.PassDate);
                            cmd.Parameters.AddWithValue("@SNILS", teacher.SNLS);
                            cmd.Parameters.AddWithValue("@toun", teacher.Toun);
                            cmd.Parameters.AddWithValue("@strit", teacher.Strit);
                            cmd.Parameters.AddWithValue("@house", teacher.House);
                            cmd.Parameters.AddWithValue("@flow", teacher.Flow);
                            cmd.Parameters.AddWithValue("@ind", teacher.Index);
                            cmd.Parameters.AddWithValue("@telephone", teacher.Telephone);
                            cmd.Parameters.AddWithValue("@diplomDate", teacher.Diplom);
                            cmd.Parameters.AddWithValue("@placeWorking", teacher.PlaceWorking);
                            cmd.Parameters.AddWithValue("@Dolgnost", teacher.dolgnost);
                            cmd.Parameters.AddWithValue("@INN", teacher.INN);
                            cmd.Parameters.AddWithValue("@dipSerNum", teacher.dipSerNum);
                            cmd.Parameters.AddWithValue("@PLB", teacher.PLB);

                            //add whatever parameters you required to update here
                            int rows = cmd.ExecuteNonQuery();
                        }
                    }
                    
                    conn.Close();

                }
            }
            teachers.Clear();
            //}
            //catch(Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
