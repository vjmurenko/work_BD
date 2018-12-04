using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;



namespace WorkingWithDB
{
    public partial class Form1 : Form
    {
       
        List<string> collect = new List<string>();
        

        public Form1()
        {
            InitializeComponent();
            collect.Add("Открытие программы: "+DateTime.Now.ToString());



        }
        OleDbConnection con = new OleDbConnection("Provider = MSDAORA;DATA SOURCE=localhost:1521/xe;" +
            "PASSWORD=palace5;PERSIST SECURITY INFO=True;USER ID=system");




        private void bt_load_Click(object sender, EventArgs e)
        {
            con.Open();
            OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT ID_PERSON ,PERSON.NAME,DATE_BIRTH AS BIRTH, MONEY," +
                " CAR.NAME AS CAR_NAME,CAR.DATE_MANUFACTURE AS CAR_DATE,CAR.COST AS CAR_COST, CAR.RATING AS CAR_RATING," +
                " PHONE.NAME AS PHONE_NAME, PHONE.DATE_MANUFACTURE AS PHONE_DATE, PHONE.COST AS PHONE_COST," +
                " PHONE.RATING AS PHONE_RATING FROM PERSON JOIN CAR ON ID_PERSON = CAR.ID_CAR " +
                "JOIN PHONE ON ID_PERSON = PHONE.ID_PHONE ", con);
            DataTable datatable = new DataTable();

            adapter.Fill(datatable);
            dataGridView1.DataSource = datatable;
            con.Close();

        }

        private void bt_export_Click(object sender, EventArgs e)
        {
            collect.Add("Нажат export: " + DateTime.Now.ToString());
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Document (*.xls)|*.xls";
            sfd.FileName = "export.xls";
            if (sfd.ShowDialog()== DialogResult.OK)
            {
                ToExcel(dataGridView1, sfd.FileName);
            }

        }

        private void ToExcel(DataGridView dGV, string filename)
        {

            string stOutput = "";
            string sHeaders = "";
            for (int j = 0; j < dGV.Columns.Count; j++)
                   sHeaders = sHeaders.ToString() + Convert.ToString(dGV.Columns[j].HeaderText) + "\t";
                stOutput += sHeaders + "\r\n";
                for (int i = 0; i < dGV.RowCount - 1; i++)
                {
                    string stline = "";
                    for (int j = 0; j < dGV.Rows[i].Cells.Count; j++)
                        stline = stline.ToString() + Convert.ToString(dGV.Rows[i].Cells[j].Value) + "\t";
                    stOutput += stline + "\r\n";
                }
                Encoding utf16 = Encoding.GetEncoding(1254);
                byte[] output = utf16.GetBytes(stOutput);
                FileStream fs = new FileStream(filename, FileMode.Create);
                BinaryWriter bw = new BinaryWriter(fs);
                bw.Write(output, 0, output.Length);
                bw.Flush();
                bw.Close();
                fs.Close();
        }

       

        private void bt_add_Click(object sender, EventArgs e)
        {
            collect.Add("Нажата Update: "  +DateTime.Now.ToString());
            con.Open();
            OleDbCommand cmd = new OleDbCommand("DEL", con);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.ExecuteReader();

            OleDbCommand cmd1 = new OleDbCommand("ADD_D", con);
            cmd1.CommandType = CommandType.StoredProcedure;
            cmd1.ExecuteReader();
            con.Close();
        }

        

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {


            switch (comboBox1.SelectedIndex)
            {
                case 0:
                    collect.Add("Выбран фильтр 10: " + DateTime.Now.ToString());
                    con.Open();
                    OleDbDataAdapter adapter1 = new OleDbDataAdapter("SELECT ID_PERSON ,PERSON.NAME,DATE_BIRTH AS BIRTH, MONEY," +
                " CAR.NAME AS CAR_NAME,CAR.DATE_MANUFACTURE AS CAR_DATE,CAR.COST AS CAR_COST, CAR.RATING AS CAR_RATING," +
                " PHONE.NAME AS PHONE_NAME, PHONE.DATE_MANUFACTURE AS PHONE_DATE, PHONE.COST AS PHONE_COST," +
                " PHONE.RATING AS PHONE_RATING FROM PERSON JOIN CAR ON ID_PERSON = CAR.ID_CAR " +
                "JOIN PHONE ON ID_PERSON = PHONE.ID_PHONE  WHERE ID_PERSON <=10", con);
                    DataTable datatable1 = new DataTable();

                    adapter1.Fill(datatable1);
                    dataGridView1.DataSource = datatable1;
                    con.Close();

                    break;

                case 1:
                    collect.Add("Выбран фильтр 20: " + DateTime.Now.ToString());
                    con.Open();
                    OleDbDataAdapter adapter2 = new OleDbDataAdapter("SELECT ID_PERSON ,PERSON.NAME,DATE_BIRTH AS BIRTH, MONEY," +
                " CAR.NAME AS CAR_NAME,CAR.DATE_MANUFACTURE AS CAR_DATE,CAR.COST AS CAR_COST, CAR.RATING AS CAR_RATING," +
                " PHONE.NAME AS PHONE_NAME, PHONE.DATE_MANUFACTURE AS PHONE_DATE, PHONE.COST AS PHONE_COST," +
                " PHONE.RATING AS PHONE_RATING FROM PERSON JOIN CAR ON ID_PERSON = CAR.ID_CAR " +
                "JOIN PHONE ON ID_PERSON = PHONE.ID_PHONE  WHERE ID_PERSON <=20", con);
                    DataTable datatable2 = new DataTable();

                    adapter2.Fill(datatable2);
                    dataGridView1.DataSource = datatable2;
                    con.Close();

                    break;

                case 2:
                    collect.Add("Выбран фильтр 30: " + DateTime.Now.ToString());
                    con.Open();
                    OleDbDataAdapter adapter3 = new OleDbDataAdapter("SELECT ID_PERSON ,PERSON.NAME,DATE_BIRTH AS BIRTH, MONEY," +
                " CAR.NAME AS CAR_NAME,CAR.DATE_MANUFACTURE AS CAR_DATE,CAR.COST AS CAR_COST, CAR.RATING AS CAR_RATING," +
                " PHONE.NAME AS PHONE_NAME, PHONE.DATE_MANUFACTURE AS PHONE_DATE, PHONE.COST AS PHONE_COST," +
                " PHONE.RATING AS PHONE_RATING FROM PERSON JOIN CAR ON ID_PERSON = CAR.ID_CAR " +
                "JOIN PHONE ON ID_PERSON = PHONE.ID_PHONE  WHERE ID_PERSON <=30", con);
                    DataTable datatable3 = new DataTable();

                    adapter3.Fill(datatable3);
                    dataGridView1.DataSource = datatable3;
                    con.Close();

                    break;


                case 3:
                    collect.Add("Выбран фильтр 40: " + DateTime.Now.ToString());
                    con.Open();
                    OleDbDataAdapter adapter4 = new OleDbDataAdapter("SELECT ID_PERSON ,PERSON.NAME,DATE_BIRTH AS BIRTH, MONEY," +
                " CAR.NAME AS CAR_NAME,CAR.DATE_MANUFACTURE AS CAR_DATE,CAR.COST AS CAR_COST, CAR.RATING AS CAR_RATING," +
                " PHONE.NAME AS PHONE_NAME, PHONE.DATE_MANUFACTURE AS PHONE_DATE, PHONE.COST AS PHONE_COST," +
                " PHONE.RATING AS PHONE_RATING FROM PERSON JOIN CAR ON ID_PERSON = CAR.ID_CAR " +
                "JOIN PHONE ON ID_PERSON = PHONE.ID_PHONE  WHERE ID_PERSON <=40", con);
                    DataTable datatable4 = new DataTable();

                    adapter4.Fill(datatable4);
                    dataGridView1.DataSource = datatable4;
                    con.Close();

                    break;


                case 4:
                    collect.Add("Выбран фильтр 50+: " + DateTime.Now.ToString());
                    con.Open();
                    OleDbDataAdapter adapter5 = new OleDbDataAdapter("SELECT ID_PERSON ,PERSON.NAME,DATE_BIRTH AS BIRTH, MONEY," +
                " CAR.NAME AS CAR_NAME,CAR.DATE_MANUFACTURE AS CAR_DATE,CAR.COST AS CAR_COST, CAR.RATING AS CAR_RATING," +
                " PHONE.NAME AS PHONE_NAME, PHONE.DATE_MANUFACTURE AS PHONE_DATE, PHONE.COST AS PHONE_COST," +
                " PHONE.RATING AS PHONE_RATING FROM PERSON JOIN CAR ON ID_PERSON = CAR.ID_CAR " +
                "JOIN PHONE ON ID_PERSON = PHONE.ID_PHONE  WHERE ID_PERSON >=50", con);
                    DataTable datatable5 = new DataTable();

                    adapter5.Fill(datatable5);
                    dataGridView1.DataSource = datatable5;
                    con.Close();

                    break;

            }

        }

  

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            collect.Add("Закрытие программы: "+DateTime.Now.ToString());

            string path = "log.txt";

            if (!File.Exists(path))
            {            
                File.WriteAllLines(path, collect.ToArray(), Encoding.UTF8);
            }
            else
            {
                File.AppendAllLines(path, collect, Encoding.UTF8);
            }
        }
    }
}
