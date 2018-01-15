using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace ExcelReaderaltqq
{
    public partial class ExcelDataReader : Form
    {
        public ExcelDataReader()
        {
            InitializeComponent();
            label2.Text = "None";
            groupBox1.Visible = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Microsoft Excel Worksheet (.xlsx)|*.xlsx";
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK) 
            {
                label2.Text = openFileDialog1.SafeFileName;
                groupBox1.Visible = true;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string constr = @"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + openFileDialog1.FileName + ";Extended Properties='Excel 8.0;HDR=Yes;';";
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = constr;
            OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = conn;
            conn.Open();
            List<string> ids = new List<string>(); //to store all the IDs fetched from excel sheet.
            OleDbDataReader dr;
            cmd.CommandText = "select [" + textBox1.Text +"] from [Sheet1$]";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                ids.Add(Convert.ToString(dr[0])); //read  from excel sheet
            }
            conn.Close();

            //write data to SQL tables
            string myfilepath = @"C:\Users\Pasha\Documents\ExcelReaderaltqqDB.mdf"; //edit your file path here.
            String SQLconstr = @"Data Source=(LocalDB)\v11.0;AttachDbFilename=" + myfilepath + ";Integrated Security=True;Connect Timeout=30";
            SqlConnection sqlconn = new SqlConnection();
            sqlconn.ConnectionString = SQLconstr;
            SqlCommand sqlcmd = new SqlCommand();
            sqlcmd.Connection = sqlconn;
            sqlconn.Open();
            foreach (string id in ids)
            {
                sqlcmd.CommandText = "UPDATE [Table1] set [columnDate] = CONVERT(DATE,GETDATE()), [columnBool] = 1, [columnString]='foo' where [id]='" + id + "'";
                sqlcmd.ExecuteNonQuery();
                sqlcmd.CommandText = "UPDATE [Table2] set [columnDate2] = CONVERT(DATE,GETDATE()), [columnBool2] = 1, [columnString2]='bar' where [id]='" + id + "'";
                sqlcmd.ExecuteNonQuery();
            }
            MessageBox.Show("Both tables are updated!");
            sqlconn.Close();

        }
    }
}
