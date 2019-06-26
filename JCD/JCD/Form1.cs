using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Configuration;
using System.IO.Ports;
using Microsoft.Office.Interop.Excel;

namespace JCD
{
    public partial class Form1 : Form
    {
        // déclaration des éléements de connexion avec l'Arduino
        OleDbDataAdapter da;
        DataSet ds;
        DataTableCollection tables;
        BindingSource source1;
        DataView view = new DataView();
        public string _productname;
        SerialPort port = new SerialPort();
        double voltagevalue;

        //excel file
        Microsoft.Office.Interop.Excel.Application oXL;
        Microsoft.Office.Interop.Excel._Workbook oWB;
        Microsoft.Office.Interop.Excel._Worksheet oSheet;
        Microsoft.Office.Interop.Excel.Range oRng;

        public Form1()
        {
            InitializeComponent();
            port.PortName = "COM3";
            port.BaudRate = 9600;
            port.DtrEnable = true;
            port.Open();
        }
        private void serialPort1_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            //string line = port.ReadLine();
            //this.BeginInvoke(new LineReceivedEvent(LineReceived), line);
        }

        private delegate void LineReceivedEvent(string line);
        private void LineReceived(string line)
        {
            //What to do with the received line here
            //voltagevalue = Math.Round(((Convert.ToDouble(line) * 5.0) / 1024.0), 2);
            //textBox2.Text = Convert.ToString(voltagevalue);
        }

        //  Lecture des mesures
        private void Button1_Click(object sender, EventArgs e)
        {
            //port.DataReceived += serialPort1_DataReceived;
            port.Write("V");
            var value = port.ReadLine();
            System.Threading.Thread.Sleep(500);
            voltagevalue = Math.Round(((Convert.ToDouble(value) * 5.0) / 1024.0), 2);
            System.Threading.Thread.Sleep(500);
            textBox2.Text = Convert.ToString(voltagevalue);
            //System.Threading.Thread.Sleep(500);
            compar();
            Writetoexcel_Click(sender, e);
        }

        private void compar()
        {
            if (voltagevalue >= 3.60 && voltagevalue <= 4.10) pictureBox1.Image = new Bitmap(@"C:\Users\ayoubexo.CEBONGROUP\source\repos\JCD\vert.PNG");
            else pictureBox1.Image = new Bitmap(@"C:\Users\ayoubexo.CEBONGROUP\source\repos\JCD\rouge.PNG");
        }

        private void Label1_Click(object sender, EventArgs e)
        {

        }

        private void Label3_Click(object sender, EventArgs e)
        {

        }

        private void DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Button1_Click_1(object sender, EventArgs e)
        {
            if (_productname == "NCM18650260")
            {
                OleDbConnection con = new OleDbConnection();
                con.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\ayoubexo.CEBONGROUP\source\repos\JCD\Fichier access.accdb;Persist Security Info = False;";
                ds = new DataSet();
                tables = ds.Tables;
                da = new OleDbDataAdapter("Select * from " + _productname + "", con);
                da.Fill(ds, _productname);
                DataView view = new DataView(tables[0]);
                source1 = new BindingSource();
                source1.DataSource = view;
                dataGridView1.DataSource = view;
            }
        }

        private void Writetoexcel_Click(object sender, EventArgs e)
        {
                //Start Excel and get Application object.
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.Visible = true;

                //Get a new workbook.
                oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                //Add table headers going cell by cell
                oSheet.Cells[1, 1] = "voltage value is : ";
                oSheet.Cells[1, 2] = textBox2.Text;
                oXL.Visible = false;
                oXL.UserControl = false;
                string date = Convert.ToString(DateTime.UtcNow);
                string[] date2 = date.Split('/', ':');
                string filename = "C:\\Users\\ayoubexo.CEBONGROUP\\source\\repos\\JCD\\JCD\\excel files\\"+ string.Join("-",date2) +"_test result.xlsx";
                oWB.SaveAs(filename, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                oWB.Close();           
    }


        private void Button2_Click(object sender, EventArgs e)
        {
            
        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void TextBox2_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
