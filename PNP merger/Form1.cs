using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Threading;
using System.Globalization;
using System.Collections;

namespace PNP_merger
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string file;
        //private int rowIndex = 0;
        
        private void openExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //int size = -1;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ReadOnlyChecked = true;
            openFileDialog1.Filter = "Office Files|*.xls;";
            // Show the dialog and get result.
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                file = openFileDialog1.FileName;
                // }
                Console.WriteLine(result); // <-- For debugging use.

                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Add();
                xlWorkbook = xlApp.Workbooks.Open(file);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int i = 1;
                int j = 1;
                while (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                {    //outputData.AppendText(xlRange.Cells[i, j].Value2.ToString() + "\t");
                    string[] row = { xlRange.Cells[i, j].Value2.ToString(), xlRange.Cells[i, 2].Value2.ToString() };
                    dataGridView1.Rows.Add(row);
                    i++;
                }

                xlApp.Quit();

                //release all memory - stop EXCEL.exe from hanging around.
                if (xlWorkbook != null) { Marshal.ReleaseComObject(xlWorkbook); } //release each workbook like this
                if (xlWorksheet != null) { Marshal.ReleaseComObject(xlWorksheet); } //release each worksheet like this
                if (xlApp != null) { Marshal.ReleaseComObject(xlApp); } //release the Excel application
                xlWorkbook = null; //set each memory reference to null.
                xlWorksheet = null;
                xlApp = null;
                GC.Collect();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //int[] combined = new int [100]; //Max 100 combined reference desginators.
            char[] delimiterChars = { ' ', ',', '.', ':', '\t' };
            ArrayList list = new ArrayList();//list of all indexes that contains combined reference desginators. 
            ArrayList partNo = new ArrayList();
            String type;
            int count = 0;
            dataGridView1.Columns[0].Name = "Reference designator";
            dataGridView1.Columns[1].Name = "Part number";
            for (int i =0; i<dataGridView1.RowCount && dataGridView1[0, i].Value!=null; i++)
            {
                if (((dataGridView1[0, i].Value).ToString()).Contains(','))
                {
                    list.Add(i); //Get an array of indexes that containes combined reference desginators.
                    partNo.Add(dataGridView1[1, i].Value.ToString());
                }
            }

            for (int i = 0; i < list.Count; i++)
            {
                string[] words = dataGridView1[0,Int32.Parse(list[i].ToString())].Value.ToString().Split(delimiterChars);//get the type of reference designators.
                string[] rowContent=new string[2];
                if (Char.IsLetter(words[0][1]))//check whethere there are 2 letter for the ref des
                    type = words[0][0].ToString() + words[0][1].ToString();
                else
                    type = words[0][0].ToString();
                for(int f=0; f < words.Length; f++)//add in the rows at the end
                {
                    if (Char.IsLetter(words[f][0]))
                        ;
                    else
                        words[f] = type + words[f];
                    rowContent[0] = words[f];
                    rowContent[1] = partNo[i].ToString();
                    dataGridView1.Rows.Add(rowContent);
                }             
            }
            for(int i=list.Count-1; i>=0; i--)
            {
                dataGridView1.Rows.Remove(dataGridView1.Rows[Int32.Parse(list[i].ToString())]);
            }
            count = dataGridView1.Rows.Count-1;
            countText.Text = count.ToString();
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //int size = -1;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ReadOnlyChecked = true;
            openFileDialog1.Filter = "Office Files|*.xls;";
            // Show the dialog and get result.
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                file = openFileDialog1.FileName;
                // }
                Console.WriteLine(result); // <-- For debugging use.

                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Add();
                xlWorkbook = xlApp.Workbooks.Open(file);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int i = 1;
                //int j = 1;
                while (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
                {    //outputData.AppendText(xlRange.Cells[i, j].Value2.ToString() + "\t");
                    string[] row = { xlRange.Cells[i, 1].Value2.ToString(), xlRange.Cells[i, 2].Value2.ToString(), xlRange.Cells[i, 3].Value2.ToString(), xlRange.Cells[i, 4].Value2.ToString(), xlRange.Cells[i, 5].Value2.ToString()};
                    dataGridView2.Rows.Add(row);
                    i++;
                }

                xlApp.Quit();

                //release all memory - stop EXCEL.exe from hanging around.
                if (xlWorkbook != null) { Marshal.ReleaseComObject(xlWorkbook); } //release each workbook like this
                if (xlWorksheet != null) { Marshal.ReleaseComObject(xlWorksheet); } //release each worksheet like this
                if (xlApp != null) { Marshal.ReleaseComObject(xlApp); } //release the Excel application
                xlWorkbook = null; //set each memory reference to null.
                xlWorksheet = null;
                xlApp = null;
                GC.Collect();
            }
        }

        private void count2Btn_Click(object sender, EventArgs e)
        {
            int count = dataGridView2.Rows.Count - 1;
            countText2.Text = count.ToString();
        }

        private void editToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView2.Rows.Count!=1 && dataGridView1.Rows.Count!=1)
            {
                mergeToolStripMenuItem.Enabled=true;
            }
        }

        private void mergeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //ArrayList componentMatch = new ArrayList();
            ArrayList componentTop = new ArrayList();
            ArrayList componentBot = new ArrayList();
            ArrayList componentMissing = new ArrayList();
            Boolean componentFound = false;
            //string[] componentDetails = new string[6];
            for (int i = 0; i < dataGridView1.RowCount && dataGridView1[0, i].Value != null; i++)
            {
                if (componentFound)
                {
                    componentMissing.Add((dataGridView1[0, i].Value).ToString());
                }
                componentFound = false;

                for (int j = 0; j < dataGridView2.RowCount && dataGridView2[0, j].Value != null; j++)
                {
                    if (((dataGridView1[0, i].Value).ToString()) == ((dataGridView2[0, j].Value).ToString())) // if a component in the BOM matches a component in the PNP file
                    {
                        string[] componentDetails = new string[6];
                        componentDetails[0] = (dataGridView1[0, i].Value).ToString(); //ref des
                        componentDetails[1] = (dataGridView2[1, j].Value).ToString(); //x
                        componentDetails[2] = (dataGridView2[2, j].Value).ToString();//y
                        componentDetails[3] = (dataGridView2[3, j].Value).ToString();//z
                        componentDetails[4] = (dataGridView2[4, j].Value).ToString();//top or bot
                        componentDetails[5] = (dataGridView1[1, i].Value).ToString();//description
                        if (componentDetails[4][0].Equals('t') || componentDetails[4][0].Equals('T'))
                        {
                            componentTop.Add(componentDetails);
                        }
                        if (componentDetails[4][0].Equals('b') || componentDetails[4][0].Equals('B'))
                        {
                            componentBot.Add(componentDetails);
                        }
                        componentFound = true;
                    }
                }

                //if (componentNotFound)
                //{
                //    componentDetails[0] = (dataGridView1[0, i].Value).ToString(); //ref des
                //    componentMissing.Add(componentDetails);
                //}
            }

        }
    }
}
