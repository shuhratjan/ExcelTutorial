using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelTutorial
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void  Form1_Load(object sender, EventArgs e)
        {

            // OpenFile();
            // WriteData();
            // OpenFile();
            CreateNewExcel("TestNewExcel");
        }

        private void CreateNewExcel(string fileName)
        {
            Excel excel = new Excel();
            excel.CreateNewFile();
            excel.CreateNewSheet();
            excel.SaveAs(@""+ fileName);
            excel.Close();
        }

        private void OpenFile()
        {
            // D:\Projects\VSProjects\ExcelTutorial\ExcelTutorial\
            Excel excel = new Excel(@"Test.xlsx", 1);
            MessageBox.Show(excel.ReadCell(0, 0));
        }

        public void WriteData()
        {
            Excel excel = new Excel(@"Test.xlsx", 1);
            excel.WriteToCell(0, 0, "TEst4 WillChange");
            excel.Save();
            excel.SaveAs(@"Test2.xlsx");
            excel.Close();
        }
    }
}
