using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

namespace ArchivosOffice
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

     private void Ptbword_Click(object sender, EventArgs e)
        {
            SaveFileDialog sf = new SaveFileDialog();
            if (sf.ShowDialog() == DialogResult.OK)
            {
                Object obj = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();

                Microsoft.Office.Interop.Word.Document doc = word.Documents.Add(ref obj);
                doc.Activate();
                word.Selection.TypeText(textBox1.Text);
                doc.SaveAs2(sf.FileName);
                word.Visible = true;
            }
        }

        private void PtbExcel_Click(object sender, EventArgs e)
        {
            SaveFileDialog sf = new SaveFileDialog();
            if (sf.ShowDialog() == DialogResult.OK)
            {
                Object obj = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook file = excel.Workbooks.Add(obj);
                Microsoft.Office.Interop.Excel.Worksheet hoja = file.Worksheets.Add();
                file.Activate();
                hoja.Cells[1, 1].Value = textBox1.Text;
                file.SaveAs(sf.FileName);
                excel.Visible = true;
            }

        }
    }
    }

