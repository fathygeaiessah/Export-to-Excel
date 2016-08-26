using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Xml;
using System.Xml.XPath;
using System.Runtime.InteropServices;


namespace Export_to_Excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

     public   Excel.Application xlApp;
       public  Excel.Workbook xlWBook;
       public  Excel.Worksheet xlWSheet;
        string FileName = "";

        Word.Application WordApp;
        Word.Document doc;
      

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                 xlApp = new Excel.Application();
            xlWBook = xlApp.Workbooks.Add();
            xlWSheet = xlWBook.ActiveSheet;
            xlApp.Visible = true;
            xlApp.UserControl = true;
           
            btnCloseExcel.Enabled = true;
            btnCreateChart.Enabled = true;
            btnSaveAs.Enabled = true;
            btnFormulas_Dialog.Enabled = true;
            btnSaveText.Enabled = true;
            button1.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = true;
            button4.Enabled = true;
            button5.Enabled = true;
            btnMacros.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            btnSaveAs_Click(sender, e);
           
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btnSaveText_Click(object sender, EventArgs e)
        {
            //xlApp = new Excel.Application();
            //xlWBook = xlApp.Workbooks.Open(FileName);
            //xlWSheet = xlWBook.ActiveSheet;

            try
            {
                 Excel.Range rng = xlWSheet.Cells.get_Range("" + cmb1.Text + (int)No1.Value + "");
                 rng.Value = txtEnter.Text;
                 rng.Formula = txtEnter.Text;
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
           
        }
  

        private void btnSaveAs_Click(object sender, EventArgs e)
        {
          
            if (saveFileDialog1.ShowDialog()==DialogResult.OK)
            {
                FileName = saveFileDialog1.FileName;
                xlWBook.SaveAs(FileName);
            }
          
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            try
            {
                 if (openFileDialog1.ShowDialog()==DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;
                 xlApp = new Excel.Application();
            xlWBook = xlApp.Workbooks.Open(FileName);
            xlWSheet = xlWBook.Worksheets.get_Item(1);
            xlApp.Visible = true;
            //xlApp.UserControl = true;
           
            btnCloseExcel.Enabled = true;
            btnCreateChart.Enabled = true;
            btnSaveAs.Enabled = true;
            btnFormulas_Dialog.Enabled = true;
            btnSaveText.Enabled = true;
            button1.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = true;
            button4.Enabled = true;
            button5.Enabled = true;
            btnMacros.Enabled = true;
            }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
          
           
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
           
            try
            {
                 Excel.Range rng = xlWSheet.Cells.get_Range("" + cmb1.Text + (int)No1.Value + "", "" + cmb2.Text + (int)No2.Value + "");
            rng.Merge();
           
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
           
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //xlApp = new Excel.Application();
            //xlWBook = xlApp.Workbooks.Open(FileName);
            //xlWSheet = xlWBook.ActiveSheet;
            try
            {
                 Excel.Range rng = xlWSheet.Cells.get_Range("" + cmb1.Text + (int)No1.Value + "", "" + cmb2.Text + (int)No2.Value + "");
            fontDialog1.ShowDialog();
            rng.Font.Name = fontDialog1.Font.Name;
            rng.Font.Bold = fontDialog1.Font.Bold;
            rng.Font.Size = fontDialog1.Font.Size;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //xlApp = new Excel.Application();
            //xlWBook = xlApp.Workbooks.Open(FileName);
            //xlWSheet = xlWBook.ActiveSheet;
            Excel.Range rng = xlWSheet.Cells.get_Range("" + cmb1.Text + (int)No1.Value + "", "" + cmb2.Text + (int)No2.Value + "");
            colorDialog1.ShowDialog();
            rng.Font.Color = colorDialog1.Color;

           
            
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //xlApp = new Excel.Application();
            //xlWBook = xlApp.Workbooks.Open(FileName);
            //xlWSheet = xlWBook.ActiveSheet;
            try
            {
                 Excel.Range rng = xlWSheet.Cells.get_Range("" + cmb1.Text + (int)No1.Value + "", "" + cmb2.Text + (int)No2.Value + "");
            colorDialog1.ShowDialog();
            rng.Interior.Color = colorDialog1.Color;
           
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
          
        }


        private void Text2Excel(string fileName)
        {
            try
            {
                  xlApp = new Excel.Application();
            xlWBook = xlApp.Workbooks.Add();
            xlWSheet = xlWBook.Worksheets.get_Item(1);
            var  lines = File.ReadAllLines(fileName);
            int rowNo = 1;
            for (int i = 0; i < lines.Length; i++)
            {
                int colNo = cmb1.SelectedIndex + 1;
                var values = lines[i].Trim().Split(' ');
                if (i == 1)
                {
                    values = lines[1].Trim().Insert(0, "z ").Split(' ');
                }

                foreach (var item in values)
                {
                    xlWSheet.Cells[rowNo, colNo] = item == "z" ? "  " : item;
                    colNo++;
                }
                rowNo++;
            }
           
            if (saveFileDialog1.ShowDialog()==DialogResult.OK)
            {
                xlWBook.SaveAs(saveFileDialog1.FileName);
            }
            xlWBook.Close();
          
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

       
        private void Word2Excel(string fileName)
        {
            try
            {
                   xlApp = new Excel.Application();
            xlWBook = xlApp.Workbooks.Add();
            xlWSheet = xlWBook.Worksheets.get_Item(1);
            WordApp = new Word.Application();
            int b = cmb1.SelectedIndex;
           
            doc = WordApp.Documents.Open(fileName);
            Word.Range[] rngPa=new Word.Range[doc.Paragraphs.Count];
            Word.Range[] rngTbl=new Word.Range[doc.Tables.Count];
            Word.Range[] rngAll=new Word.Range[doc.Tables.Count+doc.Paragraphs.Count];
         //   string[] pa = new string[doc.Paragraphs.Count];
            int t = 0;
           
            for (int i = 0; i <doc.Paragraphs.Count ; i++)
            {
                rngPa[i] = doc.Paragraphs[i + 1].Range;
                rngAll[i] = rngPa[i];
               // pa[i] = doc.Paragraphs[i + 1].Range.Text;
            }

            for (int i = 0; i < doc.Tables.Count; i++)
            {
                rngTbl[i] = doc.Tables[i + 1].Range;
                rngAll[doc.Paragraphs.Count + i] = rngTbl[i];
            }
            Word.Range[] q = (from a in rngAll orderby a.Start select a).ToArray();
            Word.Range[] qp = (from a in rngPa orderby a.Start select a).ToArray();
            Word.Range[] qt = (from a in rngTbl orderby a.Start select a).ToArray();
       
            foreach (var item in q)
            {
                Word.Range  qtbl = (from a in qt where 1 == -1 select a).SingleOrDefault();
                Word.Range  qpa = (from a in qp where 1 == -1 select a).SingleOrDefault();
                if (qt.Count() > 0)
                {
                     qtbl = (from a in qt where a == item  select a).SingleOrDefault();
                }
                if (qp.Count() > 0)
                {
                    qpa = (from a in qp where a == item  select a).SingleOrDefault();
                }
                if (qtbl==null && qpa != null)
                {
                    if (qt.Length >= 0)
                    {
                        int i=0;
                       while(i < rngTbl.Length)
                        {
                            if (item.InRange(rngTbl[i]) && item != null)
                            {
                                break;
                            }
                                i=i+1;
                       }
                            if(i >= rngTbl.Length)
                            {
                                Excel.Range rngEx = xlWSheet.get_Range("" + cmb1.Text + ((int)No1.Value + t) + "", "" + cmb2.Text + ((int)No2.Value + t) + "");
                                xlWSheet.Rows.AutoFit();
                                rngEx.Merge();
                                rngEx.Font.Size = item.Font.Size;
                                rngEx.Font.ColorIndex = item.Font.ColorIndex;
                                rngEx.Font.Bold = item.Font.Bold;

                                rngEx.Value = qpa.Text;
                                
                                qp = (from a in qp where a != item select a).ToArray();
                                q = (from a in q where a != item select a).ToArray();
                                t = t + 1;
                                goto sd;
                            }
                        }
                    
                sd: { }
                }
                else if(qtbl != null && qpa == null)
                {
                    for (int x = 1; x <= qtbl.Tables[1].Rows.Count; x++)
                    {
                        for (int y = 1; y <= qtbl.Tables[1].Columns.Count; y++)
                        {
                            decimal s = 0;
                            if (decimal.TryParse(qtbl.Tables[1].Cell(x, y).Range.Text.Remove(qtbl.Tables[1].Cell(x, y).Range.Text.Length - 1, 1), out s))
                            {
                                xlWSheet.Cells[(int)No1.Value + x + t, b + y] = s;
                            }
                            else
                            {
                                xlWSheet.Cells[(int)No1.Value + x + t, b + y] = qtbl.Tables[1].Cell(x, y).Range.Text == null ? " " : qtbl.Tables[1].Cell(x, y).Range.Text.Remove(qtbl.Tables[1].Cell(x, y).Range.Text.Length - 1, 1);

                            }

                        }

                    }
                    xlWSheet.Columns.AutoFit();
                    t = t + 1 + qtbl.Tables[1].Rows.Count;
                    qt = (from a in qt where a != item select a).ToArray();
                    q = (from a in q where a != item select a).ToArray();
                   
                }
            }
          
            if (saveFileDialog1.ShowDialog()==DialogResult.OK)
            {
                FileName = saveFileDialog1.FileName;
                 xlWBook.SaveAs(FileName);
            }
           
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void btnExportWordToExcel_Click(object sender, EventArgs e)
        {
              if (openFileDialog1.ShowDialog()==DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;
            Word2Excel(FileName);
            }
        }

        public void FormatAsTable(Excel.Range SourceRange, string TableName, string TableStyleName)
        {
            SourceRange.Worksheet.ListObjects.Add(Microsoft.Office.Interop.Excel.XlListObjectSourceType.xlSrcRange, SourceRange, System.Type.Missing, Microsoft.Office.Interop.Excel.XlYesNoGuess.xlYes, System.Type.Missing ).Name = TableName;
             //SourceRange.Select();
            SourceRange.Worksheet.ListObjects[TableName].TableStyle = TableStyleName;
    
        }


        private void btnCreateChart_Click(object sender, EventArgs e)
        {

            try
            {

                Excel.Chart oChart = xlWBook.Charts.Add();


                Excel.Range rngChart = (Excel.Range)xlWSheet.get_Range("" + cmb1.Text + (int)No1.Value + "", "" + cmb2.Text + (int)No2.Value + "");
                FormatAsTable(rngChart, "Table1", "TableStyleMedium1");
                oChart.ChartWizard(rngChart, Excel.XlChartType.xl3DColumn);
                oChart.Refresh();
                //xlWBook.AcceptAllChanges();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
          
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception occured while releasing object " + ex);
            }
            finally
            {
                GC.Collect();
            }
        }

        private void btnCloseExcel_Click(object sender, EventArgs e)
        {
            button5.Enabled = false;
            button4.Enabled = false;
            button3.Enabled = false;
            button2.Enabled = false;
            button1.Enabled = false;
            btnSaveText.Enabled = false;
            btnSaveAs.Enabled = false;
            btnFormulas_Dialog.Enabled = false;
            btnCreateChart.Enabled = false;
            btnCloseExcel.Enabled = false;
            xlWBook.Close();
            xlApp.Quit();
            releaseObject(xlWSheet);
            releaseObject(xlWBook);
            releaseObject(xlApp);
            
        }

        private void btnExportTxtToExcel_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog()==DialogResult.OK)
            {
                Text2Excel(openFileDialog1.FileName);
            }
        }
       
       

        private void button3_Click(object sender, EventArgs e)
        {
            //xlApp = new Excel.Application();
            //xlWBook = xlApp.Workbooks.Open(FileName);
            //xlWSheet = xlWBook.ActiveSheet;
            try
            {
                  Excel.Range rng = xlWSheet.Cells.get_Range("" + cmb1.Text + (int)No1.Value + "", "" + cmb2.Text + (int)No2.Value + "");
            rng.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic);
           
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
        }


        private void btnFormulas_Dialog_Click(object sender, EventArgs e)
        {
            //xlApp = new Excel.Application(); 
            //xlWBook = xlApp.Workbooks.Open(FileName);
            //xlWSheet = xlWBook.ActiveSheet;
            try
            {
                Excel.Dialog Mydialog;
      if (! xlApp.Visible)
      {
          xlApp.Visible = true;
      }  
       Mydialog  = (Excel.Dialog)xlApp.Dialogs[Excel.XlBuiltInDialog.xlDialogFunctionWizard];
      Mydialog.Show();
     
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
           
        }

        private void btnMacros_Click(object sender, EventArgs e)
        {
            frmMacros mac = new frmMacros(xlApp);
            mac.Show();
        }
        
    }
}
