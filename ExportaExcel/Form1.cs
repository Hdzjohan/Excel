using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace ExportaExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents (*.xls)|*.xls";
            //Commit
            sfd.FileName = "MiExport.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                // Copy DataGridView al clipboard
                copyAlltoClipboard();

                object misValue = System.Reflection.Missing.Value;
                Excel.Application xlexcel = new Excel.Application();

                xlexcel.DisplayAlerts = false;// Sin esto, obtendrás dos mensajes de confirmación de sobrescritura
                Excel.Workbook xlWorkBook = xlexcel.Workbooks.Add(misValue);
                Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                // Format column B  siempre será  una despues para afectar una anterior 
                Excel.Range rng = xlWorkSheet.get_Range("C:C").Cells;
                rng.NumberFormat = "0.00000";

                // Pegar resultados del portapapeles en el rango de la hoja de trabajo
                Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[1, 1];
                CR.Select();
                xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);


                //// Borra la columna en blanco
                //Excel.Range delRng = xlWorkSheet.get_Range("A:A").Cells;
                //delRng.Delete(Type.Missing);
                //xlWorkSheet.get_Range("A1").Select();

                // Guarde el archivo de Excel en la ubicación capturada desde SaveFileDialog
                xlWorkBook.SaveAs(sfd.FileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlexcel.DisplayAlerts = true;
                xlWorkBook.Close(true, misValue, misValue);
                xlexcel.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlexcel);


                // Borrar la selección del Portapapeles y DataGridView
                Clipboard.Clear();
                Malla.ClearSelection();


                // Abra el archivo de Excel recién guardado
                if (File.Exists(sfd.FileName))
                    System.Diagnostics.Process.Start(sfd.FileName);
            }
        }

        private void copyAlltoClipboard()
        {
            Malla.SelectAll();
            DataObject dataObj = Malla.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
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
                MessageBox.Show("valio verga la exportacion " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        private void Button2_Click(object sender, EventArgs e)
        {
            var r = new Random();
            Malla.Rows.Clear();

            for (int i = 0; i < 10; i++)
            {
                var p = new Producto();
                p.ProductoId = i + 1;
                p.Precio = (r.NextDouble() + 1) * 1000;
                p.Descripcion = "Descripcion " + i + 1;

                Malla.Rows.Add();
                Malla.Rows[i].Cells[0].Value = p.ProductoId;
                Malla.Rows[i].Cells[1].Value = p.Precio;
                Malla.Rows[i].Cells[2].Value = p.Descripcion;
            }


        }
    }
}
