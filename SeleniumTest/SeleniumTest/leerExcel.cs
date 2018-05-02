using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumTest
{
    public class leerExcel
    {
        //public List<SolicitudesDespacho> LeerExcel(string ArchivoExcel)
        //{
        //    List<SolicitudesDespacho> retorno = new List<SolicitudesDespacho>();
        //    SolicitudesDespacho solicitudesDespacho;

        //    string NombreArchivo = Path.GetFileName(ArchivoExcel);

        //    Microsoft.Office.Interop.Excel.Application xlApp;
        //    Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
        //    Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
        //    Microsoft.Office.Interop.Excel.Range range;

        //    xlApp = new Microsoft.Office.Interop.Excel.Application();
        //    xlWorkBook = xlApp.Workbooks.Open(ArchivoExcel, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
        //    if (xlWorkBook.Worksheets.Count > 0)
        //    {
        //        xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

        //        try
        //        {
        //            range = xlWorkSheet.UsedRange;

        //            int rows = range.Rows.Count;
        //            int cols = range.Columns.Count;

        //            if (rows > 1) //Si es que la planilla contiene titulos
        //            {
        //                for (int i = 2; i < rows + 1; i++)
        //                {
        //                    if (Convert.ToInt32(xlWorkSheet.Cells[i, 1].Value) > 0)
        //                    {
        //                        solicitudesDespacho = new SolicitudesDespacho();

        //                        solicitudesDespacho.NroOrden = Convert.ToInt32(xlWorkSheet.Cells[i, 1].Value);
        //                        solicitudesDespacho.NroVenta = Convert.ToInt32(xlWorkSheet.Cells[i, 2].Value);
        //                        solicitudesDespacho.Courier = xlWorkSheet.Cells[i, 3].Value;
        //                        solicitudesDespacho.SKU = Convert.ToString(xlWorkSheet.Cells[i, 4].Value);
        //                        solicitudesDespacho.DescProducto = xlWorkSheet.Cells[i, 5].Value;
        //                        solicitudesDespacho.Cantidad = Convert.ToInt32(xlWorkSheet.Cells[i, 6].Value);
        //                        solicitudesDespacho.Rut = xlWorkSheet.Cells[i, 7].Value;
        //                        solicitudesDespacho.Nombre = xlWorkSheet.Cells[i, 8].Value;
        //                        solicitudesDespacho.Comuna = xlWorkSheet.Cells[i, 9].Value;
        //                        solicitudesDespacho.Direccion = xlWorkSheet.Cells[i, 10].Value;
        //                        solicitudesDespacho.Telefono = Convert.ToString(xlWorkSheet.Cells[i, 11].Value);
        //                        solicitudesDespacho.Observacion = xlWorkSheet.Cells[i, 12].Value;
        //                        solicitudesDespacho.FechaVenta = xlWorkSheet.Cells[i, 13].Value;
        //                        solicitudesDespacho.FechaRetiro = xlWorkSheet.Cells[i, 14].Value;
        //                        solicitudesDespacho.RutReceptor = xlWorkSheet.Cells[i, 15].Value;
        //                        solicitudesDespacho.NombreReceptor = xlWorkSheet.Cells[i, 16].Value;
        //                        solicitudesDespacho.TelefonoReceptor = Convert.ToString(xlWorkSheet.Cells[i, 17].Value);
        //                        solicitudesDespacho.DireccionReceptor = xlWorkSheet.Cells[i, 18].Value;

        //                        retorno.Add(solicitudesDespacho);
        //                    }
        //                }
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            Log.grabaLog("Datos", "LeerExcel", "Archivo: " + NombreArchivo + "; Error: " + ex.Message);
        //            throw new Exception("Error al leer la información del archivo.");
        //        }
        //        finally
        //        {
        //            Marshal.ReleaseComObject(xlWorkSheet);
        //        }

        //        if (retorno.Count == 0)
        //        {
        //            Log.grabaLog("Datos", "LeerExcel", "Archivo: " + NombreArchivo + "; Error: no hay filas para procesar");
        //            throw new Exception("Error, no hay filas para procesar");
        //        }
        //    }
        //    else
        //    {
        //        Log.grabaLog("Datos", "LeerExcel", "Archivo: " + NombreArchivo + "; Error: no hay hojas para procesar");
        //        throw new Exception("Error, no hay hojas para procesar");
        //    }

        //    xlWorkBook.Close(true, null, null);
        //    xlApp.Quit();

        //    Marshal.ReleaseComObject(xlWorkBook);
        //    Marshal.ReleaseComObject(xlApp);

        //    return retorno;
        //}

    }
}
