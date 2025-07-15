using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using System.Linq;

public class ExcelHelper
{
    public static void RemoverLinhasAntigas(string filePath)
    {
        if (File.Exists(filePath))
        {
            DateTime sevenMonthsAgo = DateTime.Now.AddMonths(-7);

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = rowCount; row >= 2; row--)
                {
                    var dataCellValue = worksheet.Cells[row, 1].Text;

                    if (DateTime.TryParse(dataCellValue, out DateTime dataDaTarefa))
                    {
                        if (dataDaTarefa < sevenMonthsAgo)
                        {
                            worksheet.DeleteRow(row);
                        }
                    }
                }

                package.Save();
            }
        }
        else
        {
            MessageBox.Show("O arquivo não foi encontrado. Verifique os eguinte caminho \n\n \\\\marconi\\COMPANY SHARED FOLDER\\OFELIZ\\OFM\\3.SP\\7.DT\\1.Técnico\\5.CTS");
        }
    }



}

