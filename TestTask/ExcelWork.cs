using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using System.IO;
using System.Data;
using ExcelDataReader;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace TestTask
{
    public static class ExcelWork
    {
        const string filePath = "/Files/Отчет";
        static string extFile = ".xls";

        public static void SaveFile(IFormFile uploadFile, string webRoot)
        {
            if (uploadFile != null)
            {
                using (var fileStream = new FileStream(webRoot + filePath + extFile, FileMode.Create))
                {
                    uploadFile.CopyTo(fileStream);
                }
            }
        }

        //Филтрация ошибочных строк
        public static Dictionary<int, DataRow> Filter(DataTable source)
        {
            var result = new Dictionary<int, DataRow>();
            if (source.Columns.Count < 11)
            {
                return result;
            }
            int[] indexes = new int[] { 0, 1, 2, 3, 4, 5 };
            for (int i = 0; i < source.Rows.Count; i++)
            {
                if (IsX(source.Rows[i], 10))
                {
                    bool isError = true;
                    foreach (int col in indexes)
                    {
                        if (IsX(source.Rows[i], col))
                        {
                            isError = false;
                            break;
                        }
                    }
                    if (isError)
                    {
                        result.Add(i, source.Rows[i]);
                    }
                }
            }
            return result;
        }

        static bool IsX(DataRow sourceRow, int index)
        {
            return sourceRow[index].ToString().ToLower() == "x" || //en
                   sourceRow[index].ToString().ToLower() == "х";   //ru
        }

        public static DataTable ReadExcel(IFormFile uploadFile)
        {
            if (uploadFile == null)
            {
                return null;
            }
            Stream stream = uploadFile.OpenReadStream();
            IExcelDataReader reader = null;
            var fileName = uploadFile.FileName;
            if (fileName.ToLower().EndsWith(".xls"))
            {
                reader = ExcelReaderFactory.CreateBinaryReader(stream);
                extFile = ".xls";
            }
            else if (fileName.ToLower().EndsWith(".xlsx"))
            {
                reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                extFile = ".xlsx";
            }
            else
            {
                return null;
            }

            DataSet result = reader.AsDataSet();
            reader.Close();
            return result.Tables[0];
        }

        //Создание excel - отчета
        public static string ExcelReport(List<int> selected, string webRoot)
        {
            IWorkbook workbook;
            string fullPath = webRoot + filePath + extFile;
            using (var stream = new FileStream(fullPath, FileMode.Open, FileAccess.ReadWrite))
            {
                ISheet sheet;
                if (extFile == ".xls")
                {
                    workbook = new HSSFWorkbook(stream);
                    sheet = workbook.GetSheetAt(0);
                }
                else if (extFile == ".xlsx")
                {
                    workbook = new XSSFWorkbook(stream);
                    sheet = workbook.GetSheetAt(0);
                }
                else
                {
                    return "";
                }
                int lastRowNum = sheet.LastRowNum + 1; ;
                int deletedRows = 0;
                int startHeader = 6;
                for (int i = startHeader; i < lastRowNum; i++) 
                {
                    var cells = sheet.GetRow(i).Cells;
                    if (selected.Contains(i + deletedRows) || isHeaderCell(cells)) 
                    {
                        continue;
                    }
                    sheet.ShiftRows(i + 1, lastRowNum, -1);
                    lastRowNum--;
                    i--;
                    deletedRows++;
                }
            }
            using (FileStream stream = new FileStream(fullPath, FileMode.Open, FileAccess.ReadWrite))
            {
                workbook.Write(stream);
            }
            return fullPath;
        }

        //проверка на наличие объедененной ячейки
        static bool isHeaderCell(List<ICell> cells) 
        {
            if (cells.Count == 0) 
            {
                return false;
            }
            if (cells.First().CellType == CellType.String && cells.First().StringCellValue == "") 
            {
                return false;
            }
            for (int i = 1; i < cells.Count; i++) 
            {
                if (cells[i].CellType == CellType.String && cells[i].StringCellValue != "") 
                {
                    return false;
                }
            }
            return true;
        }
    }
}
