using CreateExcelSheet.Data;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.DataValidation;
using System.Data;
using System.Drawing;

namespace CreateExcelSheet.Services
{
    public class ExcelGenerationService : IExcelGenerationService
    {
        public async Task<MemoryStream> GenerateStudentList(List<Students> list)
        {
            MemoryStream memoryStream = new();
            await Task.Run(() =>
            {
                DataTable studentTable = GetDataTable(list);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelPackage package = new();
                ExcelWorksheet excelWorkSheet = package.Workbook.Worksheets.Add("STUDENTS");
                int totalRowHeader = studentTable.Rows.Count + 2;
                int totalRow = studentTable.Rows.Count;
                int totalColumn = studentTable.Columns.Count;
                for (int i = 1; i <= totalColumn; i++)
                    excelWorkSheet = Header(excelWorkSheet, i, studentTable.Columns[i - 1].ColumnName);
                for (int j = 0; j < totalRow; j++)
                {
                    for (int k = 0; k < totalColumn; k++)
                    {
                        var value = studentTable.Rows[j].ItemArray[k];
                        bool isByte = byte.TryParse(value?.ToString(), out byte byteValue);
                        excelWorkSheet.Cells[j + 3, k + 1].Value = isByte == true ? byteValue : value;
                    }
                }
                ExcelRange schoolName;
                schoolName = excelWorkSheet.Cells["A1:D1"];
                schoolName.Merge = true;
                schoolName.Value = "InterFlip Int'l School, Lagos";
                schoolName.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                schoolName.Style.Font.Size = 13;
                schoolName.Style.Font.Bold = true;
                schoolName.Style.Fill.PatternType = ExcelFillStyle.Solid;
                schoolName.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                schoolName.Style.Border.BorderAround(ExcelBorderStyle.Medium);
                CaExamProp(excelWorkSheet, totalRowHeader);
                memoryStream = LockBook(package);
            });
            return memoryStream;
        }

        private ExcelWorksheet Header(ExcelWorksheet excelWorkSheet, int i, string ColumnName)
        {
            excelWorkSheet.Cells[2, i].Value = ColumnName;
            excelWorkSheet.Cells[2, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            excelWorkSheet.Cells[2, i].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            excelWorkSheet.Cells[2, i].Style.Font.Bold = true;
            excelWorkSheet.Cells[2, i].Style.Fill.PatternType = ExcelFillStyle.Solid;
            excelWorkSheet.Cells[2, i].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
            excelWorkSheet.Cells[2, i].Style.Border.BorderAround(ExcelBorderStyle.Medium);
            excelWorkSheet.Cells[2, i].Style.WrapText = true;
            return excelWorkSheet;
        }

        private void CaExamProp(ExcelWorksheet excelWorkSheet, int row)
        {
            string colIndexCa = "B" + row;
            string colIndexExam = "C" + row;
            string colIndexSum = "D" + row;
            excelWorkSheet.Column(1).Width = 40;
            excelWorkSheet.Cells["B3" + ":" + colIndexCa].Style.Locked = false;
            excelWorkSheet.Cells["C3" + ":" + colIndexExam].Style.Locked = false;

            var rngCA = excelWorkSheet.DataValidations.AddIntegerValidation("B3" + ":" + colIndexCa);
            rngCA.ErrorStyle = ExcelDataValidationWarningStyle.stop;
            rngCA.Error = "Value not a number or greater than maximum score of 40";
            rngCA.ErrorTitle = "InterFlip Int'l School - Input Error";
            rngCA.ShowErrorMessage = true;
            rngCA.Operator = ExcelDataValidationOperator.between;
            rngCA.Formula.Value = 0;
            rngCA.Formula2.Value = 40;
            rngCA.AllowBlank = true;

            var rngExam = excelWorkSheet.DataValidations.AddIntegerValidation("C3" + ":" + colIndexExam);
            rngExam.ErrorStyle = ExcelDataValidationWarningStyle.stop;
            rngExam.Error = "Value not a number or greater than maximum score of 60";
            rngExam.ErrorTitle = "InterFlip Int'l School - Input Error";
            rngExam.ShowErrorMessage = true;
            rngExam.Operator = ExcelDataValidationOperator.between;
            rngExam.Formula.Value = 0;
            rngExam.Formula2.Value = 60;
            rngExam.AllowBlank = true;
            excelWorkSheet.Cells["D3" + ":" + colIndexSum].Formula = "=SUM(B3:C3)";
            ExcelColumn columns = excelWorkSheet.Column(5);
            columns.ColumnMax = 16384;
            columns.Hidden = true;
            excelWorkSheet.View.FreezePanes(4, 5);
            excelWorkSheet.Protection.SetPassword("2222222");
        }

        private MemoryStream LockBook(ExcelPackage package)
        {
            package.Workbook.Protection.LockStructure = true;
            package.Workbook.Protection.SetPassword("11111111");
            return new MemoryStream(package.GetAsByteArray());
        }

        private DataTable GetDataTable(List<Students> students)
        {
            DataTable tbl = new();
            tbl.Columns.Add("NAME");
            tbl.Columns.Add("CA (40)");
            tbl.Columns.Add("EXAM (60)");
            tbl.Columns.Add("TOTAL (100)");
            tbl.Columns.Add("Id");
            for (int i = 0; i < students.Count(); i++)
            {
                var dataRow = students.ElementAt(i);
                tbl.Rows.Add();
                tbl.Rows[i][0] = dataRow.Name;
                tbl.Rows[i][1] = dataRow.Ca;
                tbl.Rows[i][2] = dataRow.Exam;
                tbl.Rows[i][3] = 0;
                tbl.Rows[i][4] = dataRow.Id;
            }
            return tbl;
        }
    }
}
