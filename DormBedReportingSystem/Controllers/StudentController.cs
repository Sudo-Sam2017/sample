using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.AspNetCore.Mvc;
using DormBedReportingSystem.Models;
using DormBedReportingSystem.ViewModels;
using System.IO;
using ExcelDataReader;
using System.Data;
using OfficeOpenXml;

namespace DormBedReportingSystem.Controllers
{
    public class StudentController : Controller
    {
        public ActionResult Index()
        {
            var fileName = @"C:\Users\bigra\Downloads\CENTERsample.xlsx";
            List<Students> studentsList = new List<Students>();
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            FileStream stream1 = System.IO.File.Open(fileName, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream1);
            var dataset = excelReader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true,
                    ReadHeaderRow = excelReader =>
                    {
                        while (excelReader.Read())
                        {
                            studentsList.Add(new Students
                            {
                                RES_STUD_FEMALE = excelReader.GetValue(3).ToString(),
                                RES_STUD_MALE = excelReader.GetValue(4).ToString(),
                                TOTAL_RES_TARGET = excelReader.GetValue(5).ToString(),
                                NRES_STUD_FEMALE = excelReader.GetValue(6).ToString(),
                                NRES_STUD_MALE = excelReader.GetValue(7).ToString(),
                                TOTAL_NRES_TARGET = excelReader.GetValue(8).ToString(),
                                TOTAL_FEM_TARGET = excelReader.GetValue(9).ToString(),
                                TOTAL_MALE_TARGET = excelReader.GetValue(10).ToString(),
                                TOTAL_TARGET = excelReader.GetValue(11).ToString(),
                                AS_OF_TARGET = excelReader.GetValue(12).ToString(),
                                RES_STUD_FEMALE_OBS = excelReader.GetValue(13).ToString(),
                                RES_STUD_MALE_OBS = excelReader.GetValue(14).ToString(),
                                TOTAL_RES_ACTUAL = excelReader.GetValue(15).ToString(),
                                NRES_STUD_FEMALE_OBS = excelReader.GetValue(16).ToString(),
                                NRES_STUD_MALE_OBS = excelReader.GetValue(17).ToString(),
                                TOTAL_NRES_ACTUAL = excelReader.GetValue(18).ToString(),
                                TOTAL_FEM_ACTUAL = excelReader.GetValue(19).ToString(),
                                TOTAL_MALE_ACTUAL = excelReader.GetValue(20).ToString(),
                                TOTAL_ACTUAL = excelReader.GetValue(21).ToString(),
                                AS_OF_ACTUAL = excelReader.GetValue(22).ToString()

                            });
                        }
                        excelReader.Close();
                    }
                }
            });

            var fileName2 = @"C:\Users\bigra\Downloads\sampleBuildings.xlsx";
            List<Beds> bedsList = new List<Beds>();
            List<ReadEditRow> readRow = new List<ReadEditRow>();

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            FileStream stream2 = System.IO.File.Open(fileName2, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader2 = ExcelReaderFactory.CreateOpenXmlReader(stream2);
            var dataset2 = excelReader2.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true,
                    ReadHeaderRow = excelReader2 =>
                    {
                        while (excelReader2.Read())
                        {
                           
                            bedsList.Add(new Beds
                            {
                                REGION = excelReader2.GetValue(1).ToString(),
                                REGION_NAME = excelReader2.GetValue(2).ToString(),
                                CENTER_NAME = excelReader2.GetValue(3).ToString(),
                                BLDG_NAME = excelReader2.GetValue(4).ToString(),
                                WING_NAME = excelReader2.GetValue(5).ToString(),
                                CAP_1BED = Int32.Parse(excelReader2.GetValue(6).ToString()),
                                CAP_2BEDS = Int32.Parse(excelReader2.GetValue(7).ToString()),
                                CAP_3BEDS = Int32.Parse(excelReader2.GetValue(8).ToString()),
                                CAP_4BEDS = Int32.Parse(excelReader2.GetValue(9).ToString()),
                                CAP_5BEDS = Int32.Parse(excelReader2.GetValue(10).ToString()),
                                CAP_6BEDS = Int32.Parse(excelReader2.GetValue(11).ToString()),
                                CAP_7BEDS = Int32.Parse(excelReader2.GetValue(12).ToString()),
                                CAP_8BEDS = Int32.Parse(excelReader2.GetValue(13).ToString()),
                                CAP_9BEDS = Int32.Parse(excelReader2.GetValue(14).ToString()),
                                CAP_10BEDS = Int32.Parse(excelReader2.GetValue(15).ToString()),
                                OCC_1BED = Int32.Parse(excelReader2.GetValue(16).ToString()),
                                OCC_2BEDS = Int32.Parse(excelReader2.GetValue(17).ToString()),
                                OCC_3BEDS = Int32.Parse(excelReader2.GetValue(18).ToString()),
                                OCC_4BEDS = Int32.Parse(excelReader2.GetValue(19).ToString()),
                                OCC_5BEDS = Int32.Parse(excelReader2.GetValue(20).ToString()),
                                OCC_6BEDS = Int32.Parse(excelReader2.GetValue(21).ToString()),
                                OCC_7BEDS = Int32.Parse(excelReader2.GetValue(22).ToString()),
                                OCC_8BEDS = Int32.Parse(excelReader2.GetValue(23).ToString()),
                                OCC_9BEDS = Int32.Parse(excelReader2.GetValue(24).ToString()),
                                OCC_10BEDS = Int32.Parse(excelReader2.GetValue(25).ToString()),
                                VB_OBS = Int32.Parse(excelReader2.GetValue(26).ToString()),
                                VB_REPAIR = Int32.Parse(excelReader2.GetValue(27).ToString()),
                                VB_NOFURNITURE = Int32.Parse(excelReader2.GetValue(28).ToString()),
                                CAP_TOTAL = (Int32.Parse(excelReader2.GetValue(6).ToString()) * 1) +
                                            (Int32.Parse(excelReader2.GetValue(7).ToString()) * 2) +
                                            (Int32.Parse(excelReader2.GetValue(8).ToString()) * 3) +
                                            (Int32.Parse(excelReader2.GetValue(9).ToString()) * 4) +
                                            (Int32.Parse(excelReader2.GetValue(10).ToString()) * 5) +
                                            (Int32.Parse(excelReader2.GetValue(11).ToString()) * 6) +
                                            (Int32.Parse(excelReader2.GetValue(12).ToString()) * 7) +
                                            (Int32.Parse(excelReader2.GetValue(13).ToString()) * 8) +
                                            (Int32.Parse(excelReader2.GetValue(14).ToString()) * 9) +
                                            (Int32.Parse(excelReader2.GetValue(15).ToString()) * 10),
                                OCC_TOTAL = (Int32.Parse(excelReader2.GetValue(16).ToString()) * 1) +
                                            (Int32.Parse(excelReader2.GetValue(17).ToString()) * 2) +
                                            (Int32.Parse(excelReader2.GetValue(18).ToString()) * 3) +
                                            (Int32.Parse(excelReader2.GetValue(19).ToString()) * 4) +
                                            (Int32.Parse(excelReader2.GetValue(20).ToString()) * 5) +
                                            (Int32.Parse(excelReader2.GetValue(21).ToString()) * 6) +
                                            (Int32.Parse(excelReader2.GetValue(22).ToString()) * 7) +
                                            (Int32.Parse(excelReader2.GetValue(23).ToString()) * 8) +
                                            (Int32.Parse(excelReader2.GetValue(24).ToString()) * 9) +
                                            (Int32.Parse(excelReader2.GetValue(25).ToString()) * 10)
                            });


                            
                        }
                        excelReader2.Close();
                    }
                }
            });

            List<Beds> sortedBedsList = bedsList.OrderByDescending(a => a.BLDG_NAME).ToList();

            var StudentBedViewModel = new StudentBedViewModel()
            {
                Students = studentsList,
                Beds = bedsList,
                ReadEditRow = readRow,
            };



            return View(StudentBedViewModel);
        }

        public void AddNewRow(AddedRow addRow)
        {            

            string Region = "1";
            string RegionName = "Boston";
            string CenterID = "00205";
            string CenterName = "ONEONTA";
            string BuildingName = addRow.BLDG_NAME;
            string WingName = addRow.WING_NAME;
            int Cap_1Bed = addRow.CAP_1BED;
            int Cap_2Beds = addRow.CAP_2BEDS;
            int Cap_3Beds = addRow.CAP_3BEDS;
            int Cap_4Beds = addRow.CAP_4BEDS;
            int Cap_5Beds = addRow.CAP_5BEDS;
            int Cap_6Beds = addRow.CAP_6BEDS;
            int Cap_7Beds = addRow.CAP_7BEDS;
            int Cap_8Beds = addRow.CAP_8BEDS;
            int Cap_9Beds = addRow.CAP_9BEDS;
            int Cap_10Beds = addRow.CAP_10BEDS;
            int Occ_1Bed = addRow.OCC_1BED;
            int Occ_2Beds = addRow.OCC_2BEDS;
            int Occ_3Beds = addRow.OCC_3BEDS;
            int Occ_4Beds = addRow.OCC_4BEDS;
            int Occ_5Beds = addRow.OCC_5BEDS;
            int Occ_6Beds = addRow.OCC_6BEDS;
            int Occ_7Beds = addRow.OCC_7BEDS;
            int Occ_8Beds = addRow.OCC_8BEDS;
            int Occ_9Beds = addRow.OCC_9BEDS;
            int Occ_10Beds = addRow.OCC_10BEDS;
            int Vb_Obs = 0;
            int Vb_Repair = 0;
            int Vb_NoFurniture = 0;
            string bedComments = addRow.BedComments;


            FileInfo file = new FileInfo(@"C:\Users\bigra\Downloads\sampleBuildings.xlsx");
            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                ExcelWorkbook excelWorkBook = excelPackage.Workbook;
                ExcelWorksheet excelWorksheet = excelWorkBook.Worksheets.First();

                int lastRow = excelWorksheet.Cells.Where(cell => !string.IsNullOrEmpty(cell.Value?.ToString() ?? string.Empty)).LastOrDefault().End.Row;
                int row = lastRow + 1;
                excelWorksheet.Cells[row, 1].Value = Region;
                excelWorksheet.Cells[row, 2].Value = RegionName;
                excelWorksheet.Cells[row, 3].Value = CenterID;
                excelWorksheet.Cells[row, 4].Value = CenterName;
                excelWorksheet.Cells[row, 5].Value = BuildingName;
                excelWorksheet.Cells[row, 6].Value = WingName;
                excelWorksheet.Cells[row, 7].Value = Cap_1Bed;
                excelWorksheet.Cells[row, 8].Value = Cap_2Beds;
                excelWorksheet.Cells[row, 9].Value = Cap_3Beds;
                excelWorksheet.Cells[row, 10].Value = Cap_4Beds;
                excelWorksheet.Cells[row, 11].Value = Cap_5Beds;
                excelWorksheet.Cells[row, 12].Value = Cap_6Beds;
                excelWorksheet.Cells[row, 13].Value = Cap_7Beds;
                excelWorksheet.Cells[row, 14].Value = Cap_8Beds;
                excelWorksheet.Cells[row, 15].Value = Cap_9Beds;
                excelWorksheet.Cells[row, 16].Value = Cap_10Beds;
                excelWorksheet.Cells[row, 17].Value = Occ_1Bed;
                excelWorksheet.Cells[row, 18].Value = Occ_2Beds;
                excelWorksheet.Cells[row, 19].Value = Occ_3Beds;
                excelWorksheet.Cells[row, 20].Value = Occ_4Beds;
                excelWorksheet.Cells[row, 21].Value = Occ_5Beds;
                excelWorksheet.Cells[row, 22].Value = Occ_6Beds;
                excelWorksheet.Cells[row, 23].Value = Occ_7Beds;
                excelWorksheet.Cells[row, 24].Value = Occ_8Beds;
                excelWorksheet.Cells[row, 25].Value = Occ_9Beds;
                excelWorksheet.Cells[row, 26].Value = Occ_10Beds;
                excelWorksheet.Cells[row, 27].Value = Vb_Obs;
                excelWorksheet.Cells[row, 28].Value = Vb_Repair;
                excelWorksheet.Cells[row, 29].Value = Vb_NoFurniture;
                excelWorksheet.Cells[row, 42].Value = bedComments;

                ExcelRange range = excelWorksheet.Cells[1, 1, 1000, 1000];
                range.Sort(4, true);

                excelPackage.Save();
            }
            Response.Redirect(Url.Action("Index", "Student", null));
        }

         public ActionResult Add()
         {
            var fileName2 = @"C:\Users\bigra\Downloads\sampleBuildings.xlsx";
            List<Beds> bedsList = new List<Beds>();
            List<Beds> sortedBeds = bedsList.OrderBy(o => o.BLDG_NAME).ToList();

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            FileStream stream2 = System.IO.File.Open(fileName2, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader2 = ExcelReaderFactory.CreateOpenXmlReader(stream2);

            var dataset2 = excelReader2.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true,
                    ReadHeaderRow = excelReader2 =>
                    {
                        while (excelReader2.Read())
                        {
                            sortedBeds.Add(new Beds
                            {
                                REGION = excelReader2.GetValue(1).ToString(),
                                REGION_NAME = excelReader2.GetValue(2).ToString(),
                                CENTER_NAME = excelReader2.GetValue(3).ToString(),
                                BLDG_NAME = excelReader2.GetValue(4).ToString(),
                                WING_NAME = excelReader2.GetValue(5).ToString(),
                                CAP_1BED = Int32.Parse(excelReader2.GetValue(6).ToString()),
                                CAP_2BEDS = Int32.Parse(excelReader2.GetValue(7).ToString()),
                                CAP_3BEDS = Int32.Parse(excelReader2.GetValue(8).ToString()),
                                CAP_4BEDS = Int32.Parse(excelReader2.GetValue(9).ToString()),
                                CAP_5BEDS = Int32.Parse(excelReader2.GetValue(10).ToString()),
                                CAP_6BEDS = Int32.Parse(excelReader2.GetValue(11).ToString()),
                                CAP_7BEDS = Int32.Parse(excelReader2.GetValue(12).ToString()),
                                CAP_8BEDS = Int32.Parse(excelReader2.GetValue(13).ToString()),
                                CAP_9BEDS = Int32.Parse(excelReader2.GetValue(14).ToString()),
                                CAP_10BEDS = Int32.Parse(excelReader2.GetValue(15).ToString()),
                                OCC_1BED = Int32.Parse(excelReader2.GetValue(16).ToString()),
                                OCC_2BEDS = Int32.Parse(excelReader2.GetValue(17).ToString()),
                                OCC_3BEDS = Int32.Parse(excelReader2.GetValue(18).ToString()),
                                OCC_4BEDS = Int32.Parse(excelReader2.GetValue(19).ToString()),
                                OCC_5BEDS = Int32.Parse(excelReader2.GetValue(20).ToString()),
                                OCC_6BEDS = Int32.Parse(excelReader2.GetValue(21).ToString()),
                                OCC_7BEDS = Int32.Parse(excelReader2.GetValue(22).ToString()),
                                OCC_8BEDS = Int32.Parse(excelReader2.GetValue(23).ToString()),
                                OCC_9BEDS = Int32.Parse(excelReader2.GetValue(24).ToString()),
                                OCC_10BEDS = Int32.Parse(excelReader2.GetValue(25).ToString()),
                                CAP_TOTAL = (Int32.Parse(excelReader2.GetValue(6).ToString()) * 1) +
                                            (Int32.Parse(excelReader2.GetValue(7).ToString()) * 2) +
                                            (Int32.Parse(excelReader2.GetValue(8).ToString()) * 3) +
                                            (Int32.Parse(excelReader2.GetValue(9).ToString()) * 4) +
                                            (Int32.Parse(excelReader2.GetValue(10).ToString()) * 5) +
                                            (Int32.Parse(excelReader2.GetValue(11).ToString()) * 6) +
                                            (Int32.Parse(excelReader2.GetValue(12).ToString()) * 7) +
                                            (Int32.Parse(excelReader2.GetValue(13).ToString()) * 8) +
                                            (Int32.Parse(excelReader2.GetValue(14).ToString()) * 9) +
                                            (Int32.Parse(excelReader2.GetValue(15).ToString()) * 10),
                                OCC_TOTAL = (Int32.Parse(excelReader2.GetValue(16).ToString()) * 1) +
                                            (Int32.Parse(excelReader2.GetValue(17).ToString()) * 2) +
                                            (Int32.Parse(excelReader2.GetValue(18).ToString()) * 3) +
                                            (Int32.Parse(excelReader2.GetValue(19).ToString()) * 4) +
                                            (Int32.Parse(excelReader2.GetValue(20).ToString()) * 5) +
                                            (Int32.Parse(excelReader2.GetValue(21).ToString()) * 6) +
                                            (Int32.Parse(excelReader2.GetValue(22).ToString()) * 7) +
                                            (Int32.Parse(excelReader2.GetValue(23).ToString()) * 8) +
                                            (Int32.Parse(excelReader2.GetValue(24).ToString()) * 9) +
                                            (Int32.Parse(excelReader2.GetValue(25).ToString()) * 10)
                            });
                        }
                        excelReader2.Close();
                    }
                }
            });
            AddedRow addrow = new AddedRow();

            var StudentBedViewModel = new StudentBedViewModel()
            {
                Beds = sortedBeds,
                AddRow = addrow
            };

            return View(StudentBedViewModel);
         }


        public ActionResult ReadEdit(int id)
        {

            ReadEditRow indexRow = new ReadEditRow();
            List<ReadEditRow> sortedEditList = new List<ReadEditRow>();

            FileInfo file = new FileInfo(@"C:\Users\bigra\Downloads\sampleBuildings.xlsx");
            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                ExcelWorkbook excelWorkBook = excelPackage.Workbook;
                ExcelWorksheet excelWorksheet = excelWorkBook.Worksheets.First();

                int lastRow = excelWorksheet.Cells.Where(cell => !string.IsNullOrEmpty(cell.Value?.ToString() ?? string.Empty)).LastOrDefault().End.Row;
                int row = id;

                indexRow.id = row;
                indexRow.REGION = excelWorksheet.Cells[row, 1].Value.ToString();
                indexRow.REGION_NAME = excelWorksheet.Cells[row, 2].Value.ToString();
                indexRow.CENTER_ID = excelWorksheet.Cells[row, 3].Value.ToString();
                indexRow.CENTER_NAME = excelWorksheet.Cells[row, 4].Value.ToString();
                indexRow.BLDG_NAME = excelWorksheet.Cells[row, 5].Value.ToString();
                indexRow.WING_NAME = excelWorksheet.Cells[row, 6].Value.ToString();
                indexRow.CAP_1BED = excelWorksheet.Cells[row, 7].Value.ToString();
                indexRow.CAP_2BEDS = excelWorksheet.Cells[row, 8].Value.ToString();
                indexRow.CAP_3BEDS = excelWorksheet.Cells[row, 9].Value.ToString();
                indexRow.CAP_4BEDS = excelWorksheet.Cells[row, 10].Value.ToString();
                indexRow.CAP_5BEDS = excelWorksheet.Cells[row, 11].Value.ToString();
                indexRow.CAP_6BEDS = excelWorksheet.Cells[row, 12].Value.ToString();
                indexRow.CAP_7BEDS = excelWorksheet.Cells[row, 13].Value.ToString();
                indexRow.CAP_8BEDS = excelWorksheet.Cells[row, 14].Value.ToString();
                indexRow.CAP_9BEDS = excelWorksheet.Cells[row, 15].Value.ToString();
                indexRow.CAP_10BEDS = excelWorksheet.Cells[row, 16].Value.ToString();
                indexRow.OCC_1BED = excelWorksheet.Cells[row, 17].Value.ToString();
                indexRow.OCC_2BEDS = excelWorksheet.Cells[row, 18].Value.ToString();
                indexRow.OCC_3BEDS = excelWorksheet.Cells[row, 19].Value.ToString();
                indexRow.OCC_4BEDS = excelWorksheet.Cells[row, 20].Value.ToString();
                indexRow.OCC_5BEDS = excelWorksheet.Cells[row, 21].Value.ToString();
                indexRow.OCC_6BEDS = excelWorksheet.Cells[row, 22].Value.ToString();
                indexRow.OCC_7BEDS = excelWorksheet.Cells[row, 23].Value.ToString();
                indexRow.OCC_8BEDS = excelWorksheet.Cells[row, 24].Value.ToString();
                indexRow.OCC_9BEDS = excelWorksheet.Cells[row, 25].Value.ToString();
                indexRow.OCC_10BEDS = excelWorksheet.Cells[row, 26].Value.ToString();
                indexRow.BedComments = excelWorksheet.Cells[row, 42].Text;
                excelPackage.Save();
            }

            List<ReadEditRow> sortedReadEditRow = sortedEditList.OrderByDescending(a => a.BLDG_NAME).ToList();

            var editRowViewModel = new ReadEditViewModel()
            {
                ReadEditRow = sortedReadEditRow
            };
            return View(indexRow);
        }

        public void Update(Beds bedEntry, int id)
        {
            

            FileInfo file = new FileInfo(@"C:\Users\bigra\Downloads\sampleBuildings.xlsx");
            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                ExcelWorkbook excelWorkBook = excelPackage.Workbook;
                ExcelWorksheet excelWorksheet = excelWorkBook.Worksheets.First();
                int lastRow = excelWorksheet.Cells.Where(cell => !string.IsNullOrEmpty(cell.Value?.ToString() ?? string.Empty)).LastOrDefault().End.Row;
                int row = id;
                excelWorksheet.Cells[row, 7].Value = bedEntry.CAP_1BED;
                excelWorksheet.Cells[row, 8].Value = bedEntry.CAP_2BEDS;
                excelWorksheet.Cells[row, 9].Value = bedEntry.CAP_3BEDS;
                excelWorksheet.Cells[row, 10].Value = bedEntry.CAP_4BEDS;
                excelWorksheet.Cells[row, 11].Value = bedEntry.CAP_5BEDS;
                excelWorksheet.Cells[row, 12].Value = bedEntry.CAP_6BEDS;
                excelWorksheet.Cells[row, 13].Value = bedEntry.CAP_7BEDS;
                excelWorksheet.Cells[row, 14].Value = bedEntry.CAP_8BEDS;
                excelWorksheet.Cells[row, 15].Value = bedEntry.CAP_9BEDS;
                excelWorksheet.Cells[row, 16].Value = bedEntry.CAP_10BEDS;
                excelWorksheet.Cells[row, 17].Value = bedEntry.OCC_1BED;
                excelWorksheet.Cells[row, 18].Value = bedEntry.OCC_2BEDS;
                excelWorksheet.Cells[row, 19].Value = bedEntry.OCC_3BEDS;
                excelWorksheet.Cells[row, 20].Value = bedEntry.OCC_4BEDS;
                excelWorksheet.Cells[row, 21].Value = bedEntry.OCC_5BEDS;
                excelWorksheet.Cells[row, 22].Value = bedEntry.OCC_6BEDS;
                excelWorksheet.Cells[row, 23].Value = bedEntry.OCC_7BEDS;
                excelWorksheet.Cells[row, 24].Value = bedEntry.OCC_8BEDS;
                excelWorksheet.Cells[row, 25].Value = bedEntry.OCC_9BEDS;
                excelWorksheet.Cells[row, 26].Value = bedEntry.OCC_10BEDS;
                excelWorksheet.Cells[row, 42].Value = bedEntry.BedComments;
                ExcelRange range = excelWorksheet.Cells[1, 1, 1000, 1000];
                range.Sort(4, true);
                excelPackage.Save();

            }

            Response.Redirect(Url.Action("Index", "Student", null));
        }

        public void DeleteRow(int id)
        {
            FileInfo file = new FileInfo(@"C:\Users\bigra\Downloads\sampleBuildings.xlsx");
            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                ExcelWorkbook excelWorkBook = excelPackage.Workbook;
                ExcelWorksheet excelWorksheet = excelWorkBook.Worksheets.First();
                int lastRow = excelWorksheet.Cells.Where(cell => !string.IsNullOrEmpty(cell.Value?.ToString() ?? string.Empty)).LastOrDefault().End.Row;
                int row = id;
                excelWorksheet.DeleteRow(row);
                excelPackage.Save();
            }
            Response.Redirect(Url.Action("Index", "Student", null));
        }

    }
}
    
