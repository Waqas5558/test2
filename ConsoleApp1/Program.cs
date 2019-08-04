using Microsoft.EntityFrameworkCore;
using System;
using System.Data;
using System.IO;
using System.Linq;

namespace ConsoleApp1
{
    class Program
    {
        private MyDbContext dbContext;
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            GetDataTableFromExcel(@"F:\Backup of Exceptions.xlk");
        }
        public static void GetDataTableFromExcel(string path, bool hasHeader = true)
        {
            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                using (var stream = File.OpenRead(path))
                {
                    pck.Load(stream);
                }
                var ws = pck.Workbook.Worksheets.First();
                DataTable tbl = new DataTable();
                string range = "A1:P5";
                foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Cells[range].Columns])// Dimension.End.Column])
                {
                    tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                }
                var startRow = hasHeader ? 5 : 1;
                for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                {
                    var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Cells[range].Columns]; //ws.Dimension.End.Column];
                    DataRow row = tbl.Rows.Add();
                    foreach (var cell in wsRow)
                    {
                        try
                        {
                            row[cell.Start.Column - 1] = cell.Text;
                        }
                        catch (System.Exception)
                        {
                            continue;
                        }

                    }
                }
                //return tbl;

                // On all tables' rows
                foreach (DataRow dtRow in tbl.Rows)
                {
                    string ParcelID = dtRow[tbl.Columns["ParcelID"]].ToString();
                    using (var dbContext = new MyDbContext())
                    {
                        var exception = dbContext.Exception.Where(x => x.ParcelID.Equals(ParcelID)).FirstOrDefault();
                        if(exception != null)
                        {
                            if (!string.IsNullOrWhiteSpace(dtRow[tbl.Columns["Use"]].ToString()))
                            {
                                InsertExceptionTypes(exception.ID, 1);
                            }
                            if (!string.IsNullOrWhiteSpace(dtRow[tbl.Columns["Height"]].ToString()))
                            {
                                InsertExceptionTypes(exception.ID, 2);
                            }
                            if (!string.IsNullOrWhiteSpace(dtRow[tbl.Columns["bouncing back"]].ToString()))
                            {
                                InsertExceptionTypes(exception.ID, 3);
                            }
                            if (!string.IsNullOrWhiteSpace(dtRow[tbl.Columns["Parking"]].ToString()))
                            {
                                InsertExceptionTypes(exception.ID, 4);
                            }
                            if (!string.IsNullOrWhiteSpace(dtRow[tbl.Columns["Floor ratio"]].ToString()))
                            {
                                InsertExceptionTypes(exception.ID, 5);
                            }
                            if (!string.IsNullOrWhiteSpace(dtRow[tbl.Columns["Coverage ratio"]].ToString()))
                            {
                                InsertExceptionTypes(exception.ID, 7);
                            }
                        }
                    }


                    // On all tables' columns
                    //foreach (DataColumn dc in tbl.Columns)
                    //{
                    //    var field1 = dtRow[dc].ToString();
                    //    if(dc.ColumnName.Contains("ParcelID"))
                    //    {

                    //    }
                    //}
                }
            }
        }

        public static void InsertExceptionTypes(int ExceptionID, int TypeID)
        {
            //Save Changes Types
        }
    }

    public class MyDbContext : DbContext
    {
        public DbSet<Exception> Exception { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlServer(@"Server=DESKTOP-RQRVP0B;Database=SmartSystemGIS_Dev;Trusted_Connection=True;");
        }
    }

    public class Exception
    {
        public int ID { get; set; }
        public string ParcelID { get; set; }
    }
}
