using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using StudentApplication.Data;
using StudentApplication.Models.Excel;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Reflection;
using System.Text;

namespace StudentApplication.Controllers.Excel
{
    public class ExcelController : Controller
    {
        private readonly IConfiguration configuration;
        private readonly ApplicationContext context;

        public ExcelController(IConfiguration configuration, ApplicationContext context)
        {
            this.configuration = configuration;
            this.context = context;
        }
        public IActionResult Index()
        {
            var data = context.Students.Take(50).ToList();
            return View(data);
        }

        public class DataViewModel
        {
            public string Data { get; set; }
        }
        public IActionResult DisplayDuplicate()
        {
            string data = TempData["Data"] as string;
            DataViewModel model = new DataViewModel();
            model.Data = data;

            return View(model);
        }




        //http GET
        public IActionResult ImportExcelFile()
        {


            return View();
        }
        /* public IActionResult DisplayDuplicate()
         {
             var data1 = context.Dupchecks.Take(50).ToList();

             return View(data1);

         }*/
        [HttpPost]
        public IActionResult ImportExcelFile(IFormFile formFile)
        {

            try
            {
                var mainPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "UploadExcelFile");
                // if (formFile.FileName == "" && formFile.Length > 0)
                //{
                if (!Directory.Exists(mainPath))
                {
                    Directory.CreateDirectory(mainPath);
                }
                var filePath = Path.Combine(mainPath, formFile.FileName);
                if (filePath == null)
                {
                    TempData["message"] = "Choose a File";
                }




                using (FileStream stream = new FileStream(filePath, FileMode.Create))
                {
                    formFile.CopyTo(stream);


                }
                var fileName = formFile.FileName;
                string extension = Path.GetExtension(fileName);
                string conString = string.Empty;

                switch (extension)
                {
                    case ".xls":
                        conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties ='Excel 8.0;HDR = YES'";
                        break;
                    case ".xlsx":
                        conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties = 'Excel 8.0; HDR = YES'";
                        break;

                }
                DataTable dt = new DataTable();
                conString = String.Format(conString, filePath);
                using (OleDbConnection conExcel = new OleDbConnection(conString))
                {
                    using (OleDbCommand cmdExcel = new OleDbCommand())
                    {
                        using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                        {
                            cmdExcel.Connection = conExcel;
                            conExcel.Open();
                            DataTable dtExcelSchema = conExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                            string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                            cmdExcel.CommandText = "SELECT * FROM [" + sheetName + "]";
                            odaExcel.SelectCommand = cmdExcel;

                            odaExcel.Fill(dt);
                            conExcel.Close();


                            // Check for duplicates in the data table
                            var duplicates = dt.AsEnumerable()
                            .GroupBy(r => new { col1 = r["ExternalStudentID"], col2 = r["FirstName"], col3 = r["LastName"] })
                            .Where(g => g.Count() > 1)
                            .Select(g => g.First());

                            if (duplicates.Any())
                            {
                                foreach (var duplicate in duplicates)
                                {
                                    dt.Rows.Remove(duplicate);
                                }
                                Console.WriteLine("Duplicate rows removed successfully.");
                            }
                            else
                            {
                                conString = configuration.GetConnectionString("DefaultConnection");
                                using (SqlConnection con = new SqlConnection(conString))
                                {
                                    using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                                    {
                                        sqlBulkCopy.DestinationTableName = "Students";
                                        sqlBulkCopy.ColumnMappings.Add("ExternalStudentID", "ExternalStudentID");
                                        sqlBulkCopy.ColumnMappings.Add("FirstName", "FirstName");
                                        sqlBulkCopy.ColumnMappings.Add("LastName", "LastName");
                                        sqlBulkCopy.ColumnMappings.Add("DOB", "DOB");
                                        sqlBulkCopy.ColumnMappings.Add("SSN", "SSN");
                                        sqlBulkCopy.ColumnMappings.Add("Adddress", "Adddress");
                                        sqlBulkCopy.ColumnMappings.Add("City", "City");
                                        sqlBulkCopy.ColumnMappings.Add("State", "State");
                                        sqlBulkCopy.ColumnMappings.Add("Email", "Email");
                                        sqlBulkCopy.ColumnMappings.Add("MaritalStatus", "MaritalStatus");
                                        con.Open();
                                        SqlCommand cmd = new SqlCommand();
                                        cmd.Connection = con;
                                        cmd.CommandText = "SELECT COUNT(*) FROM Students WHERE ExternalStudentID = @ExternalStudentID";
                                        cmd.Parameters.Add("@ExternalStudentID", SqlDbType.NVarChar);
                                        DataTable dtNew = dt.Clone();
                                        DataTable dtDuplicates = dt.Clone();
                                        foreach (DataRow row in dt.Rows)
                                        {
                                            cmd.Parameters["@ExternalStudentID"].Value = row["ExternalStudentID"].ToString();
                                            int count = (int)cmd.ExecuteScalar();
                                            if (count == 0)
                                            {
                                                dtNew.ImportRow(row);
                                                TempData["message"] = "File Imported Successfully, Data Saved into DB";
                                            }
                                            else
                                            {
                                                dtDuplicates.ImportRow(row);
                                                //Duplicate Check
                                                Console.WriteLine("Duplicate Records:");
                                                StringBuilder sb = new StringBuilder();
                                                foreach (DataRow dp in dtDuplicates.Rows)
                                                {
                                                    sb.AppendLine("ExternalStudentID: " + dp["ExternalStudentID"] +
                                                                  " FirstName: " + dp["FirstName"] +
                                                                  " LastName: " + dp["LastName"] +
                                                                  " DOB: " + dp["DOB"] +
                                                                  " SSN: " + dp["SSN"] +
                                                                  " Adddress: " + dp["Adddress"] +
                                                                  " City: " + dp["City"] +
                                                                  " State: " + dp["State"] +
                                                                  " Email: " + dp["Email"] +
                                                                  " MaritalStatus: " + dp["MaritalStatus"]);
                                                }
                                                TempData["Data"] = sb.ToString();
                                                return RedirectToAction("DisplayDuplicate");


                                            }


                                        }

                                        sqlBulkCopy.WriteToServer(dtNew);
                                        con.Close();


                                    }
                                }

                            }

                        }
                    }
                }



                return RedirectToAction("Index");


            }
            catch (Exception ex)
            {
                string message = ex.Message;
            }
            return View();

        }


        public IActionResult ExportExcel()
        {
            try
            {
                var data = context.Students.ToList();
                if (data != null & data.Count > 0)
                {
                    using (XLWorkbook wb = new XLWorkbook())
                    {
                        wb.Worksheets.Add(ToConvertDataTable(data.ToList()));
                        using (MemoryStream ms = new MemoryStream())
                        {
                            wb.SaveAs(ms);
                            string fileName = $"Students_{DateTime.Now.ToString("dd/MM/yyyy")}.xlsx";
                            return File(ms.ToArray(), "application/vnd.openxmlformats-" +
                                "officedocument.spreadsheetml.sheet", fileName);
                        }
                    }

                }
                TempData["Error"] = "Data Not Found";
            }
            catch (Exception ex)
            {

            }
            return RedirectToAction("Index");
        }
        //return the datatable
        public DataTable ToConvertDataTable<T>(List<T> items)
        {
            DataTable dt = new DataTable(typeof(T).Name);
            PropertyInfo[] propInfo = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            //This Loop for getting Header/Column From DataTable
            foreach (PropertyInfo prop in propInfo)
            {
                dt.Columns.Add(prop.Name);
            }
            // for getting Rows or Data Present in the Table
            foreach (T item in items)
            {
                var value = new object[propInfo.Length];
                for (int i = 0; i < propInfo.Length; i++)
                {
                    value[i] = propInfo[i].GetValue(item, null);
                }
                dt.Rows.Add(value);

            }
            return dt;

        }
    }


}
