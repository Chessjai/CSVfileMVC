using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;


namespace CSVfileMVC.Controllers
{
    public class HomeController : Controller
    {
        // GET: Home
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult ExportCSV()
        {
            
            string constr = ConfigurationManager.ConnectionStrings["MyConnectionString"].ConnectionString;
            using (SqlConnection con = new SqlConnection(constr))
            {
                using (SqlCommand cmd = new SqlCommand("SELECT * FROM Student3"))
                {
                    using (SqlDataAdapter sda = new SqlDataAdapter())
                    {
                        cmd.Connection = con;
                        sda.SelectCommand = cmd;
                        using (DataTable dt = new DataTable())
                        {
                            sda.Fill(dt);

                            //Build the CSV file data as a Comma separated string.
                            string csv = string.Empty;

                            foreach (DataColumn column in dt.Columns)
                            {
                                //Add the Header row for CSV file.
                                csv += column.ColumnName + ',';
                            }

                            //Add new line.
                            csv += "\r\n";

                            foreach (DataRow row in dt.Rows)
                            {
                                foreach (DataColumn column in dt.Columns)
                                {
                                    //Add the Data rows.
                                    csv += row[column.ColumnName].ToString().Replace(",", ";") + ',';
                                }

                                //Add new line.
                                csv += "\r\n";
                            }

                            //Download the CSV file.
                            Response.Clear();
                            Response.Buffer = true;
                            Response.AddHeader("content-disposition", "attachment;filename=SqlExport.csv");
                            Response.Charset = "";
                            Response.ContentType = "application/text";
                            Response.Output.Write(csv);
                            Response.Flush();
                            Response.End();
                        }
                    }
                }
            }
            return View();
        }

        [HttpPost]
            public ActionResult Index(HttpPostedFileBase postedFile)
            {
            string filepath=string.Empty;
            if (postedFile!=null)
            {
                try
                {
                    string filePath = string.Empty;
                    if (postedFile != null)
                    {
                        string path = Server.MapPath("~/Uploads/");
                        if (!Directory.Exists(path))
                        {
                            Directory.CreateDirectory(path);
                        }

                        filePath = path + Path.GetFileName(postedFile.FileName);
                        string extension = Path.GetExtension(postedFile.FileName);
                        postedFile.SaveAs(filePath);

                        //Create a DataTable.
                        DataTable dt = new DataTable();
                        dt.Columns.AddRange(new DataColumn[3] {

                                new DataColumn("Name", typeof(string)),
                                new DataColumn("Age", typeof(int)),
                                new DataColumn("Address",typeof(string)) });


                        //Read the contents of CSV file.
                        string csvData = System.IO.File.ReadAllText(filePath);

                        //Execute a loop over the rows.
                        foreach (string row in csvData.Split('\n'))
                        {
                            if (!string.IsNullOrEmpty(row))
                            {
                                dt.Rows.Add();
                                int i = 0;

                                //Execute a loop over the columns.
                                foreach (string cell in row.Split(','))
                                {
                                    dt.Rows[dt.Rows.Count - 1][i] = cell;
                                    i++;
                                }
                            }
                        }


                        string conString = ConfigurationManager.ConnectionStrings["MyConnectionString"].ConnectionString;
                        using (SqlConnection con = new SqlConnection(conString))
                        {
                            using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                            {
                                //Set the database table name.
                                sqlBulkCopy.DestinationTableName = "Student3";

                                //[OPTIONAL]: Map the DataTable columns with that of the database table

                                sqlBulkCopy.ColumnMappings.Add("Name", "Name");
                                sqlBulkCopy.ColumnMappings.Add("Age", "Age");
                                sqlBulkCopy.ColumnMappings.Add("Address", "Address");

                                con.Open();
                                sqlBulkCopy.WriteToServer(dt);
                                con.Close();
                            }
                        }
                        if (dt.Rows.Count > 0)
                        {

                            ViewBag.mg = "data added sucessfully";
                        }
                        else if (dt.Rows.Count == 0)
                        {
                            ViewBag.msg = "choice file";
                        }

                    }

                }
                catch (Exception ex)
                {

                    ViewBag.msg = ex.Message;
                    return View();
                }

            }
            else
            {
                ViewBag.message = "* Please upload file";
            }





            return View();
            }
        }
}