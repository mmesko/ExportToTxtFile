using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.IO;
using System.Text;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.html.simpleparser;
using iTextSharp.text.pdf;
using System.Data.SqlClient;

namespace ExportToWordExcelPdf
{
    public partial class ExportExamples : System.Web.UI.Page
    {
        string con = @"Data Source=maja;Initial Catalog=Company;Integrated Security=True;Connect Timeout=15;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False";
        DataTable dt = new DataTable();

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                LoadData();
            }
        }

        private void LoadData()
        {
            using (SqlConnection sqlConn = new SqlConnection(con))
            {
                using (SqlCommand sqlCmd = new SqlCommand())
                {
                    sqlCmd.CommandText = "SELECT * FROM Department";
                    sqlCmd.Connection = sqlConn;
                    sqlConn.Open();
                    SqlDataAdapter da = new SqlDataAdapter(sqlCmd);
                    da.Fill(dt);
                }
            }
        }

    
        protected void btnExportToText_Click(object sender, EventArgs e)
        {
            string fileName = "ExportToText_" + DateTime.Now.ToShortDateString() + ".txt",
                   contentType = "application/text";

            //call export method with fileName and contentType
            ExportTextBasedFile(fileName, contentType);
        }



        //Method: To Export to CSV, Text file
        private void ExportTextBasedFile(string fileName, string contentType)
        {
           
            LoadData();
            Response.ClearContent();
            Response.AddHeader("content-disposition", string.Format("attachment; filename={0}", fileName));
            Response.ContentType = contentType;

            StringBuilder objSB = new StringBuilder();

            //calculate max length of column 

            int[] maxLengths = new int[dt.Columns.Count];

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                maxLengths[i] = dt.Columns[i].ColumnName.Length;

                foreach (DataRow row in dt.Rows)
                {
                    if (!row.IsNull(i))
                    {
                        int length = row[i].ToString().Length;

                        if (length > maxLengths[i])
                        {
                            maxLengths[i] = length;
                        }
                    }
                }
            }


            objSB.AppendLine("x------------------------------------------------x");
            //header 
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                objSB.Append(dt.Columns[i].ColumnName.PadRight(maxLengths[i] + 2)); 
            }
            objSB.AppendLine("x------------------------------------------------x");


            //rows filled with data
            foreach (DataRow row in dt.Rows)
            {

                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    if (!row.IsNull(i))
                    {
                        objSB.Append(row[i].ToString().PadRight(maxLengths[i] + 2));
                    }
                    else
                    {
                        objSB.Append(new string(' ', maxLengths[i] + 2));
                    }
                   
                    
                }
                objSB.AppendLine("x------------------------------------------------x");
                objSB.AppendLine(Environment.NewLine);
            }

            
            Response.Write(objSB.ToString());
            Response.End();
        }

    }
}