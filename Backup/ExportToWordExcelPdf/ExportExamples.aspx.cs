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
        string con = @"Data Source=yourdatasource;Integrated Security=true;Initial Catalog=yourdbname";
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                LoadGridData();
            }
        }

        //Bind grid data
        private void LoadGridData()
        {
            using (SqlConnection sqlConn = new SqlConnection(con))
            {
                using (SqlCommand sqlCmd = new SqlCommand())
                {
                    sqlCmd.CommandText = "SELECT * FROM SubjectDetails";
                    sqlCmd.Connection = sqlConn;
                    sqlConn.Open();
                    SqlDataAdapter da = new SqlDataAdapter(sqlCmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    grdResultDetails.DataSource = dt;
                    grdResultDetails.DataBind();
                }
            }
        }

        //Call on gridview page index change
        protected void grdResultDetails_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            grdResultDetails.PageIndex = e.NewPageIndex;
            LoadGridData();
        }

        //Method for Export to Word
        protected void btnExportToWord_Click(object sender, EventArgs e)
        {
            string fileName = "ExportToWord_" + DateTime.Now.ToShortDateString() + ".doc",
                   contentType = "application/ms-word";

            //call 1st export method with fileName and contentType
            ExportFile(fileName, contentType);
        }

        //Method for Export to Excel
        protected void btnExportToExcel_Click(object sender, EventArgs e)
        {
            string fileName = "ExportToExcel_" + DateTime.Now.ToShortDateString() + ".xls",
                   contentType = "application/vnd.ms-excel";

            //call 1st export method with fileName and contentType
            ExportFile(fileName, contentType);
        }

        /* Method for Export to CSV
         * Note: CSV file is a text representation so we can't style .csv document*/
        protected void btnExportToCSV_Click(object sender, EventArgs e)
        {
            string fileName = "ExportToCSV_" + DateTime.Now.ToShortDateString() + ".csv",
                   contentType = "application/text";

            //call 2nd export method with fileName and contentType
            ExportTextBasedFile(fileName, contentType);
        }

        /* Method for Export to Text
         * Note: TEXT file is a text representation so we can't style .txt document*/
        protected void btnExportToText_Click(object sender, EventArgs e)
        {
            string fileName = "ExportToText_" + DateTime.Now.ToShortDateString() + ".txt",
                   contentType = "application/text";

            //call 2nd export method with fileName and contentType
            ExportTextBasedFile(fileName, contentType);
        }

        //Method for Export to PDF
        protected void btnExportToPdf_Click(object sender, EventArgs e)
        {
            //disable paging to export all data and make sure to bind griddata before begin
            grdResultDetails.AllowPaging = false;
            LoadGridData();

            string fileName = "ExportToPdf_" + DateTime.Now.ToShortDateString();
            Response.ContentType = "application/pdf";
            Response.AddHeader("content-disposition", string.Format("attachment; filename={0}", fileName + ".pdf"));
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            StringWriter objSW = new StringWriter();
            HtmlTextWriter objTW = new HtmlTextWriter(objSW);
            grdResultDetails.RenderControl(objTW);
            StringReader objSR = new StringReader(objSW.ToString());
            Document objPDF = new Document(PageSize.A4, 100f, 100f, 100f, 100f);
            HTMLWorker objHW = new HTMLWorker(objPDF);
            PdfWriter.GetInstance(objPDF, Response.OutputStream);
            objPDF.Open();
            objHW.Parse(objSR);
            objPDF.Close();
            Response.Write(objPDF);
            Response.End();
        }

        //1st Method: To Export to Word, Excel file
        private void ExportFile(string fileName, string contentType)
        {
            //disable paging to export all data and make sure to bind griddata before begin
            grdResultDetails.AllowPaging = false;
            LoadGridData();

            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", string.Format("attachment; filename={0}", fileName));
            Response.ContentType = contentType;
            StringWriter objSW = new StringWriter();
            HtmlTextWriter objTW = new HtmlTextWriter(objSW);
            grdResultDetails.RenderControl(objTW);
            Response.Write(objSW);
            Response.End();
        }

        //2nd Method: To Export to CSV, Text file
        private void ExportTextBasedFile(string fileName, string contentType)
        {
            //disable paging to export all data and make sure to bind griddata before begin
            grdResultDetails.AllowPaging = false;
            LoadGridData();

            Response.ClearContent();
            Response.AddHeader("content-disposition", string.Format("attachment; filename={0}", fileName));
            Response.ContentType = contentType;
            StringBuilder objSB = new StringBuilder();
            for (int i = 0; i < grdResultDetails.Columns.Count; i++)
            {
                objSB.Append(grdResultDetails.Columns[i].HeaderText + ',');
            }
            objSB.Append("\n");
            for (int j = 0; j < grdResultDetails.Rows.Count; j++)
            {
                for (int k = 0; k < grdResultDetails.Columns.Count; k++)
                {
                    objSB.Append(grdResultDetails.Rows[j].Cells[k].Text + ',');
                }
                objSB.Append("\n");
            }
            Response.Write(objSB.ToString());
            Response.End();
        }

        /* Added to resolve following error:
           Control 'grdResultDetails' of type 'GridView' must be placed inside a form tag with runat=server.
           http://www.aspneto.com/control-gridview1-of-type-gridview-must-be-placed-inside-a-form-tag-with-runatserver.html
        */
        public override void VerifyRenderingInServerForm(Control control)
        {
            //Required to verify that the control is rendered properly on page
        }
    }
}