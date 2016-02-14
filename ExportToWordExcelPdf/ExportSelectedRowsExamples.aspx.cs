using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.html.simpleparser;
using iTextSharp.text.pdf;
using System.Collections;
using System.Data.SqlClient;
using System.Net.Mail;

namespace ExportToWordExcelPdf
{
    public partial class ExportSelectedRowsExamples : System.Web.UI.Page
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

        //Call on gridview page rowIndex change
        protected void grdResultDetails_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            //Save checked rows before page change
            SaveCheckedStates();
            grdResultDetails.PageIndex = e.NewPageIndex;
            LoadGridData();
            //Populate cheked items with its checked status
            PopulateCheckedStates();
        }

        //Method for Export to Word
        protected void btnExportToWord_Click(object sender, EventArgs e)
        {
            string fileName = "ExportToWord_" + DateTime.Now.ToShortDateString() + ".doc",
                   contentType = "application/ms-word";

            //call 1st export method with filename and contenttype
            ExportFile(fileName, contentType);
        }

        //Method for Export to Excel
        protected void btnExportToExcel_Click(object sender, EventArgs e)
        {
            string fileName = "ExportToExcel_" + DateTime.Now.ToShortDateString() + ".xls",
                   contentType = "application/vnd.ms-excel";

            //call 1st export method with filename and contenttype
            ExportFile(fileName, contentType);
        }

        /* Method for Export to CSV
         * Note: CSV file is a text representation so we can't style .csv document*/
        protected void btnExportToCSV_Click(object sender, EventArgs e)
        {
            string fileName = "ExportToCsv_" + DateTime.Now.ToShortDateString() + ".csv",
                   contentType = "application/text";

            //call 2nd export method with filename and contenttype
            ExportTextBasedFile(fileName, contentType);
        }

        /* Method for Export to Text
         * Note: TEXT file is a text representation so we can't style .txt document*/
        protected void btnExportToText_Click(object sender, EventArgs e)
        {
            string fileName = "ExportToText_" + DateTime.Now.ToShortDateString() + ".txt",
                   contentType = "application/text";

            //call 2nd export method with filename and contenttype
            ExportTextBasedFile(fileName, contentType);
        }

        //Method for Export to PDF
        protected void btnExportToPdf_Click(object sender, EventArgs e)
        {
            SaveCheckedStates();
            //disable paging to export all data and make sure to bind griddata before begin
            grdResultDetails.AllowPaging = false;
            LoadGridData();

            string fileName = "ExportToPdf_" + DateTime.Now.ToShortDateString();
            Response.ContentType = "application/pdf";
            Response.AddHeader("content-disposition", string.Format("attachment; filename={0}", fileName + ".pdf"));
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            StringWriter objSW = new StringWriter();
            HtmlTextWriter objTW = new HtmlTextWriter(objSW);
            grdResultDetails.Columns[0].Visible = false;
            if (ViewState["SELECTED_ROWS"] != null)
            {
                ArrayList objSelectedRowsAL = (ArrayList)ViewState["SELECTED_ROWS"];
                for (int j = 0; j < grdResultDetails.Rows.Count; j++)
                {
                    GridViewRow row = grdResultDetails.Rows[j];
                    int rowIndex = Convert.ToInt32(grdResultDetails.DataKeys[row.RowIndex].Value);
                    if (!objSelectedRowsAL.Contains(rowIndex))
                    {
                        //make invisible because row is not checked
                        row.Visible = false;
                    }
                }
            }
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

        //Populate the saved checked checkbox status
        private void PopulateCheckedStates()
        {
            ArrayList objSubjectAL = (ArrayList)ViewState["SELECTED_ROWS"];
            if (objSubjectAL != null && objSubjectAL.Count > 0)
            {
                foreach (GridViewRow row in grdResultDetails.Rows)
                {
                    int rowIndex = Convert.ToInt32(grdResultDetails.DataKeys[row.RowIndex].Value);
                    if (objSubjectAL.Contains(rowIndex))
                    {
                        CheckBox chkSelectRow = (CheckBox)row.FindControl("chkSelectRow");
                        chkSelectRow.Checked = true;
                    }
                }
            }
        }

        //Save the state of row checkboxes
        private void SaveCheckedStates()
        {
            ArrayList objSubjectAL = new ArrayList();
            int rowIndex = -1;
            foreach (GridViewRow row in grdResultDetails.Rows)
            {
                rowIndex = Convert.ToInt32(grdResultDetails.DataKeys[row.RowIndex].Value);
                bool isSelected = ((CheckBox)row.FindControl("chkSelectRow")).Checked;
                if (ViewState["SELECTED_ROWS"] != null)
                {
                    objSubjectAL = (ArrayList)ViewState["SELECTED_ROWS"];
                }
                if (isSelected)
                {
                    if (!objSubjectAL.Contains(rowIndex))
                    {
                        objSubjectAL.Add(rowIndex);
                    }
                }
                else
                {
                    objSubjectAL.Remove(rowIndex);
                }
            }
            if (objSubjectAL != null && objSubjectAL.Count > 0)
            {
                ViewState["SELECTED_ROWS"] = objSubjectAL;
            }
        }

        //1st Method: To Export to Word, Excel file
        private void ExportFile(string fileName, string contentType)
        {
            SaveCheckedStates();
            //disable paging to export all pages data and make sure to bind griddata before begin
            grdResultDetails.AllowPaging = false;
            LoadGridData();
            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", string.Format("attachment; filename={0}", fileName));
            Response.ContentType = contentType;
            StringWriter objSW = new StringWriter();
            HtmlTextWriter objHW = new HtmlTextWriter(objSW);
            grdResultDetails.HeaderRow.Style.Add("background-color", "#fff");
            grdResultDetails.Columns[0].Visible = false;
            for (int i = 0; i < grdResultDetails.HeaderRow.Cells.Count; i++)
            {
                grdResultDetails.HeaderRow.Cells[i].Style.Add("background-color", "#9a9a9a");
            }
            if (ViewState["SELECTED_ROWS"] != null)
            {
                ArrayList objSelectedRowsAL = (ArrayList)ViewState["SELECTED_ROWS"];
                for (int j = 0; j < grdResultDetails.Rows.Count; j++)
                {
                    GridViewRow row = grdResultDetails.Rows[j];
                    int rowIndex = Convert.ToInt32(grdResultDetails.DataKeys[row.RowIndex].Value);
                    if (!objSelectedRowsAL.Contains(rowIndex))
                    {
                        //make invisible because row is not checked
                        row.Visible = false;
                    }
                }
            }
            grdResultDetails.RenderControl(objHW);
            Response.Write(objSW);
            Response.End();
        }

        //2nd Method: To Export to CSV, Text file
        private void ExportTextBasedFile(string fileName, string contentType)
        {
            SaveCheckedStates();
            //disable paging to export all data and make sure to bind griddata before begin
            grdResultDetails.AllowPaging = false;
            LoadGridData();

            Response.ClearContent();
            Response.AddHeader("content-disposition", string.Format("attachment; filename={0}", fileName));
            Response.ContentType = contentType;
            StringBuilder objSB = new StringBuilder();
            grdResultDetails.Columns[0].Visible = false;
            for (int i = 1; i < grdResultDetails.Columns.Count; i++)
            {
                objSB.Append(grdResultDetails.Columns[i].HeaderText + ',');
            }
            objSB.Append("\n");

            ArrayList objSelectedRowsAL = (ArrayList)ViewState["SELECTED_ROWS"];
            for (int j = 0; j < grdResultDetails.Rows.Count; j++)
            {
                bool isRowSelected = true;
                if (ViewState["SELECTED_ROWS"] != null)
                {
                    GridViewRow row = grdResultDetails.Rows[j];
                    int rowIndex = Convert.ToInt32(grdResultDetails.DataKeys[row.RowIndex].Value);
                    isRowSelected = objSelectedRowsAL.Contains(rowIndex);
                }
                if (isRowSelected)
                {
                    //if row is selected then add row to csv file, else ignore row
                    for (int k = 1; k < grdResultDetails.Columns.Count; k++)
                    {
                        objSB.Append(grdResultDetails.Rows[j].Cells[k].Text + ',');
                    }
                    objSB.Append("\n");
                }
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

        //Method To Convert Gridview To HTML formatted String
        private string GridViewToHtml(GridView grdResultDetails)
        {
            SaveCheckedStates();
            grdResultDetails.AllowPaging = false;
            LoadGridData();
            StringBuilder objSB = new StringBuilder();
            StringWriter objSW = new StringWriter(objSB);
            HtmlTextWriter objHW = new HtmlTextWriter(objSW);
            if (ViewState["SELECTED_ROWS"] != null)
            {
                ArrayList objSelectedRowsAL = (ArrayList)ViewState["SELECTED_ROWS"];
                for (int j = 0; j < grdResultDetails.Rows.Count; j++)
                {
                    GridViewRow row = grdResultDetails.Rows[j];
                    int rowIndex = Convert.ToInt32(grdResultDetails.DataKeys[row.RowIndex].Value);
                    if (!objSelectedRowsAL.Contains(rowIndex))
                    {
                        //make invisible because row is not checked
                        row.Visible = false;
                    }
                }
            }
            grdResultDetails.RenderControl(objHW);
            return objSB.ToString();
        }

        //Method To Send Mail
        protected void btnSendMail_Click(object sender, EventArgs e)
        {
            try
            {
                string Subject = "This is test mail with gridview data",
                Body = GridViewToHtml(grdResultDetails),
                ToEmail = "toemail@domain.com";

                string SMTPUser = "email@domain.com", SMTPPassword = "password";

                //Now instantiate a new instance of MailMessage
                MailMessage mail = new MailMessage();

                //set the sender address of the mail message
                mail.From = new System.Net.Mail.MailAddress(SMTPUser, "Webblogsforyou");

                //set the recepient addresses of the mail message
                mail.To.Add(ToEmail);

                //set the subject of the mail message
                mail.Subject = Subject;

                //set the body of the mail message
                mail.Body = Body;

                //leave as it is even if you are not sending HTML message
                mail.IsBodyHtml = true;

                //set the priority of the mail message to normal
                mail.Priority = System.Net.Mail.MailPriority.Normal;

                //instantiate a new instance of SmtpClient
                SmtpClient smtp = new SmtpClient();

                //if you are using your smtp server, then change your host like "smtp.yourdomain.com"
                smtp.Host = "smtp.gmail.com";

                //chnage your port for your host
                smtp.Port = 25; //or you can also use port# 587

                //provide smtp credentials to authenticate to your account
                smtp.Credentials = new System.Net.NetworkCredential(SMTPUser, SMTPPassword);

                //if you are using secure authentication using SSL/TLS then "true" else "false"
                smtp.EnableSsl = true;

                smtp.Send(mail);

                lblMsg.Text = "Success: Mail sent successfully!";
                lblMsg.ForeColor = System.Drawing.Color.Green;
            }
            catch (SmtpException ex)
            {
                //catched smtp exception
                lblMsg.Text = "SMTP Exception: " + ex.Message.ToString();
                lblMsg.ForeColor = System.Drawing.Color.Red;
            }
            catch (Exception ex)
            {
                lblMsg.Text = "Error: " + ex.Message.ToString();
                lblMsg.ForeColor = System.Drawing.Color.Red;
            }
        }
    }
}