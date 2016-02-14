<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ExportExamples.aspx.cs"
    Inherits="ExportToWordExcelPdf.ExportExamples" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Export Data to Text File </title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table>
            <tr>
                <td colspan="3">
                    <h4>
                        Export Data to Text File 
                    </h4>
                </td>
            </tr>
            <tr>
                <td colspan="3">                  
                    <asp:Button ID="btnExportToText" runat="server" Text="ExportToText" OnClick="btnExportToText_Click" />&nbsp;&nbsp;                  
                </td>
            </tr>
        </table>
    </div>
    </form>
</body>
</html>
