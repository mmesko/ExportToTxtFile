<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ExportExamples.aspx.cs"
    Inherits="ExportToWordExcelPdf.ExportExamples" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Export Grid Data to Word, Excel, CSV, Pdf, Text File Examples </title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table>
            <tr>
                <td colspan="3">
                    <h4>
                        Export Grid Data to Word, Excel, CSV, Pdf File Examples
                    </h4>
                </td>
            </tr>
            <tr>
                <td colspan="3">
                    <asp:Button ID="btnExportToWord" runat="server" Text="ExportToWord" OnClick="btnExportToWord_Click" />&nbsp;&nbsp;
                    <asp:Button ID="btnExportToExcel" runat="server" Text="ExportToExcel" OnClick="btnExportToExcel_Click" />&nbsp;&nbsp;
                    <asp:Button ID="btnExportToCSV" runat="server" Text="ExportToCSV" OnClick="btnExportToCSV_Click" />&nbsp;&nbsp;
                    <asp:Button ID="btnExportToText" runat="server" Text="ExportToText" OnClick="btnExportToText_Click" />&nbsp;&nbsp;
                    <asp:Button ID="btnExportToPdf" runat="server" Text="ExportToPdf" OnClick="btnExportToPdf_Click" />
                </td>
            </tr>
            <tr>
                <td colspan="3">
                    <asp:GridView ID="grdResultDetails" runat="server" AutoGenerateColumns="false" AllowPaging="true"
                        PageSize="5" OnPageIndexChanging="grdResultDetails_PageIndexChanging">
                        <HeaderStyle BackColor="#9a9a9a" ForeColor="White" Font-Bold="true" Height="30" />
                        <PagerStyle HorizontalAlign="Center" />
                        <AlternatingRowStyle BackColor="#f5f5f5" />
                        <Columns>
                            <asp:BoundField DataField="SubjectName" HeaderText="SubjectName" ItemStyle-Width="200"
                                ItemStyle-HorizontalAlign="Center" />
                            <asp:BoundField DataField="Marks" HeaderText="Marks" ItemStyle-Width="200" ItemStyle-HorizontalAlign="Center" />
                            <asp:BoundField DataField="Grade" HeaderText="Grade" ItemStyle-Width="200" ItemStyle-HorizontalAlign="Center" />
                        </Columns>
                    </asp:GridView>
                </td>
            </tr>
        </table>
    </div>
    </form>
</body>
</html>
