<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ExportSelectedRowsExamples.aspx.cs"
    Inherits="ExportToWordExcelPdf.ExportSelectedRowsExamples" EnableEventValidation="false" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Export selected row data from grid to Word, Excel, CSV, Pdf File Examples
    </title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table>
            <tr>
                <td colspan="3">
                    <h4>
                        Export selected row data from grid to Word, Excel, CSV, Pdf File Examples
                    </h4>
                    ** By default it'll export all data
                </td>
            </tr>
            <tr>
                <td colspan="3">
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td colspan="3">
                    <asp:Button ID="btnExportToWord" runat="server" Text="ExportToWord" OnClick="btnExportToWord_Click" />&nbsp;&nbsp;
                    <asp:Button ID="btnExportToExcel" runat="server" Text="ExportToExcel" OnClick="btnExportToExcel_Click" />&nbsp;&nbsp;
                    <asp:Button ID="btnExportToCSV" runat="server" Text="ExportToCSV" OnClick="btnExportToCSV_Click" />&nbsp;&nbsp;
                    <asp:Button ID="btnExportToText" runat="server" Text="ExportToText" OnClick="btnExportToText_Click" />&nbsp;&nbsp;
                    <asp:Button ID="btnExportToPdf" runat="server" Text="ExportToPdf" OnClick="btnExportToPdf_Click" />&nbsp;&nbsp;
                    <asp:Button ID="btnSendMail" runat="server" Text="Send Mail" OnClick="btnSendMail_Click" />
                </td>
            </tr>
            <tr>
                <td colspan="3">
                    <asp:GridView ID="grdResultDetails" runat="server" AutoGenerateColumns="false" AllowPaging="true"
                        PageSize="5" OnPageIndexChanging="grdResultDetails_PageIndexChanging" DataKeyNames="SubjectId">
                        <HeaderStyle BackColor="#9a9a9a" ForeColor="White" Font-Bold="true" Height="30" />
                        <PagerStyle HorizontalAlign="Center" />
                        <Columns>
                            <asp:TemplateField>
                                <ItemTemplate>
                                    <asp:CheckBox ID="chkSelectRow" runat="server" />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:BoundField DataField="SubjectId" HeaderText="SubjectID" ItemStyle-Width="100"
                                ItemStyle-HorizontalAlign="Center" />
                            <asp:BoundField DataField="SubjectName" HeaderText="SubjectName" ItemStyle-Width="200"
                                ItemStyle-HorizontalAlign="Center" />
                            <asp:BoundField DataField="Marks" HeaderText="Marks" ItemStyle-Width="200" ItemStyle-HorizontalAlign="Center" />
                            <asp:BoundField DataField="Grade" HeaderText="Grade" ItemStyle-Width="200" ItemStyle-HorizontalAlign="Center" />
                        </Columns>
                    </asp:GridView>
                </td>
            </tr>
        </table>
        <asp:Label ID="lblMsg" runat="server" />
    </div>
    </form>
</body>
</html>
