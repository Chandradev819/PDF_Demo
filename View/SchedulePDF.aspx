<%@ Page Title="" Language="C#" MasterPageFile="~/Website.Master" AutoEventWireup="true" CodeBehind="SchedulePDF.aspx.cs" Inherits="PDF_Demo.View.SchedulePDF" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
     <div class="legend-width">
        <fieldset>
            <legend><b>Schedule Pdf File</b></legend>
            <table>
                <tr>
                    <td>Select the Pdf File</td>
                    <td>
                        <asp:FileUpload ID="FileUpload1" runat="server" />
                    </td>
                </tr>
                <tr>
                    <td> </td>
                    <td><asp:Button ID="Button1" runat="server" ValidationGroup="CCCFrame" OnClick="Button1_Click" Text="Upload" /></td>
                </tr>
            </table>
        </fieldset>
    </div>
</asp:Content>
