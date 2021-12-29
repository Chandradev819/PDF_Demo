﻿<%@ Page Title="" Language="C#" MasterPageFile="~/Website.Master" AutoEventWireup="true"  CodeBehind="CCCPdf.aspx.cs" Inherits="PDF_Demo.View.CCCPdf" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div class="legend-width">
        <fieldset>
            <legend><b>CCC Pdf File</b></legend>
            <table>
                <tr>
                    <td>Select the Pdf File</td>
                    <td>
                        <asp:FileUpload ID="FileUpload1" runat="server" />
                    </td>
                </tr>
                <tr>
                    <td>Excel File Path:</td>
                    <td>
                        <asp:TextBox ID="txtFilePath" runat="server"></asp:TextBox>
                    </td>
                    <td>
                      <asp:RequiredFieldValidator ID="reqFilePath" ControlToValidate="txtFilePath" 
                          ValidationGroup="CCCFrame" runat="server" ErrorMessage="Please give the File Path to save file">
                      </asp:RequiredFieldValidator>  
                    </td>
                </tr>

                <tr>
                    <td>Excel File Name:</td>
                    <td>
                        <asp:TextBox ID="txtFileName" runat="server"></asp:TextBox></td>
                    <td>
                         <asp:RequiredFieldValidator ID="reqFileName" ControlToValidate="txtFilePath" 
                          ValidationGroup="CCCFrame" runat="server" ErrorMessage="Give the Excel FileName">
                      </asp:RequiredFieldValidator>  
                    </td>
                </tr>
                <tr>
                    <td> <asp:Button ID="Button1" runat="server" ValidationGroup="CCCFrame" OnClick="Button1_Click" Text="Upload" /></td>
                </tr>
            </table>
        </fieldset>
    </div>
</asp:Content>