<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Schedule.aspx.cs" Inherits="PDF_Demo.View.Schedule" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div style="margin-left:300px;"> 
            <asp:FileUpload ID="FileUpload1" runat="server" /> <br /> <br />
            <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="Click Here upload pdf file" />
        </div>
    </form>
</body>
</html>
