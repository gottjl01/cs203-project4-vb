<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Default.aspx.vb" Inherits="PFW.CSIST203.Project4._Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <table align="center" border="1">
            <tr>
                <td align="right">
                    <asp:Button ID="btnPrevious" runat="server" Text="Button" CausesValidation="False" />
                    </td>
                <td align="center"><asp:Label ID="lblID" runat="server" Text="0"></asp:Label>
                </td>
                <td align="left"><asp:Button ID="btnNext" runat="server" Text="Button" CausesValidation="False" />
                </td>
            </tr>
            <tr>
                <td align="left">First Name:</td>
                <td colspan="2"><asp:TextBox ID="txtFirstName" runat="server"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td align="left">Last Name:</td>
                <td colspan="2"><asp:TextBox ID="txtLastName" runat="server"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td align="left">E-Mail Address:</td>
                <td colspan="2"><asp:TextBox ID="txtEmailAddress" runat="server"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td align="left">Business Phone:</td>
                <td colspan="2"><asp:TextBox ID="txtBusinessPhone" runat="server"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td align="left">Company:</td>
                <td colspan="2"><asp:TextBox ID="txtCompany" runat="server"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td align="left">Job Title:</td>
                <td colspan="2"><asp:TextBox ID="txtJobTitle" runat="server"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Button ID="btnSave" runat="server" Text="Button" />
                    </td>
                <td align="center">
                    <asp:Button ID="btnReset" runat="server" Text="Button" CausesValidation="False" />
                    </td>
                <td align="left"><asp:Button ID="btnNewEntry" runat="server" Text="Button" CausesValidation="False" />
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
