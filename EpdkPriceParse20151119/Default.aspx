<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="_Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style>
        body {
            font-family: arial;
            font-size: 11px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <div style="padding: 10px 0px; border-bottom: solid 1px blue; margin-bottom: 20px; font-size: 10pt; font-weight: bold;">EPDK Sayfasından Birim Fiyatları Sorgulama(Opet-Sunpet)</div>
        <table>
            <tr>
                <td>İl:</td>
                <td><asp:DropDownList runat="server" ID="ddlCity"/></td>
                <td>&nbsp;</td>
                <td>
                    <asp:Button runat="server" ID="btnSorgula" Text="Seçili İli Sorgula" OnClick="btnSorgula_OnClick" />
                </td>
                <td>&nbsp;</td>
                <td>
                    <asp:Button runat="server" ID="btnTumIllerdeSorgula" Text="Tüm İllerde Sorgula" OnClick="btnTumIllerdeSorgula_OnClick" />
                </td>
                <td>&nbsp;</td>
                <td>
                    <asp:Button runat="server" ID="btnExcel" Text="Verileri Excele Aktar" OnClick="btnExcel_OnClick" />
                </td>
            </tr>
        </table>
        <div style="margin-top: 10px; padding: 5px 5px; border: solid 1px gray; background-color: lightgrey;"><span style="font-weight: bold;">Not:</span> <span> Her il için 3 saniyede bir sorgulama yapılır. </span> </div>
        <div style="height: 20px;"></div>
        <div style="padding-top: 10px; font-size: 10pt; font-weight: bold; border-bottom: dotted 1px black; margin-bottom: 5px;">Elde Edilen Html Kodları</div>
        <div runat="server" id="divContent1" style="height: 200px; overflow: scroll;"></div>
        <div style="height: 40px;"></div>
        <div style="padding-top: 10px; font-size: 10pt; font-weight: bold; border-bottom: dotted 1px black; margin-bottom: 5px;">Toplo İl Listesi</div>
        <div>
            Kayıt Adeti:
            <asp:Label ID="lblRecordCount" runat="server" Font-Bold="True" Text="0"></asp:Label>
        </div>
        <div style="height: 20px;"></div>
        <asp:GridView runat="server" ID="gv1"></asp:GridView>
        <div style="height: 40px;"></div>
    </div>
    </form>
</body>
</html>
