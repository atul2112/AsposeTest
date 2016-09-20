<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="PyramidMSChart.aspx.cs" Inherits="AsposeTest.UI.PyramidMSChart" %>

<%@ Register Assembly="System.Web.DataVisualization, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" Namespace="System.Web.UI.DataVisualization.Charting" TagPrefix="asp" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Chart ID="xyChart" runat="server" Width="500px" Height="500px">
            <Series>
                <asp:Series Name="Series1" ChartType="StackedBar100" Color="Transparent"></asp:Series>
                <asp:Series Name="Series2" ChartType="StackedBar100" Color="Green"></asp:Series>
                <asp:Series Name="Series3" ChartType="StackedBar100" Color="Transparent"></asp:Series>
            </Series>
            <ChartAreas>
                <asp:ChartArea Name="ChartArea1"></asp:ChartArea>
            </ChartAreas>
        </asp:Chart>
    </div>
    </form>
</body>
</html>
