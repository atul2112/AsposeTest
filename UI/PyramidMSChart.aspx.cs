using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.DataVisualization.Charting;
using System.Web.UI.WebControls;

namespace AsposeTest.UI
{
    public partial class PyramidMSChart : System.Web.UI.Page
    {

        protected void Page_Load(object sender, EventArgs e)
        {
            xyChart.Series["Series1"].ChartArea = "ChartArea1";
            xyChart.Series["Series2"].ChartArea = "ChartArea1";
            xyChart.Series["Series3"].ChartArea = "ChartArea1";
            xyChart.Series["Series1"].ChartType = SeriesChartType.StackedBar100;
            xyChart.Series["Series2"].ChartType = SeriesChartType.StackedBar100;
            xyChart.Series["Series3"].ChartType = SeriesChartType.StackedBar100;


            xyChart.Series["Series1"].Points.AddY(30);
            xyChart.Series["Series2"].Points.AddY(40);
            xyChart.Series["Series3"].Points.AddY(30);

            xyChart.Series["Series1"].Points.AddY(40);
            xyChart.Series["Series2"].Points.AddY(20);
            xyChart.Series["Series3"].Points.AddY(40);

            xyChart.Series["Series1"].Points.AddY(45);
            xyChart.Series["Series2"].Points.AddY(10);
            xyChart.Series["Series3"].Points.AddY(45);


            xyChart.Series["Series1"]["PointWidth"] = "1.0";
            xyChart.Series["Series2"]["PointWidth"] = "1.0";
            xyChart.Series["Series3"]["PointWidth"] = "1.0"; 

            xyChart.ChartAreas["ChartArea1"].AxisY.MajorGrid.Enabled = false;
            xyChart.ChartAreas["ChartArea1"].AxisX.MajorGrid.Enabled = false;

        }
           

    }
}