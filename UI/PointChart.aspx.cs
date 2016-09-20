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
    public partial class PointChart : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            GetChart();
        }

        private void GetChart()
        {
            PointChartNew.Series.Clear();
            PointChartNew.ChartAreas.Clear();

            PointChartNew.Series.Add("Aguila");
            PointChartNew.Series.Add("AguilaLight");
            PointChartNew.Series.Add("clubColombia");
            PointChartNew.Series.Add("corona");
            

            PointChartNew.ChartAreas.Add("ChartArea1");

            //PointChartNew.Series["Aguila"].Color = Color.LightBlue;
            //PointChartNew.Series["AguilaLight"].Color = Color.LightGreen;
            //PointChartNew.Series["clubColombia"].Color = Color.LightPink;
            //PointChartNew.Series["corona"].Color = Color.LightGray;



            PointChartNew.Series["Aguila"].ChartArea = "ChartArea1";
            PointChartNew.Series["AguilaLight"].ChartArea = "ChartArea1";
            PointChartNew.Series["clubColombia"].ChartArea = "ChartArea1";
            PointChartNew.Series["corona"].ChartArea = "ChartArea1";
           

            PointChartNew.Series["Aguila"].ChartType = SeriesChartType.Point;
            PointChartNew.Series["AguilaLight"].ChartType = SeriesChartType.Point;
            PointChartNew.Series["clubColombia"].ChartType = SeriesChartType.Point;
            PointChartNew.Series["corona"].ChartType = SeriesChartType.Point;



            PointChartNew.Series["Aguila"].Points.AddY(12);
            PointChartNew.Series["AguilaLight"].Points.AddY(40);
            PointChartNew.Series["clubColombia"].Points.AddY(30);
            PointChartNew.Series["corona"].Points.AddY(20);

            PointChartNew.Series["Aguila"].Points.AddY(21);
            PointChartNew.Series["AguilaLight"].Points.AddY(22);
            PointChartNew.Series["clubColombia"].Points.AddY(23);
            PointChartNew.Series["corona"].Points.AddY(24);

            PointChartNew.Series["Aguila"].Points.AddY(13);
            PointChartNew.Series["AguilaLight"].Points.AddY(42);
            PointChartNew.Series["clubColombia"].Points.AddY(32);
            PointChartNew.Series["corona"].Points.AddY(30);

            PointChartNew.Series["Aguila"].Points.AddY(28);
            PointChartNew.Series["AguilaLight"].Points.AddY(12);
            PointChartNew.Series["clubColombia"].Points.AddY(48);
            PointChartNew.Series["corona"].Points.AddY(26);

            PointChartNew.Series["Aguila"].Points.AddY(15);
            PointChartNew.Series["AguilaLight"].Points.AddY(15);
            PointChartNew.Series["clubColombia"].Points.AddY(33);
            PointChartNew.Series["corona"].Points.AddY(30);

            PointChartNew.Series["Aguila"].AxisLabel = "Name1";
            PointChartNew.Series["AguilaLight"].AxisLabel = "Name2";
            PointChartNew.Series["clubColombia"].AxisLabel = "Name3";
            PointChartNew.Series["corona"].AxisLabel = "Name4";



            PointChartNew.Series["Aguila"].MarkerStyle = MarkerStyle.Circle;
            PointChartNew.Series["Aguila"].MarkerSize = 30;
          //  PointChartNew.Series["Aguila"].MarkerColor = Color.LightBlue;

            PointChartNew.Series["AguilaLight"].MarkerStyle = MarkerStyle.Circle;
            PointChartNew.Series["AguilaLight"].MarkerSize = 30;
        //    PointChartNew.Series["AguilaLight"].MarkerColor = Color.LightGreen;

            PointChartNew.Series["clubColombia"].MarkerStyle = MarkerStyle.Circle;
            PointChartNew.Series["clubColombia"].MarkerSize = 30;
         //   PointChartNew.Series["clubColombia"].MarkerColor = Color.LightPink;

            PointChartNew.Series["corona"].MarkerStyle = MarkerStyle.Circle;
            PointChartNew.Series["corona"].MarkerSize = 30;
        //    PointChartNew.Series["corona"].MarkerColor = Color.LightGray;

            //PointChartNew.ChartAreas[0].Area3DStyle.Enable3D = true;
            //PointChartNew.ChartAreas[0].Area3DStyle.Rotation = 40;

            PointChartNew.ChartAreas[0].AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dot;
            PointChartNew.ChartAreas[0].AxisY.MajorGrid.Enabled = false;

            PointChartNew.ChartAreas[0].AxisY.Maximum = 100;
            PointChartNew.ChartAreas[0].AxisY.Minimum = 0;
            PointChartNew.ChartAreas[0].AxisY.Interval = 25;


           
            SetChartTransparency(PointChartNew, 200);
        }

        private void SetChartTransparency(Chart chart, byte alpha)
        {
            // Apply palette colors so that they are populated into chart before being manipulated
            chart.ApplyPaletteColors();

            // Iterate through data points and set alpha values for each
            foreach (Series series in chart.Series)
                foreach (DataPoint point in series.Points)
                    point.Color = Color.FromArgb(alpha, point.Color);
        }



    }
}