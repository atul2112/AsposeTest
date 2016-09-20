using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using Aspose.Slides;
using Aspose.Slides.Charts;
using System.Drawing;
using System.IO;

namespace AsposeTest.UI
{
    public partial class LineChartFinalNew : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            GetPPT();
        }


        private void GetPPT()
        {
            try
            {
                //Get the Presentation PPTX file
                Presentation pres = GetPresentation();

                var RandomValue = new Random();
                MemoryStream objMemoryStream = new MemoryStream();

                pres.Save(objMemoryStream, Aspose.Slides.Export.SaveFormat.Pptx);

                byte[] buffer = objMemoryStream.ToArray();

                HttpContext.Current.Response.Clear();
                HttpContext.Current.Response.Buffer = true;
                HttpContext.Current.Response.AddHeader("Content-disposition", "attachment; filename=demo.pptx");

                HttpContext.Current.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                HttpContext.Current.Response.AddHeader("Content-Length", buffer.Length.ToString());
                HttpContext.Current.Response.BinaryWrite(buffer);
                HttpContext.Current.Response.Flush();
                HttpContext.Current.Response.Close();

            }
            catch (Exception ex)
            {
                LblError.Text = ex.Message.ToString();
            }
        }


        private Presentation GetPresentation()
        {
            Presentation pres = new Presentation();

            #region Data Table Values
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("Year");
            dt.Columns.Add("Pepsi");
            dt.Columns.Add("Coke");
            dt.Columns.Add("Dew");
            dt.Rows.Add(2000, 85, 80, 75);
            dt.Rows.Add(2003, 72, 90, 80);
            dt.Rows.Add(2006, 84, 87, 85);
            dt.Rows.Add(2009, 92, 88, 85);
            dt.Rows.Add(2012, 84, 80, 80);
            #endregion


            //Access First Slide
            ISlide slide1 = pres.Slides[0];


            #region Add Line Chart to Slide 1
            //Add Chart to the Slide
            IChart LineChart1 = slide1.Shapes.AddChart(ChartType.LineWithMarkers, 50, 50, 600, 450);

            #region Chart Title
            //Assign Chart title
            LineChart1.HasTitle = true;
            LineChart1.ChartTitle.AddTextFrameForOverriding("SoftDrinks Analysis");
            LineChart1.ChartTitle.TextFrameForOverriding.TextFrameFormat.CenterText = NullableBool.True;
            LineChart1.ChartTitle.TextFrameForOverriding.Paragraphs[0].Portions[0].PortionFormat.FillFormat.FillType = FillType.Solid;
            LineChart1.ChartTitle.TextFrameForOverriding.Paragraphs[0].Portions[0].PortionFormat.FillFormat.SolidFillColor.Color = Color.Blue;
            LineChart1.ChartTitle.TextFrameForOverriding.Paragraphs[0].Portions[0].PortionFormat.FontBold = NullableBool.True;
            LineChart1.ChartTitle.TextFrameForOverriding.Paragraphs[0].Portions[0].PortionFormat.FontHeight = 20;
            LineChart1.ChartTitle.X = 400;
            LineChart1.ChartTitle.Y = 200;
            #endregion

            //Getting the Chart Data Workbook
            IChartDataWorkbook workbook1 = LineChart1.ChartData.ChartDataWorkbook;

            //Setting the index of chart data sheet
            int defaultWorksheetIndex = 0;

            //Delete Default generated series and categories
            LineChart1.ChartData.Series.Clear();
            LineChart1.ChartData.Categories.Clear();

            IChartSeries Series;
            IDataLabel lbl;

            #region Add Series
            //Setting Series Name
            for (int i = 1; i < dt.Columns.Count; i++)
            {
                LineChart1.ChartData.Series.Add(workbook1.GetCell(defaultWorksheetIndex, 0, i, dt.Columns[i].ColumnName), LineChart1.Type);
            }
            #endregion

            #region Add Category
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //Setting Category Name
                LineChart1.ChartData.Categories.Add(workbook1.GetCell(defaultWorksheetIndex, i + 1, 0, dt.Rows[i][0].ToString()));
            }

            #endregion

            #region Add Series Data
            //Add DataPoints for the Series
            for (int i = 1; i < dt.Columns.Count; i++)
            {
                Series = LineChart1.ChartData.Series[i - 1];

                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    string Val = dt.Rows[j][i].ToString();
                    Series.DataPoints.AddDataPointForLineSeries(workbook1.GetCell(defaultWorksheetIndex, j + 1, i, Convert.ToDouble(dt.Rows[j][i])));

                    //Set Data Point Label Style
                    lbl = Series.DataPoints[j].Label;
                    lbl.DataLabelFormat.Position = LegendDataLabelPosition.Top;
                    lbl.DataLabelFormat.ShowValue = true;
                    lbl.DataLabelFormat.TextFormat.PortionFormat.FillFormat.FillType = FillType.Solid;
                    lbl.DataLabelFormat.TextFormat.PortionFormat.FillFormat.SolidFillColor.Color = Color.Green;
                    lbl.DataLabelFormat.TextFormat.PortionFormat.FontBold = NullableBool.True;
                    lbl.DataLabelFormat.TextFormat.PortionFormat.FontHeight = 20;
                }

                //Set DataPoint Marker Style
                Series.Marker.Symbol = MarkerStyleType.Circle;
                Series.Marker.Size = 10;
            }

            #endregion

            #region Setting chart maximum, minimum values
            //Set chart Maximum, Minimum values
            LineChart1.Axes.VerticalAxis.IsAutomaticMajorUnit = false;
            LineChart1.Axes.VerticalAxis.IsAutomaticMaxValue = false;
            LineChart1.Axes.VerticalAxis.IsAutomaticMinorUnit = false;
            LineChart1.Axes.VerticalAxis.IsAutomaticMinValue = false;

            LineChart1.Axes.VerticalAxis.MaxValue = 125;
            LineChart1.Axes.VerticalAxis.MinValue = 70;
            LineChart1.Axes.VerticalAxis.MinorUnit = 5;
            LineChart1.Axes.VerticalAxis.MajorUnit = 10;
            #endregion

            #region Seting Value and Category Text Properties

            //Setting Value Axis Text Properties         
            IChartPortionFormat YaxisVal = LineChart1.Axes.VerticalAxis.TextFormat.PortionFormat;
            YaxisVal.FontBold = NullableBool.True;
            YaxisVal.FontHeight = 20;
            YaxisVal.FontItalic = NullableBool.True;
            YaxisVal.FillFormat.FillType = FillType.Solid;
            YaxisVal.FillFormat.SolidFillColor.Color = Color.Blue;
            YaxisVal.LatinFont = new FontData("Times New Roman");


            //Setting Category Axis Text Properties
            IChartPortionFormat XAxisVal = LineChart1.Axes.HorizontalAxis.TextFormat.PortionFormat;
            XAxisVal.FontBold = NullableBool.True;
            XAxisVal.FontHeight = 16;
            XAxisVal.FontItalic = NullableBool.True;
            XAxisVal.FillFormat.FillType = FillType.Solid; ;
            XAxisVal.FillFormat.SolidFillColor.Color = Color.Blue;
            XAxisVal.LatinFont = new FontData("Arial");

            #endregion

            #region Set Chart Axis Properties
            //Set VerticalAxis Axis GridLines Style
            LineChart1.Axes.VerticalAxis.MajorGridLinesFormat.Line.FillFormat.FillType = FillType.Solid;
            LineChart1.Axes.VerticalAxis.MajorGridLinesFormat.Line.FillFormat.SolidFillColor.Color = Color.Blue;
            LineChart1.Axes.VerticalAxis.MajorGridLinesFormat.Line.DashStyle = LineDashStyle.Dot;

            //Set Vertical Axis Title
            LineChart1.Axes.VerticalAxis.HasTitle = true;
            LineChart1.Axes.VerticalAxis.Title.AddTextFrameForOverriding("Values in Percentage");
            LineChart1.Axes.VerticalAxis.Title.TextFrameForOverriding.Paragraphs[0].Portions[0].PortionFormat.FontHeight = 20;
            LineChart1.Axes.VerticalAxis.Title.TextFrameForOverriding.Paragraphs[0].Portions[0].PortionFormat.FillFormat.FillType = FillType.Solid;
            LineChart1.Axes.VerticalAxis.Title.TextFrameForOverriding.Paragraphs[0].Portions[0].PortionFormat.FillFormat.SolidFillColor.Color = Color.DarkViolet;

            //Set Horizontal axis Title
            LineChart1.Axes.HorizontalAxis.HasTitle = true;
            LineChart1.Axes.HorizontalAxis.Title.AddTextFrameForOverriding("Year");
            LineChart1.Axes.HorizontalAxis.Title.TextFrameForOverriding.Paragraphs[0].Portions[0].PortionFormat.FontHeight = 20;

            LineChart1.Axes.HorizontalAxis.Title.TextFrameForOverriding.Paragraphs[0].Portions[0].PortionFormat.FillFormat.FillType = FillType.Solid;
            LineChart1.Axes.HorizontalAxis.Title.TextFrameForOverriding.Paragraphs[0].Portions[0].PortionFormat.FillFormat.SolidFillColor.Color = Color.DarkViolet;

            #endregion

            #region Set Legend
            //Set Legend Style
            LineChart1.HasLegend = true;
            LineChart1.Legend.Position = LegendPositionType.Right;
            LineChart1.Legend.TextFormat.PortionFormat.FontBold = NullableBool.True;
            LineChart1.Legend.TextFormat.PortionFormat.FontHeight = 20;
            #endregion


            #endregion



            return pres;
        }
    }
}