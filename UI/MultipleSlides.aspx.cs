using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace AsposeTest.UI
{
    public partial class MultipleSlides : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
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

                //Instantiate Presentation class that represents PPTX file
                Presentation pres = new Presentation();

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
                    LineChart1.ChartData.Series.Add(workbook1.GetCell(defaultWorksheetIndex, 0, i, "Series" + i.ToString()), LineChart1.Type);
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

                LineChart1.Axes.VerticalAxis.MaxValue = 100;
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


                Presentation Pres2 = new Presentation();

                //Access Second Slide
                ISlide slide2 = Pres2.Slides[0];

                #region Add Column Chart to Slide 2

                //Add Chart to the Slide
                IChart ColumnChart = slide2.Shapes.AddChart(ChartType.ClusteredColumn, 50, 50, 600, 450);

                #region Chart Title
                //Assign Chart title
                ColumnChart.HasTitle = true;
                ColumnChart.ChartTitle.AddTextFrameForOverriding("SoftDrinks Analysis");
                ColumnChart.ChartTitle.TextFrameForOverriding.TextFrameFormat.CenterText = NullableBool.True;
                ColumnChart.ChartTitle.TextFrameForOverriding.Paragraphs[0].Portions[0].PortionFormat.FillFormat.FillType = FillType.Solid;
                ColumnChart.ChartTitle.TextFrameForOverriding.Paragraphs[0].Portions[0].PortionFormat.FillFormat.SolidFillColor.Color = Color.Blue;
                ColumnChart.ChartTitle.TextFrameForOverriding.Paragraphs[0].Portions[0].PortionFormat.FontBold = NullableBool.True;
                ColumnChart.ChartTitle.TextFrameForOverriding.Paragraphs[0].Portions[0].PortionFormat.FontHeight = 20;

                #endregion

                //Getting the Chart Data Workbook
                IChartDataWorkbook workbook2 = ColumnChart.ChartData.ChartDataWorkbook;

                //Setting the index of chart data sheet
                int defaultWorksheetIndex2 = 0;

                //Delete Default generated series and categories
                ColumnChart.ChartData.Series.Clear();
                ColumnChart.ChartData.Categories.Clear();

                IChartSeries Series2;
                IDataLabel lbl2;

                #region Add Series
                //Setting Series Name
                for (int i = 1; i < dt.Columns.Count; i++)
                {
                    ColumnChart.ChartData.Series.Add(workbook2.GetCell(defaultWorksheetIndex2, 0, i, "Series" + i.ToString()), ColumnChart.Type);
                }
                #endregion

                #region Add Category
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    //Setting Category Name
                    ColumnChart.ChartData.Categories.Add(workbook2.GetCell(defaultWorksheetIndex2, i + 1, 0, dt.Rows[i][0].ToString()));
                }

                #endregion

                #region Add Series Data
                //Add DataPoints for the Series
                for (int i = 1; i < dt.Columns.Count; i++)
                {
                    Series2 = ColumnChart.ChartData.Series[i - 1];

                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        string Val = dt.Rows[j][i].ToString();
                        Series2.DataPoints.AddDataPointForBarSeries(workbook2.GetCell(defaultWorksheetIndex2, j + 1, i, Convert.ToDouble(dt.Rows[j][i])));

                        //Set Data Point Label Style
                        lbl2 = Series2.DataPoints[j].Label;
                        lbl2.DataLabelFormat.Position = LegendDataLabelPosition.OutsideEnd;
                        lbl2.DataLabelFormat.ShowValue = true;
                        lbl2.DataLabelFormat.TextFormat.PortionFormat.FillFormat.FillType = FillType.Solid;
                        lbl2.DataLabelFormat.TextFormat.PortionFormat.FillFormat.SolidFillColor.Color = Color.Green;
                        lbl2.DataLabelFormat.TextFormat.PortionFormat.FontBold = NullableBool.True;
                        lbl2.DataLabelFormat.TextFormat.PortionFormat.FontHeight = 20;
                    }

                    //Set DataPoint Marker Style
                    Series2.Marker.Symbol = MarkerStyleType.Circle;
                    Series2.Marker.Size = 10;
                }

                #endregion

                #region Setting chart maximum, minimum values
                //Set chart Maximum, Minimum values
                ColumnChart.Axes.VerticalAxis.IsAutomaticMajorUnit = false;
                ColumnChart.Axes.VerticalAxis.IsAutomaticMaxValue = false;
                ColumnChart.Axes.VerticalAxis.IsAutomaticMinorUnit = false;
                ColumnChart.Axes.VerticalAxis.IsAutomaticMinValue = false;

                ColumnChart.Axes.VerticalAxis.MaxValue = 100;
                ColumnChart.Axes.VerticalAxis.MinValue = 70;
                ColumnChart.Axes.VerticalAxis.MinorUnit = 5;
                ColumnChart.Axes.VerticalAxis.MajorUnit = 10;
                #endregion

                #region Seting Value and Category Text Properties

                //Setting Value Axis Text Properties         
                IChartPortionFormat YaxisVal2 = ColumnChart.Axes.VerticalAxis.TextFormat.PortionFormat;
                YaxisVal2.FontBold = NullableBool.True;
                YaxisVal2.FontHeight = 20;
                YaxisVal2.FontItalic = NullableBool.True;
                YaxisVal2.FillFormat.FillType = FillType.Solid;
                YaxisVal2.FillFormat.SolidFillColor.Color = Color.Blue;
                YaxisVal2.LatinFont = new FontData("Times New Roman");


                //Setting Category Axis Text Properties
                IChartPortionFormat XAxisVal2 = ColumnChart.Axes.HorizontalAxis.TextFormat.PortionFormat;
                XAxisVal2.FontBold = NullableBool.True;
                XAxisVal2.FontHeight = 16;
                XAxisVal2.FontItalic = NullableBool.True;
                XAxisVal2.FillFormat.FillType = FillType.Solid; ;
                XAxisVal2.FillFormat.SolidFillColor.Color = Color.Blue;
                XAxisVal2.LatinFont = new FontData("Arial");

                #endregion

                #region Set Chart Axis Properties
                //Set VerticalAxis Axis GridLines Style
                ColumnChart.Axes.VerticalAxis.MajorGridLinesFormat.Line.FillFormat.FillType = FillType.Solid;
                ColumnChart.Axes.VerticalAxis.MajorGridLinesFormat.Line.FillFormat.SolidFillColor.Color = Color.Blue;
                ColumnChart.Axes.VerticalAxis.MajorGridLinesFormat.Line.DashStyle = LineDashStyle.Dot;

                //Set Vertical Axis Title
                ColumnChart.Axes.VerticalAxis.HasTitle = true;
                ColumnChart.Axes.VerticalAxis.Title.AddTextFrameForOverriding("Values in Percentage");
                ColumnChart.Axes.VerticalAxis.Title.TextFrameForOverriding.Paragraphs[0].Portions[0].PortionFormat.FontHeight = 20;
                ColumnChart.Axes.VerticalAxis.Title.TextFrameForOverriding.Paragraphs[0].Portions[0].PortionFormat.FillFormat.FillType = FillType.Solid;
                ColumnChart.Axes.VerticalAxis.Title.TextFrameForOverriding.Paragraphs[0].Portions[0].PortionFormat.FillFormat.SolidFillColor.Color = Color.DarkViolet;

                //Set Horizontal axis Title
                ColumnChart.Axes.HorizontalAxis.HasTitle = true;
                ColumnChart.Axes.HorizontalAxis.Title.AddTextFrameForOverriding("Year");
                ColumnChart.Axes.HorizontalAxis.Title.TextFrameForOverriding.Paragraphs[0].Portions[0].PortionFormat.FontHeight = 20;
                ColumnChart.Axes.HorizontalAxis.Title.TextFrameForOverriding.Paragraphs[0].Portions[0].PortionFormat.FillFormat.FillType = FillType.Solid;
                ColumnChart.Axes.HorizontalAxis.Title.TextFrameForOverriding.Paragraphs[0].Portions[0].PortionFormat.FillFormat.SolidFillColor.Color = Color.DarkViolet;
                #endregion

                #region Set Legend
                //Set Legend Style
                ColumnChart.HasLegend = true;
                ColumnChart.Legend.Position = LegendPositionType.Right;
                ColumnChart.Legend.TextFormat.PortionFormat.FontBold = NullableBool.True;
                ColumnChart.Legend.TextFormat.PortionFormat.FontHeight = 20;
                #endregion


                #endregion

                //Add the Slide to the presentation
                pres.Slides.InsertClone(1, slide2);

                var RandomValue = new Random();
                pres.Save(@"D:\Aspose PPT\Multiple Slide\AsposeMultiSlide" + RandomValue.Next(1, 99).ToString() + ".pptx", Aspose.Slides.Export.SaveFormat.Pptx);

            }
            catch (Exception ex)
            {
                LblError.Text = ex.Message.ToString();
            }
            
        }
    }
}