using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Aspose.Slides;
using Aspose.Slides.Export;
using Aspose.Slides.Charts;
using System.Drawing;

namespace AsposeTest.UI
{
    public partial class Asposechartonly : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                GetPPT();
            }
        }

        private void GetPPT()
        {
            try
            {
                //Instantiate Presentation class that represents PPTX file
                Presentation pres = new Presentation();

                //Access first slide
                ISlide sld = pres.Slides[0];

                // Add chart with default data
                IChart chart2 = sld.Shapes.AddChart(ChartType.LineWithMarkers, 10, 10, 500, 500);
                //Setting chart Title
                //chart.ChartTitle.TextFrameForOverriding.Text = "Sample Title";
                chart2.ChartTitle.AddTextFrameForOverriding("Line Chart");
                chart2.ChartTitle.TextFrameForOverriding.TextFrameFormat.CenterText = NullableBool.True;
                chart2.ChartTitle.Height = 5;
                chart2.HasTitle = true;


                //Set first series to Show Values
                chart2.ChartData.Series[0].Labels.DefaultDataLabelFormat.ShowValue = true;

                //Setting the index of chart data sheet
                int defaultWorksheetIndex = 0;

                //Getting the chart data worksheet
                IChartDataWorkbook fact2 = chart2.ChartData.ChartDataWorkbook;

                //Delete default generated series and categories
                chart2.ChartData.Series.Clear();
                chart2.ChartData.Categories.Clear();
                //int s2 = chart2.ChartData.Series.Count;
                //s2 = chart2.ChartData.Categories.Count;



                #region Setting chart maximum, minimum values
                //Setting chart maximum, minimum values
                chart2.Axes.VerticalAxis.IsAutomaticMajorUnit = false;
                chart2.Axes.VerticalAxis.IsAutomaticMaxValue = false;
                chart2.Axes.VerticalAxis.IsAutomaticMinorUnit = false;
                chart2.Axes.VerticalAxis.IsAutomaticMinValue = false;

                chart2.Axes.VerticalAxis.MaxValue = 60;
                chart2.Axes.VerticalAxis.MinValue = 10;
                chart2.Axes.VerticalAxis.MinorUnit = 5;
                chart2.Axes.VerticalAxis.MajorUnit = 10; 
                #endregion


                #region Add Series and Categories
                //Adding new series
                chart2.ChartData.Series.Add(fact2.GetCell(defaultWorksheetIndex, 0, 1, "Series 1"), chart2.Type);
                chart2.ChartData.Series.Add(fact2.GetCell(defaultWorksheetIndex, 0, 2, "Series 2"), chart2.Type);
                chart2.ChartData.Series.Add(fact2.GetCell(defaultWorksheetIndex, 0, 3, "Series 3"), chart2.Type);

                //Adding new categories
                chart2.ChartData.Categories.Add(fact2.GetCell(defaultWorksheetIndex, 1, 0, "category 1"));
                chart2.ChartData.Categories.Add(fact2.GetCell(defaultWorksheetIndex, 2, 0, "category 2"));
                chart2.ChartData.Categories.Add(fact2.GetCell(defaultWorksheetIndex, 3, 0, "Caetegoty 3"));
                #endregion

                #region FirstChartSeries
                //Take first chart series
                IChartSeries series2 = chart2.ChartData.Series[0];

                //Now populating series data

                series2.DataPoints.AddDataPointForLineSeries(fact2.GetCell(defaultWorksheetIndex, 1, 1, 45));
                series2.DataPoints.AddDataPointForLineSeries(fact2.GetCell(defaultWorksheetIndex, 2, 1, 50));
                series2.DataPoints.AddDataPointForLineSeries(fact2.GetCell(defaultWorksheetIndex, 3, 1, 30));

                //Setting fill color for series
                series2.Format.Fill.FillType = FillType.Solid;
                series2.Format.Fill.SolidFillColor.Color = Color.Red;


                //create custom labels for each of categories for new series

                //first label will be show Category name
                IDataLabel lbl2 = series2.DataPoints[0].Label;
                lbl2.DataLabelFormat.ShowCategoryName = false;
                lbl2.DataLabelFormat.ShowValue = true;
                lbl2.DataLabelFormat.ShowSeriesName = false;


                lbl2 = series2.DataPoints[1].Label;
                lbl2.DataLabelFormat.ShowCategoryName = false;
                lbl2.DataLabelFormat.ShowValue = true;
                lbl2.DataLabelFormat.ShowSeriesName = false;

                //Show value for third label
                lbl2 = series2.DataPoints[2].Label;
                lbl2.DataLabelFormat.ShowCategoryName = false;
                lbl2.DataLabelFormat.ShowValue = true;
                lbl2.DataLabelFormat.ShowSeriesName = false;
                lbl2.DataLabelFormat.Position = LegendDataLabelPosition.Right;
                lbl2.TextFormat.PortionFormat.FillFormat.SolidFillColor.Color = Color.Red;
                //lbl2.DataLabelFormat.Separator = "/"; 

                #endregion


                #region Second Series

                //Take second chart series
                series2 = chart2.ChartData.Series[1];

                //Now populating series data
                series2.DataPoints.AddDataPointForLineSeries(fact2.GetCell(defaultWorksheetIndex, 1, 2, 30));
                series2.DataPoints.AddDataPointForLineSeries(fact2.GetCell(defaultWorksheetIndex, 2, 2, 40));
                series2.DataPoints.AddDataPointForLineSeries(fact2.GetCell(defaultWorksheetIndex, 3, 2, 60));

                //Setting fill color for series
                series2.Format.Fill.FillType = FillType.Solid;
                series2.Format.Fill.SolidFillColor.Color = Color.Green;

                chart2.Axes.VerticalAxis.IsNumberFormatLinkedToSource = false;



                //create custom labels for each of categories for new series

                //first label will be show Category name
                IDataLabel lbl1 = series2.DataPoints[0].Label;
                lbl1.DataLabelFormat.ShowCategoryName = false;
                lbl1.DataLabelFormat.ShowValue = true;
                lbl1.DataLabelFormat.ShowSeriesName = false;


                lbl1 = series2.DataPoints[1].Label;
                lbl1.DataLabelFormat.ShowCategoryName = false;
                lbl1.DataLabelFormat.ShowValue = true;
                lbl1.DataLabelFormat.ShowSeriesName = false;

                //Show value for third label
                lbl1 = series2.DataPoints[2].Label;
                lbl1.DataLabelFormat.ShowCategoryName = false;
                lbl1.DataLabelFormat.ShowValue = false;
                lbl1.DataLabelFormat.ShowSeriesName = false;
                //lbl2.DataLabelFormat.Separator = "/"; 
                #endregion


                #region Third Series
                //Take second chart series
                series2 = chart2.ChartData.Series[2];
                
                //Now populating series data
                series2.DataPoints.AddDataPointForLineSeries(fact2.GetCell(defaultWorksheetIndex, 1, 3, 10));
                series2.DataPoints.AddDataPointForLineSeries(fact2.GetCell(defaultWorksheetIndex, 2, 3, 20));
                series2.DataPoints.AddDataPointForLineSeries(fact2.GetCell(defaultWorksheetIndex, 3, 3, 25));

                //Setting fill color for series
                series2.Format.Fill.FillType = FillType.Solid;
                series2.Format.Fill.SolidFillColor.Color = Color.Green;

                chart2.Axes.VerticalAxis.IsNumberFormatLinkedToSource = false;



                //create custom labels for each of categories for new series

                //first label will be show Category name
                IDataLabel lbl3 = series2.DataPoints[0].Label;
                lbl3.DataLabelFormat.ShowCategoryName = false;
                lbl3.DataLabelFormat.ShowValue = true;
                lbl3.DataLabelFormat.ShowSeriesName = false;


                lbl3 = series2.DataPoints[1].Label;
                lbl3.DataLabelFormat.ShowCategoryName = false;
                lbl3.DataLabelFormat.ShowValue = true;
                lbl3.DataLabelFormat.ShowSeriesName = false;

                //Show value for third label
                lbl3 = series2.DataPoints[2].Label;
                lbl3.DataLabelFormat.ShowCategoryName = false;
                lbl3.DataLabelFormat.ShowValue = true;
                lbl3.DataLabelFormat.ShowSeriesName = false;
                //lbl2.DataLabelFormat.Separator = "/"; 
                #endregion

                chart2.ChartData.Series[2].Marker.Symbol = MarkerStyleType.Circle;
                chart2.ChartData.Series[2].Marker.Size = 15;


                #region Value and categories Text Properties
                //Setting Value Axis Text Properties
                IChartPortionFormat txtVal = chart2.Axes.VerticalAxis.TextFormat.PortionFormat;
                txtVal.FontBold = NullableBool.True;
                txtVal.FontHeight = 16;
                txtVal.FontItalic = NullableBool.True;
                txtVal.FillFormat.FillType = FillType.Solid; ;
                txtVal.FillFormat.SolidFillColor.Color = Color.DarkGreen;
                txtVal.LatinFont = new FontData("Times New Roman");


                //Setting Category Axis Text Properties
                IChartPortionFormat txtCat = chart2.Axes.HorizontalAxis.TextFormat.PortionFormat;
                txtCat.FontBold = NullableBool.True;
                txtCat.FontHeight = 16;
                txtCat.FontItalic = NullableBool.True;
                txtCat.FillFormat.FillType = FillType.Solid; ;
                txtCat.FillFormat.SolidFillColor.Color = Color.Blue;
                txtCat.LatinFont = new FontData("Arial");

                #endregion


                var val = new Random();
                pres.Save(@"D:\Aspose PPT\AsposeChartonly" + val.Next(1, 99).ToString() + ".pptx", Aspose.Slides.Export.SaveFormat.Pptx);

            }
            catch (Exception ex)
            {
                LblError.Text = ex.Message.ToString();
            }
        }
    }
}