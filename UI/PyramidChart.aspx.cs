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
    public partial class PyramidChart : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            GetPPT();
        }

        private void GetPPT()
        {
            //Instantiate Presentation class that represents PPTX file//Instantiate Presentation class that represents PPTX file
            Presentation pres = new Presentation();

            //Access first slide
            ISlide sld = pres.Slides[0];

            // Add chart with default data
            IChart chart = sld.Shapes.AddChart(ChartType.StackedBar, 100, 100, 200, 200);

            //Setting chart Title
           
            chart.ChartTitle.AddTextFrameForOverriding("Pyramid Chart");
            chart.ChartTitle.TextFormat.PortionFormat.FontBold = NullableBool.True;
            chart.ChartTitle.Height = 20;
            chart.HasTitle = true;
            chart.HasLegend = false;
           
            //Set first series to Show Values
            chart.ChartData.Series[0].Labels.DefaultDataLabelFormat.ShowValue = true;


            //Setting the index of chart data sheet
            int defaultWorksheetIndex = 0;

            //Getting the chart data worksheet
            IChartDataWorkbook fact = chart.ChartData.ChartDataWorkbook;



            //Delete default generated series and categories
            chart.ChartData.Series.Clear();
            chart.ChartData.Categories.Clear();

            //#region Setting chart maximum, minimum values
            ////Setting chart maximum, minimum values
            //chart.Axes.VerticalAxis.IsAutomaticMajorUnit = false;
            //chart.Axes.VerticalAxis.IsAutomaticMaxValue = false;
            //chart.Axes.VerticalAxis.IsAutomaticMinorUnit = false;
            //chart.Axes.VerticalAxis.IsAutomaticMinValue = false;

            //chart.Axes.VerticalAxis.MaxValue = 60;
            //chart.Axes.VerticalAxis.MinValue = 1;
            //chart.Axes.VerticalAxis.MinorUnit = 5;
            //chart.Axes.VerticalAxis.MajorUnit = 10;

            //chart.Axes.VerticalAxis.IsAutomaticMajorUnit = false;
            //chart.Axes.VerticalAxis.IsAutomaticMaxValue = false;
            //chart.Axes.VerticalAxis.IsAutomaticMinorUnit = false;
            //chart.Axes.VerticalAxis.IsAutomaticMinValue = false;

            //chart.Axes.VerticalAxis.MaxValue = 100;
            //chart.Axes.VerticalAxis.MinValue = 0;
            //chart.Axes.VerticalAxis.MinorUnit = 5;
            //chart.Axes.VerticalAxis.MajorUnit = 10;

            //chart.Axes.VerticalAxis.IsAutomaticTickLabelSpacing = false;
            //chart.Axes.VerticalAxis.TickLabelSpacing = 20;

             

            //#endregion

            //Adding new series
            chart.ChartData.Series.Add(fact.GetCell(defaultWorksheetIndex, 0, 1, "Series 1"), chart.Type);
            chart.ChartData.Series.Add(fact.GetCell(defaultWorksheetIndex, 0, 2, "Series 2"), chart.Type);
            chart.ChartData.Series.Add(fact.GetCell(defaultWorksheetIndex, 0, 3, "Series 3"), chart.Type);
            //Adding new categories
            chart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, 1, 0, "Caetegoty 1"));
            chart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, 2, 0, "Caetegoty 2"));
            chart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, 3, 0, "Caetegoty 3"));
            chart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, 4, 0, "Caetegoty 4"));

            //Take first chart series
            IChartSeries series1 = chart.ChartData.Series[0];
            IChartSeries series2 = chart.ChartData.Series[1];
            IChartSeries series3 = chart.ChartData.Series[2];

            //Now populating series data
            var valuePoint = 50;

            //Category 1
            series1.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 0, 1, (100 - valuePoint) / 2));
            series2.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 1, 1, valuePoint));
            series3.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 2, 1, (100 - valuePoint) / 2));
            
            //Category 2
            valuePoint = 40;
            series1.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 0, 2, (100 - valuePoint) / 2));
            series2.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 1, 2, valuePoint));
            series3.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 2, 2, (100 - valuePoint) / 2));

            //Category 3
            valuePoint = 30;
            series1.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 0, 3, (100 - valuePoint) / 2));
            series2.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 1, 3, valuePoint));
            series3.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 2, 3, (100 - valuePoint) / 2));


            valuePoint = 10;
            series1.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 0, 4, (100 - valuePoint) / 2));
            series2.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 1, 4, valuePoint));
            series3.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 2, 4, (100 - valuePoint) / 2));

            IDataLabel Lbl;
            for (int i = 0; i < 4; i++)
            {
                Lbl = series2.DataPoints[i].Label;
                Lbl.DataLabelFormat.ShowValue = true;
                Lbl.DataLabelFormat.Position = LegendDataLabelPosition.Center;
                Lbl.TextFormat.PortionFormat.FontBold = NullableBool.True;
                Lbl.TextFormat.PortionFormat.FontHeight = 20;
                Lbl.TextFormat.PortionFormat.FillFormat.FillType = FillType.Solid;
                Lbl.TextFormat.PortionFormat.FillFormat.SolidFillColor.Color = Color.White;

            }

            //Setting fill color for series
            series1.Format.Fill.FillType = FillType.Solid;
            series1.Format.Fill.SolidFillColor.Color = Color.Transparent;

            series2.Format.Fill.FillType = FillType.Solid;
            series2.Format.Fill.SolidFillColor.Color = Color.Green;

            series3.Format.Fill.FillType = FillType.Solid;
            series3.Format.Fill.SolidFillColor.Color = Color.Transparent;

            series1.ParentSeriesGroup.GapWidth = 0;
            series2.ParentSeriesGroup.GapWidth = 0;
            series3.ParentSeriesGroup.GapWidth = 0;

            //series1.ParentSeriesGroup.GapDepth = 0;
            //series2.ParentSeriesGroup.GapDepth = 50;
            //series3.ParentSeriesGroup.GapDepth = 0;

            series1.ParentSeriesGroup.Overlap = 100;
            series2.ParentSeriesGroup.Overlap = 100;
            series3.ParentSeriesGroup.Overlap = 100;

            //Take second chart series
            //   series = chart.ChartData.Series[1];

            //Now populating series data
            //series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 1, 1, 30));
            //series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 2, 1, 10));
            //series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 3, 1, 60));

            //Setting fill color for series
            //series.Format.Fill.FillType = FillType.Solid;
            //series.Format.Fill.SolidFillColor.Color = Color.Green;

            chart.Axes.VerticalAxis.MajorGridLinesFormat.Line.FillFormat.FillType = FillType.NoFill;
            chart.Axes.VerticalAxis.MinorGridLinesFormat.Line.FillFormat.FillType = FillType.NoFill;

            chart.Axes.HorizontalAxis.MajorGridLinesFormat.Line.FillFormat.FillType = FillType.NoFill;
            chart.Axes.HorizontalAxis.MinorGridLinesFormat.Line.FillFormat.FillType = FillType.NoFill;

            //chart.Axes.HorizontalAxis.IsVisible = false;
            //chart.Axes.VerticalAxis.IsVisible = false;


            IChartPortionFormat txtval = chart.Axes.VerticalAxis.TextFormat.PortionFormat;
            txtval.FontBold = NullableBool.True;
            txtval.FontHeight = 20;
            txtval.FontItalic = NullableBool.True;
            txtval.FillFormat.FillType = FillType.Solid;
            txtval.FillFormat.SolidFillColor.Color = Color.Green;
            txtval.LatinFont = new FontData("Times New Roman");

             txtval = chart.Axes.HorizontalAxis.TextFormat.PortionFormat;
            txtval.FontBold = NullableBool.True;
            txtval.FontHeight = 20;
            txtval.FontItalic = NullableBool.True;
            txtval.FillFormat.FillType = FillType.Solid;
            txtval.FillFormat.SolidFillColor.Color = Color.Blue;
            txtval.LatinFont = new FontData("Times New Roman");

            chart.Axes.HorizontalAxis.IsVisible = false;

           
            //create custom labels for each of categories for new series

            //first label will be show Category name
            //IDataLabel lbl = series.DataPoints[0].Label;
            //lbl.DataLabelFormat.ShowCategoryName = false;
            //lbl.DataLabelFormat.ShowValue = true;
            //lbl.DataLabelFormat.ShowSeriesName = false;

            //lbl = series.DataPoints[1].Label;
            //lbl.DataLabelFormat.ShowSeriesName = true;

            ////Show value for third label
            //lbl = series.DataPoints[2].Label;
            //lbl.DataLabelFormat.ShowValue = true;
            //lbl.DataLabelFormat.ShowSeriesName = true;
            //lbl.DataLabelFormat.Separator = "/";


            DownloadAspose(pres);
        }

        private void DownloadAspose(Presentation pres)
        {
            this.Response.ContentType = "application/vnd.ms-powerpoint";

            //Appending the header of the Http Response to contain the presentation file name
            this.Response.AppendHeader("Content-Disposition", "attachment; filename=demo.pptx");

            //Flushing the buffers of Http Response
            this.Response.Flush();

            //Accessing the output stream of Http Response
            System.IO.Stream st = this.Response.OutputStream;

            //Saving the presentation to the output stream of Http Response
            pres.Save(st, Aspose.Slides.Export.SaveFormat.Pptx);

            //Closing the Http Response
            this.Response.End();
        }



    }
}