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
    public partial class AspostChartandTable : System.Web.UI.Page
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

                #region Chart1
               
                // Add chart with default data
                IChart chart = sld.Shapes.AddChart(ChartType.ClusteredColumn, 0, 300, 300, 200);

                //Setting chart Title
                //chart.ChartTitle.TextFrameForOverriding.Text = "Sample Title";
                chart.ChartTitle.AddTextFrameForOverriding("Sample Title");
                chart.ChartTitle.TextFrameForOverriding.TextFrameFormat.CenterText = NullableBool.True;
                chart.ChartTitle.Height = 20;
                chart.HasTitle = true;

                //Set first series to Show Values
                chart.ChartData.Series[0].Labels.DefaultDataLabelFormat.ShowValue = true;

                //Setting the index of chart data sheet
                int defaultWorksheetIndex = 0;

                //Getting the chart data worksheet
                IChartDataWorkbook fact = chart.ChartData.ChartDataWorkbook;

                //Delete default generated series and categories
                chart.ChartData.Series.Clear();
                chart.ChartData.Categories.Clear();
                int s = chart.ChartData.Series.Count;
                s = chart.ChartData.Categories.Count;

                //Adding new series
                chart.ChartData.Series.Add(fact.GetCell(defaultWorksheetIndex, 0, 1, "Series 1"), chart.Type);
                chart.ChartData.Series.Add(fact.GetCell(defaultWorksheetIndex, 0, 2, "Series 2"), chart.Type);

                //Adding new categories
                chart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, 1, 0, "Caetegoty 1"));
                chart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, 2, 0, "Caetegoty 2"));
                chart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, 3, 0, "Caetegoty 3"));

                //Take first chart series
                IChartSeries series = chart.ChartData.Series[0];

                //Now populating series data

                series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 1, 1, 20));
                series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 2, 1, 50));
                series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 3, 1, 30));

                //Setting fill color for series
                series.Format.Fill.FillType = FillType.Solid;
                series.Format.Fill.SolidFillColor.Color = Color.Red;


                //Take second chart series
                series = chart.ChartData.Series[1];

                //Now populating series data
                series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 1, 2, 30));
                series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 2, 2, 10));
                series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 3, 2, 60));

                //Setting fill color for series
                series.Format.Fill.FillType = FillType.Solid;
                series.Format.Fill.SolidFillColor.Color = Color.Green;


                //create custom labels for each of categories for new series

                //first label will be show Category name
                IDataLabel lbl = series.DataPoints[0].Label;
                lbl.DataLabelFormat.ShowCategoryName = true;

                lbl = series.DataPoints[1].Label;
                lbl.DataLabelFormat.ShowSeriesName = true;

                //Show value for third label
                lbl = series.DataPoints[2].Label;
                lbl.DataLabelFormat.ShowValue = true;
                lbl.DataLabelFormat.ShowSeriesName = true;
                lbl.DataLabelFormat.Separator = "/";
                #endregion

                #region Chart2

                // Add chart with default data
                 chart = sld.Shapes.AddChart(ChartType.Line, 310, 300, 400, 200);
                //Setting chart Title
                //chart.ChartTitle.TextFrameForOverriding.Text = "Sample Title";
                 chart.ChartTitle.AddTextFrameForOverriding("Line Chart");

                //chartTitle.PortionFormat.FillFormat.FillType = FillType.Solid;
                // chartTitle.PortionFormat.FillFormat.SolidFillColor.Color = Color.Gray;

                 chart.ChartTitle.TextFrameForOverriding.TextFrameFormat.CenterText = NullableBool.True;
                 chart.ChartTitle.Height = 5;
                 chart.HasTitle = true;

                //Set first series to Show Values
                 chart.ChartData.Series[0].Labels.DefaultDataLabelFormat.ShowValue = true;

                //Setting the index of chart data sheet
                //int defaultWorksheetIndex = 0;

                //Getting the chart data worksheet
                 IChartDataWorkbook fact2 = chart.ChartData.ChartDataWorkbook;

                //Delete default generated series and categories
                chart.ChartData.Series.Clear();
                chart.ChartData.Categories.Clear();
                int s2 = chart.ChartData.Series.Count;
                s2 = chart.ChartData.Categories.Count;

                //Adding new series
                chart.ChartData.Series.Add(fact2.GetCell(defaultWorksheetIndex, 0, 1, "Series 1"), chart.Type);
                chart.ChartData.Series.Add(fact2.GetCell(defaultWorksheetIndex, 0, 2, "Series 2"), chart.Type);

                //Adding new categories
                chart.ChartData.Categories.Add(fact2.GetCell(defaultWorksheetIndex, 1, 0, "Caetegoty 1"));
                chart.ChartData.Categories.Add(fact2.GetCell(defaultWorksheetIndex, 2, 0, "Caetegoty 2"));
                chart.ChartData.Categories.Add(fact2.GetCell(defaultWorksheetIndex, 3, 0, "Caetegoty 3"));

                //Take first chart series
                IChartSeries series2 = chart.ChartData.Series[0];

                //Now populating series data

                series2.DataPoints.AddDataPointForLineSeries(fact2.GetCell(defaultWorksheetIndex, 1, 1, 20));
                series2.DataPoints.AddDataPointForLineSeries(fact2.GetCell(defaultWorksheetIndex, 2, 1, 50));
                series2.DataPoints.AddDataPointForLineSeries(fact2.GetCell(defaultWorksheetIndex, 3, 1, 30));

                //Setting fill color for series
                series2.Format.Fill.FillType = FillType.Solid;
                series2.Format.Fill.SolidFillColor.Color = Color.Red;


                //Take second chart series
                series2 = chart.ChartData.Series[1];

                //Now populating series data
                series2.DataPoints.AddDataPointForLineSeries(fact2.GetCell(defaultWorksheetIndex, 1, 2, 30));
                series2.DataPoints.AddDataPointForLineSeries(fact2.GetCell(defaultWorksheetIndex, 2, 2, 10));
                series2.DataPoints.AddDataPointForLineSeries(fact2.GetCell(defaultWorksheetIndex, 3, 2, 60));

                //Setting fill color for series
                series2.Format.Fill.FillType = FillType.Solid;
                series2.Format.Fill.SolidFillColor.Color = Color.Green;

                chart.Axes.VerticalAxis.IsNumberFormatLinkedToSource = false;

                //Setting Value Axis Text Properties
                IChartPortionFormat txtVal = chart.Axes.VerticalAxis.TextFormat.PortionFormat;
                txtVal.FontBold = NullableBool.True;
                txtVal.FontHeight = 16;
                txtVal.FontItalic = NullableBool.True;
                txtVal.FillFormat.FillType = FillType.Solid; ;
                txtVal.FillFormat.SolidFillColor.Color = Color.DarkGreen;
                txtVal.LatinFont = new FontData("Times New Roman");


                IChartPortionFormat txtValhorz = chart.Axes.HorizontalAxis.TextFormat.PortionFormat;
                txtValhorz.FontBold = NullableBool.True;
                txtValhorz.FontHeight = 12;
                txtValhorz.FontItalic = NullableBool.True;
                txtValhorz.FillFormat.FillType = FillType.Solid; ;
                txtValhorz.FillFormat.SolidFillColor.Color = Color.DarkGreen;
                txtValhorz.LatinFont = new FontData("Times New Roman");



                //Setting Category Axis Text Properties
                IChartPortionFormat txtCat = chart.Axes.HorizontalAxis.TextFormat.PortionFormat;
                txtCat.FontBold = NullableBool.True;
                txtCat.FontHeight = 16;
                txtCat.FontItalic = NullableBool.True;
                txtCat.FillFormat.FillType = FillType.Solid; ;
                txtCat.FillFormat.SolidFillColor.Color = Color.Blue;
                txtCat.LatinFont = new FontData("Arial");


                //create custom labels for each of categories for new series

                //first label will be show Category name
                IDataLabel lbl2 = series2.DataPoints[0].Label;
                lbl2.DataLabelFormat.ShowCategoryName = true;

                lbl2 = series2.DataPoints[1].Label;
                lbl2.DataLabelFormat.ShowSeriesName = true;

                //Show value for third label
                lbl2 = series2.DataPoints[2].Label;
                lbl2.DataLabelFormat.ShowValue = true;
                lbl2.DataLabelFormat.ShowSeriesName = true;
                lbl2.DataLabelFormat.Separator = "/";
                #endregion

                #region Table
                //ISlide sld = pres.Slides[0];

                //Define columns with widths and rows with heights
                double[] dblCols = { 100, 100, 100, 100, 100, 100, 100 };
                double[] dblRows = { 50, 30, 30, 30, 30, 30, 30, 30 };

                //Add table shape to slide
                ITable tbl = sld.Shapes.AddTable(10, 10, dblCols, dblRows);

                //Set border format for each cell
                foreach (IRow row in tbl.Rows)
                    foreach (ICell cell in row)
                    {
                        cell.BorderTop.FillFormat.FillType = FillType.Solid;
                        cell.BorderTop.FillFormat.SolidFillColor.Color = Color.Red;
                        cell.BorderTop.Width = 5;

                        cell.BorderBottom.DashStyle = LineDashStyle.DashDot;
                        cell.BorderBottom.FillFormat.FillType = FillType.Solid;
                        cell.BorderBottom.FillFormat.SolidFillColor.Color = Color.Blue;
                        cell.BorderBottom.Width = 5;

                        cell.BorderLeft.FillFormat.FillType = FillType.Solid;
                        cell.BorderLeft.FillFormat.SolidFillColor.Color = Color.Red;
                        cell.BorderLeft.Width = 5;

                        cell.BorderRight.DashStyle = LineDashStyle.DashDot;
                        cell.BorderRight.FillFormat.FillType = FillType.Solid;
                        cell.BorderRight.FillFormat.SolidFillColor.Color = Color.Blue;
                        cell.BorderRight.Width = 5;
                    }

                //Merge cells 1 & 2 of row 1
                tbl.MergeCells(tbl[0, 0], tbl[1, 0], false);

                //Add text to the merged cell
                tbl[0, 0].TextFrame.Text = "Merged Cells";
                #endregion

                //var val = new Random();
                //pres.Save(@"D:\Aspose PPT\AsposeCharttest" + val.Next(1, 99).ToString() + ".pptx", Aspose.Slides.Export.SaveFormat.Pptx);

                DownloadAspose(pres);

            }
            catch (Exception ex)
            {
                LblError.Text = ex.Message.ToString();
            }
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