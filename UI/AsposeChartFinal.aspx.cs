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
using System.IO;

namespace AsposeTest.UI
{
    public partial class AsposeChartFinal : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            GetPPT();
        }

        private void GetPPT()
        {
            //Instantiate the License class
            Aspose.Slides.License license = new Aspose.Slides.License();

            //Pass only the name of the license file embedded in the assembly
            license.SetLicense("Aspose.Slides.lic");

            Presentation pres = new Presentation();

            //Access First Slide
            ISlide Slide1 = pres.Slides[0];

            //Add Chart

            IChart LineChart = Slide1.Shapes.AddChart(ChartType.LineWithMarkers, 5, 5, 500, 500);

            //Setting Chart Title

            LineChart.ChartTitle.AddTextFrameForOverriding("Line Chart");
            LineChart.ChartTitle.TextFrameForOverriding.TextFrameFormat.CenterText = NullableBool.True;
            LineChart.ChartTitle.Height = 20;
            LineChart.ChartTitle.TextFormat.PortionFormat.FontBold = NullableBool.True;            
            LineChart.HasTitle = true;

            //Getting the chart data worksheet
            IChartDataWorkbook fact2 = LineChart.ChartData.ChartDataWorkbook;

            //Setting the index of chart data sheet
            int defaultWorksheetIndex = 0;

            //Delete default generated series and categories
            LineChart.ChartData.Series.Clear();
            LineChart.ChartData.Categories.Clear();

            //Set Series & Name
            LineChart.ChartData.Series.Add(fact2.GetCell(defaultWorksheetIndex, 0, 1, "Series 1"), LineChart.Type);
            LineChart.ChartData.Series.Add(fact2.GetCell(defaultWorksheetIndex, 0, 2, "Series 2"), LineChart.Type);
            LineChart.ChartData.Series.Add(fact2.GetCell(defaultWorksheetIndex, 0, 3, "Series 3"), LineChart.Type);

            //set Category & Category Name
            LineChart.ChartData.Categories.Add(fact2.GetCell(defaultWorksheetIndex, 1, 0, "cat 1"));
            LineChart.ChartData.Categories.Add(fact2.GetCell(defaultWorksheetIndex, 2, 0, "cat 2"));
            LineChart.ChartData.Categories.Add(fact2.GetCell(defaultWorksheetIndex, 3, 0, "Caet 3"));


            #region Setting chart maximum, minimum values
            //Setting chart maximum, minimum values
            LineChart.Axes.VerticalAxis.IsAutomaticMajorUnit = false;
            LineChart.Axes.VerticalAxis.IsAutomaticMaxValue = false;
            LineChart.Axes.VerticalAxis.IsAutomaticMinorUnit = false;
            LineChart.Axes.VerticalAxis.IsAutomaticMinValue = false;

            LineChart.Axes.VerticalAxis.MaxValue = 60;
            LineChart.Axes.VerticalAxis.MinValue = 1;
            LineChart.Axes.VerticalAxis.MinorUnit = 5;
            LineChart.Axes.VerticalAxis.MajorUnit = 10;
            #endregion

            IChartSeries Series;

            #region Setting Series Data
            //Take first chart series
            for (int i = 0; i < 3; i++)
            {
                Series = LineChart.ChartData.Series[i];

                //add data
                Series.DataPoints.AddDataPointForLineSeries(fact2.GetCell(defaultWorksheetIndex, 1, i + 1, (i + 1) * 5));
                Series.DataPoints.AddDataPointForLineSeries(fact2.GetCell(defaultWorksheetIndex, 2, i + 1, (i + 1) * 10));
                Series.DataPoints.AddDataPointForLineSeries(fact2.GetCell(defaultWorksheetIndex, 3, i + 1, (i + 1) * 15));

                //Set Marker Style
                Series.Marker.Symbol = MarkerStyleType.Circle;
                Series.Marker.Size = 10;

                IDataLabel lbl;
                for (int j = 0; j < 3; j++)
                {
                    //label will be show the value
                    lbl = Series.DataPoints[j].Label;
                    lbl.DataLabelFormat.ShowCategoryName = false;
                    lbl.DataLabelFormat.ShowSeriesName = false;
                    lbl.TextFormat.PortionFormat.FontHeight = 20;
                    lbl.TextFormat.PortionFormat.FontBold = NullableBool.True;
                    lbl.TextFormat.PortionFormat.FillFormat.FillType = FillType.Solid;
                    lbl.TextFormat.PortionFormat.FillFormat.SolidFillColor.Color = Color.Red;


                    //Data Label Position
                    lbl.DataLabelFormat.Position = i == 0 ? LegendDataLabelPosition.Right : i == 1 ? LegendDataLabelPosition.Left : LegendDataLabelPosition.Top;

                    //Data Point Backcolor for First Series
                    if (j == 0)
                        lbl.TextFormat.PortionFormat.HighlightColor.Color = Color.LightBlue;

                    if (i == 1)
                    {
                        lbl.DataLabelFormat.ShowValue = false;
                    }
                    else
                    {
                        lbl.DataLabelFormat.ShowValue = true;
                    }
                }


                ////Setting Points Styles
                //IChartDataPoint Point;
                //for(int m=0;m<3;m++)
                //{
                //    Point = Series.DataPoints[m];
                //    Point.Format.Fill.FillType = FillType.Solid;
                //    Point.Format.Fill.SolidFillColor.Color = Color.Brown;
                //    Point.Label.DataLabelFormat.ShowValue = true;
                    
                //    //Setting Sector border
                //    Point.Format.Line.FillFormat.FillType = FillType.Solid;
                //    Point.Format.Line.FillFormat.SolidFillColor.Color = Color.Blue;
                //    Point.Format.Line.Width = 3.0;
                //    Point.Format.Line.Style = LineStyle.Single;
                //    Point.Format.Line.DashStyle = LineDashStyle.LargeDashDot;

                //}

            }

            #endregion


            #region Seeting Value and Category Text Properties
            //Setting Value Axis Text Properties
            IChartPortionFormat txtval = LineChart.Axes.VerticalAxis.TextFormat.PortionFormat;
            txtval.FontBold = NullableBool.True;
            txtval.FontHeight = 20;
            txtval.FontItalic = NullableBool.True;
            txtval.FillFormat.FillType = FillType.Solid;
            txtval.FillFormat.SolidFillColor.Color = Color.Green;
            txtval.LatinFont = new FontData("Times New Roman");

            //Setting Category Axis Text Properties
            IChartPortionFormat txtCat = LineChart.Axes.HorizontalAxis.TextFormat.PortionFormat;
            txtCat.FontBold = NullableBool.True;
            txtCat.FontHeight = 16;
            txtCat.FontItalic = NullableBool.True;
            txtCat.FillFormat.FillType = FillType.Solid; ;
            txtCat.FillFormat.SolidFillColor.Color = Color.Blue;
            txtCat.LatinFont = new FontData("Arial");
            
            #endregion

       
            

            //var val = new Random();
            //pres.Save("D:/Projects/AsposeTest/AsposeTest/Aspose PPT/AsposeChartFinal" + val.Next(1, 99).ToString() + ".pptx", Aspose.Slides.Export.SaveFormat.Pptx);

           


              
                //Setting the content type of the Http Response
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

        //public void DownloadFile(Presentation pres, string FileName)
        //{
        //    if (pres != null)
        //    {
        //        HttpResponse httpResponse = HttpContext.Current.Response;
        //        httpResponse.Clear();
        //        httpResponse.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        //        httpResponse.AddHeader("content-disposition", "attachment;filename=\"" + FileName + ".xlsx\"");

        //        using (MemoryStream memoryStream = new MemoryStream())
        //        {
                    
        //            pres.Save(memoryStream);
        //            memoryStream.WriteTo(httpResponse.OutputStream);
        //            memoryStream.Close();
        //        }

        //        httpResponse.End();

        //    }
        //    else
        //       // DownloadExcelFile(NoDataExcel());
        //}
    }
}