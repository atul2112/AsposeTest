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
    public partial class AsposeTableOnly : System.Web.UI.Page
    {
        string[] Colorcode = { "#FF3300", "#3366FF", "#FFCC00", "#47B547" };
        string[] FirstRowColor = { "#989898", "#484848" };

        List<string> groupNames = new List<string>();
        List<string> brandNames = new List<string>();

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

            //Instantiate Presentation class that represents PPTX file
            Presentation pres = new Presentation();

            //Access first slide
            ISlide sld = pres.Slides[0];

            #region Table
            //ISlide sld = pres.Slides[0];

            //Define columns with widths and rows with heights
            double[] dblCols = { 80, 100, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45 };
            double[] dblRows = { 40, 50, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30 };

            //Add table shape to slide
            ITable tbl = sld.Shapes.AddTable(10, 10, dblCols, dblRows);

            //Set border format for each cell
            foreach (IRow row in tbl.Rows)
            {
                foreach (ICell cell in row)
                {
                    cell.BorderTop.FillFormat.FillType = FillType.Solid;
                    cell.BorderTop.FillFormat.SolidFillColor.Color = Color.Black;
                    cell.BorderTop.Width = 2;

                    cell.BorderBottom.FillFormat.FillType = FillType.Solid;
                    cell.BorderBottom.FillFormat.SolidFillColor.Color = Color.Black;
                    cell.BorderBottom.Width = 2;

                    cell.BorderLeft.FillFormat.FillType = FillType.Solid;
                    cell.BorderLeft.FillFormat.SolidFillColor.Color = Color.Black;
                    cell.BorderLeft.Width = 2;

                    cell.BorderRight.FillFormat.FillType = FillType.Solid;
                    cell.BorderRight.FillFormat.SolidFillColor.Color = Color.Black;
                    cell.BorderRight.Width = 2;                  

                }

            }


            #region Merge
            //settings Merge

            //Row 1
            tbl.MergeCells(tbl[0, 0], tbl[1, 0], false);

            tbl.MergeCells(tbl[2, 0], tbl[3, 0], false);
            tbl.MergeCells(tbl[2, 0], tbl[4, 0], false);
            tbl.MergeCells(tbl[2, 0], tbl[5, 0], false);

            tbl.MergeCells(tbl[6, 0], tbl[7, 0], false);
            tbl.MergeCells(tbl[6, 0], tbl[8, 0], false);
            tbl.MergeCells(tbl[6, 0], tbl[9, 0], false);

            tbl.MergeCells(tbl[10, 0], tbl[11, 0], false);
            tbl.MergeCells(tbl[10, 0], tbl[12, 0], false);
            tbl.MergeCells(tbl[10, 0], tbl[13, 0], false);

            //Row 2
            tbl.MergeCells(tbl[0, 1], tbl[1, 1], false);


            //Row 4
            tbl.MergeCells(tbl[0, 3], tbl[0, 4], false);
            tbl.MergeCells(tbl[0, 3], tbl[0, 5], false);
            tbl.MergeCells(tbl[0, 3], tbl[0, 5], false);
            tbl.MergeCells(tbl[0, 3], tbl[0, 6], false);

            //Row 8
            tbl.MergeCells(tbl[0, 7], tbl[0, 8], true);
            tbl.MergeCells(tbl[0, 7], tbl[0, 9], true);
            tbl.MergeCells(tbl[0, 7], tbl[0, 10], true);
            tbl.MergeCells(tbl[0, 7], tbl[0, 11], true);
            tbl.MergeCells(tbl[0, 7], tbl[0, 12], true);
            #endregion


            //Add Text to the cell
            //tbl.StylePreset = Aspose.Slides.TableStylePreset.LightStyle2Accent3;

            tbl[2, 0].TextFrame.Text = "Total Population";
            tbl[2, 0].FillFormat.FillType = FillType.Solid;
            tbl[2, 0].FillFormat.SolidFillColor.Color = ColorTranslator.FromHtml(FirstRowColor[0]);

            tbl[6, 0].TextFrame.Text = "African American";
            tbl[6, 0].FillFormat.FillType = FillType.Solid;
            tbl[6, 0].FillFormat.SolidFillColor.Color = ColorTranslator.FromHtml(FirstRowColor[1]);

            tbl[10, 0].TextFrame.Text = "Hispanions";
            tbl[10, 0].FillFormat.FillType = FillType.Solid;
            tbl[10, 0].FillFormat.SolidFillColor.Color = ColorTranslator.FromHtml(FirstRowColor[0]);

            tbl[0, 1].TextFrame.Text = "";
            tbl[0, 2].TextFrame.Text = "Awareness";
            tbl[0, 2].TextFrame.Paragraphs[0].Portions[0].PortionFormat.FontHeight = 10;
            tbl[0, 3].TextFrame.Text = "Frequency";
            tbl[0, 3].TextFrame.Paragraphs[0].Portions[0].PortionFormat.FontHeight = 10;
            tbl[0, 7].TextFrame.Text = "Imagery";
            tbl[0, 7].TextFrame.Paragraphs[0].Portions[0].PortionFormat.FontHeight = 10;
            tbl[0, 13].TextFrame.Text = "Preference";
            tbl[0, 13].TextFrame.Paragraphs[0].Portions[0].PortionFormat.FontHeight = 10;


            //Values & Font Size Settings

            string[] BrandNameVal = { "Cocacola", "Pepsi", "Dr.Peper", "Mtn Dew" };
            int loopcount = 0;
            for (int a = 2; a <= 13; a++)
            {
                tbl[a, 1].TextFrame.Text = BrandNameVal[loopcount];
                tbl[a, 1].FillFormat.FillType = FillType.Solid;
                tbl[a, 1].FillFormat.SolidFillColor.Color = ColorTranslator.FromHtml(Colorcode[loopcount]);
                tbl[a, 1].TextFrame.Paragraphs[0].Portions[0].PortionFormat.FontHeight = 10;



                loopcount++;
                if (loopcount == 4)
                    loopcount = 0;
                else
                {
                    tbl[a, 1].BorderRight.DashStyle = LineDashStyle.Dot;
                    tbl[a, 1].BorderRight.FillFormat.SolidFillColor.Color = Color.Blue;
                    tbl[a, 1].BorderRight.Width = 1;

                }
            }

            //Add text to the Categories
            for (int c = 2; c <= 13; c++)
            {
                tbl[1, c].TextFrame.Text = "Cat";
                tbl[1, c].FillFormat.SolidFillColor.Color = Color.Red;
                tbl[1, c].BorderBottom.DashStyle = LineDashStyle.Dot;
                tbl[1, c].BorderBottom.FillFormat.SolidFillColor.Color = Color.Blue;
                tbl[1, c].BorderBottom.Width = 1;
                tbl[1, c].TextFrame.Paragraphs[0].Portions[0].PortionFormat.FontHeight = 10;

                //tf.Paragraphs[0].Portions[0].PortionFormat.FontHeight = 10;
                //tf.Paragraphs[0].ParagraphFormat.Bullet.Type = BulletType.None;
            }

            //Add Values to the Inner Cell
            int Val = 0;
            int columnloop = 0;
            int Rowloop = 0;
            for (int RowNumber = 2; RowNumber <= 13; RowNumber++)
            {
                Rowloop++;
                for (int columnnum = 2; columnnum <= 13; columnnum++)
                {
                    columnloop++;
                    tbl[columnnum, RowNumber].TextFrame.Text = Convert.ToString(Val++);
                    tbl[columnnum, RowNumber].TextFrame.Paragraphs[0].Portions[0].PortionFormat.FontHeight = 10;
                    tbl[columnnum, RowNumber].TextFrame.Paragraphs[0].Portions[0].PortionFormat.FillFormat.FillType = FillType.Solid;
                    tbl[columnnum, RowNumber].TextFrame.Paragraphs[0].Portions[0].PortionFormat.FillFormat.SolidFillColor.Color = Val <= 30 ? Color.Green : Val <= 60 ? Color.Blue : Color.Red;

                    if (columnloop < 4)
                    {
                        tbl[columnnum, RowNumber].BorderRight.DashStyle = LineDashStyle.Dot;
                        tbl[columnnum, RowNumber].BorderRight.FillFormat.SolidFillColor.Color = Color.Blue;
                        tbl[columnnum, RowNumber].BorderRight.Width = 1;
                    }
                    else
                        columnloop = 0;

                    if (Rowloop > 1 && Rowloop < 5)
                    {
                        tbl[columnnum, RowNumber].BorderBottom.DashStyle = LineDashStyle.Dot;
                        tbl[columnnum, RowNumber].BorderBottom.FillFormat.SolidFillColor.Color = Color.Blue;
                        tbl[columnnum, RowNumber].BorderBottom.Width = 1;
                    }
                }
                if (Rowloop == 5)
                    Rowloop = -1;
            }
            #endregion



            //var val = new Random();
            //pres.Save(@"D:\Aspose PPT\TableOnly\AsposeTableTest" + val.Next(1, 99).ToString() + ".pptx", Aspose.Slides.Export.SaveFormat.Pptx);

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