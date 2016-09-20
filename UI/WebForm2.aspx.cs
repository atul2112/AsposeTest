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
    public partial class WebForm2 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            GetPPT();
        }

        private void GetPPT()
        {

            //Instantiate Presentation class that represents PPTX file
            Presentation pres = new Presentation();

            //Access first slide
            ISlide sld = pres.Slides[0];
            #region Table
            //ISlide sld = pres.Slides[0];

            //Define columns with widths and rows with heights
            double[] dblCols = { 50, 50, 50, 50 };
            double[] dblRows = { 50, 50, 50, 50 };

            //Add table shape to slide
            ITable tbl = sld.Shapes.AddTable(110, 110, dblCols, dblRows);

            int RowNum = 0;

            //Set border format for each cell
            foreach (IRow row in tbl.Rows)
            {
                foreach (ICell cell in row)
                {
                    cell.BorderTop.FillFormat.FillType = FillType.Solid;
                    cell.BorderTop.FillFormat.SolidFillColor.Color = Color.Blue;
                    cell.BorderTop.Width = 2;

                    cell.BorderBottom.FillFormat.FillType = FillType.Solid;
                    cell.BorderBottom.FillFormat.SolidFillColor.Color = Color.Blue;
                    cell.BorderBottom.Width = 2;

                    cell.BorderLeft.FillFormat.FillType = FillType.Solid;
                    cell.BorderLeft.FillFormat.SolidFillColor.Color = Color.Blue;
                    cell.BorderLeft.Width = 2;

                    cell.BorderRight.FillFormat.FillType = FillType.Solid;
                    cell.BorderRight.FillFormat.SolidFillColor.Color = Color.Blue;
                    cell.BorderRight.Width = 2;

                }

            }

            tbl.MergeCells(tbl[0, 0], tbl[1, 0], false);
            tbl.MergeCells(tbl[2, 0], tbl[3, 0], false);

            tbl.MergeCells(tbl[0, 3], tbl[1, 3], false);

            ////Merge cells 1 & 2 of row 1
            //tbl.MergeCells(tbl[0, 0], tbl[1, 0], false);

            ////Add text to the merged cell
            //tbl[0, 0].TextFrame.Text = "Merged Cells";
            #endregion


            var val = new Random();
            pres.Save(@"D:\Aspose PPT\TableOnly\AsposeTableTest" + val.Next(1, 99).ToString() + ".pptx", Aspose.Slides.Export.SaveFormat.Pptx);
        }



    }
}