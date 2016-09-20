using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Drawing;
using Aspose.Slides;
using System.Data;

namespace AsposeTest.UI
{
    public partial class TableRowClone : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            //Add DataTable Values

            DataTable dt = new DataTable();
          


            //Creating empty presentation
            Presentation pres = new Presentation();

            //Access first slide
            ISlide sld = pres.Slides[0];

            //Define columns with widths and rows with heights
            double[] dblCols = { 50, 50 };
            double[] dblRows = { 50, 30};

            //Add table shape to slide
            ITable tbl = sld.Shapes.AddTable(50, 50, dblCols, dblRows);

            //Set border format for each cell
            foreach (IRow row in tbl.Rows)
                foreach (ICell cell in row)
                {
                    cell.BorderTop.FillFormat.FillType = FillType.Solid;
                    cell.BorderTop.FillFormat.SolidFillColor.Color = Color.Red;
                    cell.BorderTop.Width = 5;

                    cell.BorderBottom.FillFormat.FillType = FillType.Solid;
                    cell.BorderBottom.FillFormat.SolidFillColor.Color = Color.Red;
                    cell.BorderBottom.Width = 5;

                    cell.BorderLeft.FillFormat.FillType = FillType.Solid;
                    cell.BorderLeft.FillFormat.SolidFillColor.Color = Color.Red;
                    cell.BorderLeft.Width = 5;

                    cell.BorderRight.FillFormat.FillType = FillType.Solid;
                    cell.BorderRight.FillFormat.SolidFillColor.Color = Color.Red;
                    cell.BorderRight.Width = 5;
                }


            tbl[0, 0].TextFrame.Text = "00";
            tbl[0, 1].TextFrame.Text = "01";
            tbl[1, 0].TextFrame.Text = "10";
            //tbl[2, 0].TextFrame.Text = "20";

            //AddClone adds a row in the end of the table
            tbl.Rows.AddClone(tbl.Rows[0], false);

            //InsertClone adds a row at specific position in a table
            tbl.Rows.InsertClone(2, tbl.Rows[0], false);

            //AddClone adds a column in the end of the table
            tbl.Columns.AddClone(tbl.Columns[0], false);

            //InsertClone adds a column at specific position in a table
            tbl.Columns.InsertClone(2, tbl.Columns[0], false);

            var val = new Random();
            pres.Save(@"D:\Aspose PPT\RowClone\RowCloneTest" + val.Next(1, 99).ToString() + ".pptx", Aspose.Slides.Export.SaveFormat.Pptx);
        }
    }
}