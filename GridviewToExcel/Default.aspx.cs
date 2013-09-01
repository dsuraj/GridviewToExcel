using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace GridviewToExcel
{
    public partial class Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            
        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", string.Format("attachment; filename={0}", "EmployeesData.xls"));
            Response.ContentType = "application/ms-excel";

            StringWriter stringWriter = new StringWriter();
            HtmlTextWriter htmlTextWriter = new HtmlTextWriter(stringWriter);

            GridView1.AllowPaging = false;
            GridView1.DataBind();

            //This will change the header background color
            GridView1.HeaderRow.Style.Add("background-color", "#FFFFFF");

            //This will apply style to gridview header cells
            for (int index = 0; index < GridView1.HeaderRow.Cells.Count; index++)
            {
                GridView1.HeaderRow.Cells[index].Style.Add("background-color", "#d17250");
            }

            int index2 = 1;
            //This will apply style to alternate rows
            foreach (GridViewRow gridViewRow in GridView1.Rows)
            {
                gridViewRow.BackColor = Color.White;
                if (index2 <= GridView1.Rows.Count)
                {
                    if (index2 % 2 != 0)
                    {
                        for (int index3 = 0; index3 < gridViewRow.Cells.Count; index3++)
                        {
                            gridViewRow.Cells[index3].Style.Add("background-color", "#eed0bb");
                        }
                    }
                }
                index2++;
            }

            GridView1.RenderControl(htmlTextWriter);

            Response.Write(stringWriter.ToString());
            Response.End();
        }

        public override void VerifyRenderingInServerForm(Control control)
        {
            /* Verifies that the control is rendered */
        }
    }
}