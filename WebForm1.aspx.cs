
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Xceed.Words.NET;

namespace Probar
{
    public partial class WebForm1 : System.Web.UI.Page
    {

        protected void Page_Load(object sender, EventArgs e)
        {
            DocX doc = DocX.Load(Server.MapPath("APROBACIÓN_ACTUALIZACIÓN_DE_CONOCIMIENTOS.docx"));
            doc.Bookmarks["CarreraCoordinador"].SetText("Holabebe");
            doc.SaveAs(Server.MapPath("Auxiliar.docx"));

            doc = DocX.Load(Server.MapPath("APROBACIÓN_ACTUALIZACIÓN_DE_CONOCIMIENTOS.docx"));
            doc.Bookmarks["CarreraCoordinador"].SetText("Dieguin");
            //t1.Text = doc.Bookmarks["Cuerpo1"].Paragraph.Text;
            doc.SaveAs(Server.MapPath("Auxiliar2.docx"));






            using (FileStream fileStream = File.OpenRead(Server.MapPath("Auxiliar.docx")))
            {
                MemoryStream memStream = new MemoryStream();
                memStream.SetLength(fileStream.Length);
                fileStream.Read(memStream.GetBuffer(), 0, (int)fileStream.Length);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                Response.AddHeader("Content-Disposition", "attachment; filename=dieguin.docx");
                Response.BinaryWrite(memStream.ToArray());
                Response.Flush();
                Response.Write("<script type='text/javascript'> setTimeout('location.reload(true); ', 0);</script>");
                Response.Close();
                Response.End();
            }
        }
       
    }
}