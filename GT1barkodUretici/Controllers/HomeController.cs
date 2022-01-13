using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace GT1barkodUretici.Controllers
{
    public class barkodModel
    {
        public string Baslangic { get; set; }
        public string GuvenlikNumarasi { get; set; }
        public string Barkod { get; set; }
    }

    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Excel(double baslangic, double bitis)
        {
            List<barkodModel> barkods = new List<barkodModel>();

            for (double i = baslangic; i < bitis; i++)
            {
                var factor = 3;
                double sum = 0;

                for (int index = i.ToString().Count(); index > 0; --index)
                {
                    sum = sum + Convert.ToDouble(i.ToString().Substring(index - 1, 1)) * factor;
                    factor = 4 - factor;
                }

                var cc = ((1000 - sum) % 10);
                var barkod = i.ToString() + cc.ToString();

                barkods.Add(new barkodModel
                {
                    Baslangic = i.ToString(),
                    Barkod = barkod,
                    GuvenlikNumarasi = cc.ToString()
                });
            }

            DataTable dt = new DataTable();
            dt.TableName = "Sayfa1";
            dt.Columns.Add("Ön Barkod", typeof(string));
            dt.Columns.Add("Güvenlik-Kontrol Numarası", typeof(string));
            dt.Columns.Add("Barkod", typeof(string));

            foreach (var item in barkods)
            {
                string[] row = new string[3];
                row[0] = item.Baslangic;
                row[1] = item.GuvenlikNumarasi;
                row[2] = item.Barkod; 
                dt.Rows.Add(row);
            }

            string dosyaAdi = "GTIN-TopluBarkodUretici-" + DateTime.Now.ToString().Replace(".", "-").Replace(" ", "");
            var grid = new GridView();
            grid.DataSource = dt;
            grid.DataBind();

            Response.ClearContent();
            Response.ContentEncoding = Encoding.UTF8;
            Response.BinaryWrite(Encoding.UTF8.GetPreamble());
            Response.AddHeader("content-disposition", "attachment; filename=" + dosyaAdi + ".xls");

            Response.ContentType = "application/vnd.ms-excel";
            StringWriter sw = new StringWriter();
            Encoding.GetEncoding(1254).GetBytes(sw.ToString());
            HtmlTextWriter htw = new HtmlTextWriter(sw);

            grid.RenderControl(htw);

            Response.Write(sw.ToString());
            Response.End();

            return RedirectToAction("Index");
        }

    }
}