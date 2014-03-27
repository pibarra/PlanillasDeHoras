using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WebApplication2.Infrastructure;

namespace WebApplication2.Controllers
{
    public class PlantillaController : Controller
    {
        //
        // GET: /Plantilla/
        public ActionResult Index()
        {
            var workbook = WookBook.GenerarWookBook();
            return new ExcelResult(workbook, "Prueba");
        }
	}
}