using Microsoft.AspNetCore.Mvc;

namespace RelatoriosRosset.Controllers
{
    public class FiscalController : Controller
    {
        public IActionResult Fiscal()
        {
            return View();
        }
    }
}
