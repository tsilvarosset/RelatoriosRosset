using Microsoft.AspNetCore.Mvc;

namespace RelatoriosRosset.Controllers
{
    public class PaginaemDesenvolvimentoController : Controller
    {
        private readonly ILogger<PaginaemDesenvolvimentoController> _logger;

        public PaginaemDesenvolvimentoController(ILogger<PaginaemDesenvolvimentoController> logger)
        {
            _logger = logger;
        }
        public IActionResult PaginaemDesenvolvimento()
        {
            return View();
        }
    }
}
