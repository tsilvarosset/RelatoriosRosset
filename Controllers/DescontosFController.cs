using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace RelatoriosRosset.Controllers
{
    public class DescontosFController : Controller
    {
        private readonly ApplicationDbContext _context;

        public DescontosFController(ApplicationDbContext context)
        {
            _context = context;
        }

        [HttpGet]
        public async Task<IActionResult> DescontosF()
        {
            var DescontosP = await _context.V_DESCONTOS_FRANQUIAS
                .OrderBy(f => f.FILIAL)
                .ToListAsync();
            return View(DescontosP);
        }
    }
}
