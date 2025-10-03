using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace RelatoriosRosset.Controllers
{
    public class DescontosPController : Controller
    {
        private readonly ApplicationDbContext _context;

        public DescontosPController(ApplicationDbContext context)
        {
            _context = context;
        }

        [HttpGet]
        public async Task<IActionResult> DescontosP()
        {
            var DescontosP = await _context.V_DESCONTOS_PROPRIAS
                .OrderBy(f => f.FILIAL)
                .ToListAsync();
            return View(DescontosP);
        }
    }
}
