using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace RelatoriosRosset.Controllers
{
    public class FiliaisAtivasFController : Controller
    {
        private readonly ApplicationDbContext _context;

        public FiliaisAtivasFController(ApplicationDbContext context)
        {
            _context = context;
        }

        [HttpGet]
        public async Task<IActionResult> FiliaisAtivasF()
        {
            var lojasAtivasP = await _context.V_FILIAIS_ATIVAS_FRANQUIAS
                .OrderBy(f => f.Filial)
                .ToListAsync();
            return View(lojasAtivasP);
        }
    }
}
