using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace RelatoriosRosset.Controllers
{
    public class FiliaisAtivasPController : Controller
    {
        private readonly ApplicationDbContext _context;

        public FiliaisAtivasPController(ApplicationDbContext context)
        {
            _context = context;
        }

        [HttpGet]
        public async Task<IActionResult> FiliaisAtivasP()
        {
            var lojasAtivasP = await _context.V_FILIAIS_ATIVAS_PROPRIAS
                .OrderBy(f => f.Filial)
                .ToListAsync();
            return View(lojasAtivasP);
        }
        
    }
}
