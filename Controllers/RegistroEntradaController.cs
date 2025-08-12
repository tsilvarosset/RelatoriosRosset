using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace RelatoriosRosset.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class RegistroEntradaController : Controller
    {
        private readonly ApplicationDbContext _context;
        private readonly ILogger<RegistroEntradaController> _logger;

        public RegistroEntradaController(ApplicationDbContext context, ILogger<RegistroEntradaController> logger)
        {
            _context = context;
            _logger = logger;
        }

        [HttpGet]
        public async Task<IActionResult> GetRegistroEntradas([FromQuery] DateTime? recebimentoInicio, [FromQuery] DateTime? recebimentoFim)
        {
            try
            {
                _logger.LogInformation("Starting GetRegistroEntradas with recebimentoInicio: {0}, recebimentoFim: {1}", recebimentoInicio, recebimentoFim);
                var query = _context.W_LF_REGISTRO_ENTRADA_IMPOSTO_ITEM.AsQueryable();

                if (recebimentoInicio.HasValue)
                {
                    var dataInicio = recebimentoInicio.Value.Date;
                    query = query.Where(n => n.RECEBIMENTO >= dataInicio);
                    _logger.LogInformation("Applied recebimentoInicio filter: {0}", dataInicio);
                }

                if (recebimentoFim.HasValue)
                {
                    var dataFim = recebimentoFim.Value.Date.AddDays(1);
                    query = query.Where(n => n.RECEBIMENTO < dataFim);
                    _logger.LogInformation("Applied recebimentoFim filter: {0}", dataFim);
                }

                var notasFiscais = await query.ToListAsync();
                _logger.LogInformation("Retrieved {0} records", notasFiscais.Count);

                if (!notasFiscais.Any())
                {
                    return NotFound("Nenhuma nota fiscal encontrada para o intervalo de datas fornecido.");
                }

                return Ok(notasFiscais);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in GetRegistroEntradas");
                return StatusCode(500, $"Erro interno: {ex.Message}");
            }
        }
    }
}
