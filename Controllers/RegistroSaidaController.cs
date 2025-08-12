using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;

namespace RelatoriosRosset.Controllers
{
    [ApiController]
    [Route("[controller]")]
    
    public class RegistroSaidaController : Controller
    {
        private readonly ApplicationDbContext _context;
        private readonly ILogger<RegistroSaidaController> _logger;

        public RegistroSaidaController(ApplicationDbContext context, ILogger<RegistroSaidaController> logger)
        {
            _context = context;
            _logger = logger;
        }

        [HttpGet]
        public async Task<IActionResult> GetRegistroSaidas([FromQuery] DateTime? emissaoInicio, [FromQuery] DateTime? emissaoFim)
        {
            try
            {
                _logger.LogInformation("Starting GetRegistroSaidas with emissaoInicio: {0}, emissaoFim: {1}", emissaoInicio, emissaoFim);
                var query = _context.W_LF_REGISTRO_SAIDA_IMPOSTO_ITEM.AsQueryable();

                if (emissaoInicio.HasValue)
                {
                    var dataInicio = emissaoInicio.Value.Date;
                    query = query.Where(n => n.EMISSAO >= dataInicio);
                    _logger.LogInformation("Applied emissaoInicio filter: {0}", dataInicio);
                }

                if (emissaoFim.HasValue)
                {
                    var dataFim = emissaoFim.Value.Date.AddDays(1);
                    query = query.Where(n => n.EMISSAO < dataFim);
                    _logger.LogInformation("Applied emissaoFim filter: {0}", dataFim);
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
                _logger.LogError(ex, "Error in GetRegistroSaidas");
                return StatusCode(500, $"Erro interno: {ex.Message}");
            }
        }
    }
}

