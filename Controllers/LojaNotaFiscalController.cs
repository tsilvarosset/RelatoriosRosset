using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System;
using System.Threading.Tasks;

namespace RelatoriosRosset.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class LojaNotaFiscalController : ControllerBase
    {
        private readonly ApplicationDbContext _context;

        public LojaNotaFiscalController(ApplicationDbContext context)
        {
            _context = context;
        }

        [HttpGet]
        public async Task<IActionResult> GetLojaNotaFiscal([FromQuery] DateTime? emissaoInicio, [FromQuery] DateTime? emissaoFim)
        {
            try
            {
                var query = _context.LOJA_NOTA_FISCAL.AsQueryable();

                if (emissaoInicio.HasValue)
                {
                    var dataInicio = emissaoInicio.Value.Date;
                    query = query.Where(n => n.EMISSAO >= dataInicio);
                }

                if (emissaoFim.HasValue)
                {
                    var dataFim = emissaoFim.Value.Date.AddDays(1); // Até o final do dia
                    query = query.Where(n => n.EMISSAO < dataFim);
                }

                var notasFiscais = await query
                    .ToListAsync();

                if (!notasFiscais.Any())
                {
                    return NotFound("Nenhuma nota fiscal encontrada para o intervalo de datas fornecido.");
                }

                return Ok(notasFiscais);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Erro interno: {ex.Message}\nStackTrace: {ex.StackTrace}\nInnerException: {ex.InnerException?.Message}");
            }
        }
    }
}