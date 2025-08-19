using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace RelatoriosRosset.Controllers
{
    public class TicketsNotasController : Controller
    {
        private readonly ApplicationDbContext _context;

        public TicketsNotasController(ApplicationDbContext context)
        {
            _context = context;
        }

        public async Task<IActionResult> TicketsNotas(DateTime? dataInicio, DateTime? dataFim)
        {
            var query = _context.V_TICKETS_NOTAS.AsQueryable();

            if (dataInicio.HasValue)
            {
                query = query.Where(v => v.DATA >= dataInicio.Value);
            }

            if (dataFim.HasValue)
            {
                query = query.Where(v => v.DATA <= dataFim.Value);
            }

            var tickets = await query.OrderByDescending(v => v.DATA)
                .ThenByDescending(V => V.VALOR_PAGO)
                .Take(10)
                .ToListAsync();
            ViewBag.DataInicio = dataInicio;
            ViewBag.DataFim = dataFim;

            return View(tickets);
        }
        // GET: LojaVendas/ExportarExcel

        public async Task<IActionResult> ExportarExcel(DateTime? dataInicio, DateTime? dataFim)
        {
            try
            {
                var query = _context.V_TICKETS_NOTAS.AsQueryable();

                if (dataInicio.HasValue)
                {
                    query = query.Where(v => v.DATA >= dataInicio.Value);
                }

                if (dataFim.HasValue)
                {
                    query = query.Where(v => v.DATA <= dataFim.Value);
                }

                var tickets = await query.OrderBy(v => v.FILIAL).ToListAsync();

                if (!tickets.Any())
                {
                    return RedirectToAction(nameof(TicketsNotas), new { dataInicio, dataFim });
                }

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Vendas");

                    // Adicionar cabeçalhos
                    worksheet.Cell(1, 1).Value = "FILIAL";
                    worksheet.Cell(1, 2).Value = "DATA";
                    worksheet.Cell(1, 3).Value = "TICKET";
                    worksheet.Cell(1, 4).Value = "VALOR_PAGO";
                    worksheet.Cell(1, 5).Value = "NUMERO_FISCAL_VENDA";
                    worksheet.Cell(1, 6).Value = "NUMERO_FISCAL_TROCA";
                    worksheet.Cell(1, 7).Value = "NUMERO_FISCAL_CANCELAMENTO";

                    // Adicionar dados
                    for (int i = 0; i < tickets.Count; i++)
                    {
                        worksheet.Cell(i + 2, 1).Value = tickets[i].FILIAL;
                        worksheet.Cell(i + 2, 2).Value = tickets[i].DATA.ToString("dd/MM/yyyy");
                        worksheet.Cell(i + 2, 3).Value = tickets[i].TICKET;
                        worksheet.Cell(i + 2, 4).Value = tickets[i].VALOR_PAGO;
                        worksheet.Cell(i + 2, 5).Value = tickets[i].NUMERO_FISCAL_VENDA;
                        worksheet.Cell(i + 2, 6).Value = tickets[i].NUMERO_FISCAL_TROCA;
                        worksheet.Cell(i + 2, 7).Value = tickets[i].NUMERO_FISCAL_CANCELAMENTO;
                    }

                    // Ajustar formato das colunas
                    worksheet.Columns().AdjustToContents();

                    // Configurar cabeçalhos como negrito
                    worksheet.Row(1).Style.Font.Bold = true;

                    // Converter o workbook para um array de bytes
                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        stream.Position = 0;

                        string fileName = $"Relatorio_Vendas_{DateTime.Now:yyyyMMdd}.xlsx";
                        return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                    }
                }
            }
            catch (Exception ex)
            {
                return Content($"Erro ao gerar o relatório: {ex.Message}\nStack Trace: {ex.StackTrace}");
            }
        }
    }

}























