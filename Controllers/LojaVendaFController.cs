using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;

namespace RelatoriosRosset.Controllers
{
    public class LojaVendaFController : Controller
    {
        private readonly ApplicationDbContext _context;

        public LojaVendaFController(ApplicationDbContext context)
        {
            _context = context;
        }

        //GET: LojaVendas/Search
        public async Task<IActionResult> LojaVendaF(DateTime? dataInicio, DateTime? dataFim)
        {
            var query = _context.V_VENDAS_FRANQUIAS.AsQueryable();

            if (dataInicio.HasValue)
            {
                query = query.Where(v => v.DATA_VENDA >= dataInicio.Value);
            }

            if (dataFim.HasValue)
            {
                query = query.Where(v => v.DATA_VENDA <= dataFim.Value);
            }

            var lojaVendas = await query.OrderByDescending(v => v.DATA_VENDA).Take(10).ToListAsync();
            ViewBag.DataInicio = dataInicio;
            ViewBag.DataFim = dataFim;

            return View(lojaVendas);
        }
        // GET: LojaVendas/ExportarExcel
        public async Task<IActionResult> ExportarExcel(DateTime? dataInicio, DateTime? dataFim)
        {
            try
            {
                var query = _context.V_VENDAS_PROPRIAS.AsQueryable();

                if (dataInicio.HasValue)
                {
                    query = query.Where(v => v.DATA_VENDA >= dataInicio.Value);
                }

                if (dataFim.HasValue)
                {
                    query = query.Where(v => v.DATA_VENDA <= dataFim.Value);
                }

                var lojaVendas = await query.OrderBy(v => v.DATA_VENDA).ToListAsync();

                if (!lojaVendas.Any())
                {
                    return RedirectToAction(nameof(LojaVendaF), new { dataInicio, dataFim });
                }

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Vendas");

                    // Adicionar cabeçalhos
                    worksheet.Cell(1, 1).Value = "Data da Venda";
                    worksheet.Cell(1, 2).Value = "Filial";
                    worksheet.Cell(1, 3).Value = "Valor";

                    // Adicionar dados
                    for (int i = 0; i < lojaVendas.Count; i++)
                    {
                        worksheet.Cell(i + 2, 1).Value = lojaVendas[i].DATA_VENDA.ToString("dd/MM/yyyy");
                        worksheet.Cell(i + 2, 2).Value = lojaVendas[i].FILIAL;
                        worksheet.Cell(i + 2, 3).Value = lojaVendas[i].VALOR_PAGO;
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

