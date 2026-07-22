using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using RelatoriosRosset.Models;

namespace RelatoriosRosset.Controllers
{
    public class PedidosGradeController : Controller
    {
        private readonly ApplicationDbContext _context;

        public PedidosGradeController(ApplicationDbContext context)
        {
            _context = context;
        }


        public async Task<IActionResult> PedidosGrade(DateTime? dataInicio, DateTime? dataFim)
        {
            if (!dataInicio.HasValue && !dataFim.HasValue)
            {
                ViewBag.DataInicio = null;
                ViewBag.DataFim = null;
                return View(new List<PedidosGradeModel>()); // Return empty list to avoid null Model
            }

            var query = _context.PEDIDOS_GRADE.AsQueryable();

            if (dataInicio.HasValue)
            {
                query = query.Where(v => v.EMISSAO >= dataInicio.Value);
            }

            if (dataFim.HasValue)
            {
                query = query.Where(v => v.EMISSAO <= dataFim.Value);
            }

            var pedido = await query.OrderByDescending(v => v.EMISSAO)
                .ThenByDescending(V => V.PEDIDO)
                .Take(10)
                .ToListAsync();
            ViewBag.DataInicio = dataInicio;
            ViewBag.DataFim = dataFim;

            return View(pedido);
        }
        // GET: LojaVendas/ExportarExcel

        public async Task<IActionResult> ExportarExcel(DateTime? dataInicio, DateTime? dataFim)
        {
            try
            {
                var query = _context.PEDIDOS_GRADE.AsQueryable();

                if (dataInicio.HasValue)
                {
                    query = query.Where(v => v.EMISSAO >= dataInicio.Value);
                }

                if (dataFim.HasValue)
                {
                    query = query.Where(v => v.EMISSAO <= dataFim.Value);
                }

                var pedido = await query.OrderBy(v => v.PEDIDO).ToListAsync();

                if (!pedido.Any())
                {
                    return RedirectToAction(nameof(PedidosGrade), new { dataInicio, dataFim });
                }

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Vendas");

                    // Adicionar cabeçalhos
                    worksheet.Cell(1, 1).Value = "GRADE";
                    worksheet.Cell(1, 2).Value = "CLIENTE ATACADO";
                    worksheet.Cell(1, 3).Value = "PEDIDO";
                    worksheet.Cell(1, 4).Value = "PRODUTO";
                    worksheet.Cell(1, 5).Value = "COR PRODUTO";
                    worksheet.Cell(1, 6).Value = "TAMANHO";
                    worksheet.Cell(1, 7).Value = "QTDE";
                    worksheet.Cell(1, 8).Value = "VALOR UNITARIO";
                    worksheet.Cell(1, 9).Value = "VALOR PEDIDO";
                    worksheet.Cell(1, 10).Value = "EMISSAO";

                    // Adicionar dados
                    for (int i = 0; i < pedido.Count; i++)
                    {
                        worksheet.Cell(i + 2, 1).Value = pedido[i].GRADE;
                        worksheet.Cell(i + 2, 2).Value = pedido[i].CLIENTE_ATACADO;
                        worksheet.Cell(i + 2, 3).Value = pedido[i].PEDIDO;
                        worksheet.Cell(i + 2, 4).Value = pedido[i].PRODUTO;
                        worksheet.Cell(i + 2, 5).Value = pedido[i].COR_PRODUTO;
                        worksheet.Cell(i + 2, 6).Value = pedido[i].TAMANHO;
                        worksheet.Cell(i + 2, 7).Value = pedido[i].QTDE;
                        worksheet.Cell(i + 2, 8).Value = pedido[i].VALOR_UNITARIO;
                        worksheet.Cell(i + 2, 9).Value = pedido[i].VALOR_PEDIDO;
                        worksheet.Cell(i + 2, 10).Value = pedido[i].EMISSAO.ToString("dd/MM/yyyy");
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

                        string fileName = $"Relatorio_Pedidos_{DateTime.Now:yyyyMMdd}.xlsx";
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
