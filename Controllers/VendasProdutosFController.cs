using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using RelatoriosRosset.Models;

namespace RelatoriosRosset.Controllers
{
    public class VendasProdutosFController : Controller
    {
        private readonly ApplicationDbContext _context;

        public VendasProdutosFController(ApplicationDbContext context)
        {
            _context = context;
        }


        public async Task<IActionResult> VendasProdutosF(DateTime? dataInicio, DateTime? dataFim)
        {
            if (!dataInicio.HasValue && !dataFim.HasValue)
            {
                ViewBag.DataInicio = null;
                ViewBag.DataFim = null;
                return View(new List<VendasProdutosFModel>()); // Return empty list to avoid null Model
            }

            var query = _context.V_VENDAS_PRODUTOS_F.AsQueryable();

            if (dataInicio.HasValue)
            {
                query = query.Where(v => v.DATA_VENDA >= dataInicio.Value);
            }

            if (dataFim.HasValue)
            {
                query = query.Where(v => v.DATA_VENDA <= dataFim.Value);
            }

            var lojaVendas = await query.OrderByDescending(v => v.DATA_VENDA)
                .ThenByDescending(V => V.FILIAL)
                .Take(10)
                .ToListAsync();
            ViewBag.DataInicio = dataInicio;
            ViewBag.DataFim = dataFim;

            return View(lojaVendas);
        }
        // GET: LojaVendas/ExportarExcel

        public async Task<IActionResult> ExportarExcel(DateTime? dataInicio, DateTime? dataFim)
        {
            try
            {
                var query = _context.V_VENDAS_PRODUTOS_F.AsQueryable();

                if (dataInicio.HasValue)
                {
                    query = query.Where(v => v.DATA_VENDA >= dataInicio.Value);
                }

                if (dataFim.HasValue)
                {
                    query = query.Where(v => v.DATA_VENDA <= dataFim.Value);
                }

                var lojaVendas = await query.OrderBy(v => v.FILIAL).ToListAsync();

                if (!lojaVendas.Any())
                {
                    return RedirectToAction(nameof(VendasProdutosF), new { dataInicio, dataFim });
                }

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Vendas");

                    // Adicionar cabeçalhos
                    worksheet.Cell(1, 1).Value = "CODIGO FILIAL";
                    worksheet.Cell(1, 2).Value = "FILIAL";
                    worksheet.Cell(1, 3).Value = "DATA VENDA";
                    worksheet.Cell(1, 4).Value = "PRODUTO";
                    worksheet.Cell(1, 5).Value = "COR PRODUTO";
                    worksheet.Cell(1, 6).Value = "TAMANHO";
                    worksheet.Cell(1, 7).Value = "CODIGO BARRA";
                    worksheet.Cell(1, 8).Value = "QTDE VENDIDA";
                    worksheet.Cell(1, 9).Value = "PRECO LIQUIDO";
                    worksheet.Cell(1, 10).Value = "QTDE TICKETS";

                    // Adicionar dados
                    for (int i = 0; i < lojaVendas.Count; i++)
                    {
                        worksheet.Cell(i + 2, 1).Value = lojaVendas[i].CODIGO_FILIAL;
                        worksheet.Cell(i + 2, 2).Value = lojaVendas[i].FILIAL;
                        worksheet.Cell(i + 2, 3).Value = lojaVendas[i].DATA_VENDA.ToString("dd/MM/yyyy");
                        worksheet.Cell(i + 2, 4).Value = lojaVendas[i].PRODUTO;
                        worksheet.Cell(i + 2, 5).Value = lojaVendas[i].COR_PRODUTO;
                        worksheet.Cell(i + 2, 6).Value = lojaVendas[i].TAMANHO;
                        worksheet.Cell(i + 2, 7).Value = lojaVendas[i].CODIGO_BARRA;
                        worksheet.Cell(i + 2, 8).Value = lojaVendas[i].QTDE_VENDIDA;
                        worksheet.Cell(i + 2, 9).Value = lojaVendas[i].PRECO_LIQUIDO;
                        worksheet.Cell(i + 2, 10).Value = lojaVendas[i].QTDE_TICKETS;
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
