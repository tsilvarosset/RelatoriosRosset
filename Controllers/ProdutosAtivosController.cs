using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using RelatoriosRosset.Models;

namespace RelatoriosRosset.Controllers
{
    public class ProdutosAtivosController : Controller
    {
        private readonly ApplicationDbContext _context;

        public ProdutosAtivosController(ApplicationDbContext context)
        {
            _context = context;
        }


        public async Task<IActionResult> ProdutosAtivos()
        {
            var produtosAtivos = await _context.V_PRODUTOS_ATIVOS
                .OrderBy(f => f.PRODUTO)
                .Take(10)
                .ToListAsync();

            return View(produtosAtivos);
        }
        // GET: LojaVendas/ExportarExcel

        public async Task<IActionResult> ExportarExcel()
        {
            try
            {
                var query = _context.V_PRODUTOS_ATIVOS.AsQueryable();

                var produto = await query.OrderBy(v => v.PRODUTO).ToListAsync();

                if (!produto.Any())
                {
                    return RedirectToAction(nameof(ProdutosAtivos));
                }

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Produtos");

                    // Adicionar cabeçalhos
                    worksheet.Cell(1, 1).Value = "PRODUTO";
                    worksheet.Cell(1, 2).Value = "DESC PRODUTO";
                    worksheet.Cell(1, 3).Value = "COR PRODUTO";
                    worksheet.Cell(1, 4).Value = "CODIGO BARRA";
                    worksheet.Cell(1, 5).Value = "GRADE";

                    // Adicionar dados
                    for (int i = 0; i < produto.Count; i++)
                    {
                        worksheet.Cell(i + 2, 1).Value = produto[i].PRODUTO;
                        worksheet.Cell(i + 2, 2).Value = produto[i].DESC_PRODUTO;
                        worksheet.Cell(i + 2, 3).Value = produto[i].COR_PRODUTO;
                        worksheet.Cell(i + 2, 4).Value = produto[i].CODIGO_BARRA;
                       worksheet.Cell(i + 2, 5).Value = produto[i].GRADE;
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

                        string fileName = $"Relatorio_Produtos_{DateTime.Now:yyyyMMdd}.xlsx";
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
