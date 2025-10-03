using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using RelatoriosRosset.Models;

namespace RelatoriosRosset.Controllers
{
    public class DevEntradasController : Controller
    {
        private readonly ApplicationDbContext _context;
        public DevEntradasController(ApplicationDbContext context)
        {
            _context = context;
        }
        
        public async Task<IActionResult> DevEntradas(DateTime? dataInicio, DateTime? dataFim)
        {
            try
            {
                var query = _context.V_DEVOLUCOES_ORIGENS.AsQueryable();

                if (dataInicio.HasValue)
                    query = query.Where(V => V.EMISSAO >= dataInicio.Value);

                if (dataFim.HasValue)
                    query = query.Where(V => V.EMISSAO <= dataFim.Value);

                var notasP = await query
                    .OrderByDescending(V => V.EMISSAO)
                    .Take(10)
                    .Select(V => new DevEntradasModel
                    {
                        CODIGO_FILIAL = V.CODIGO_FILIAL,
                        FILIAL = V.FILIAL,
                        EMISSAO = V.EMISSAO,
                        NF_DEVOLUCAO = V.NF_DEVOLUCAO,
                        ENTRADA_ORIGEM = V.ENTRADA_ORIGEM,
                        CHAVE_DEVOLUCAO = V.CHAVE_DEVOLUCAO,
                        CHAVE_ENTRADA = V.CHAVE_ENTRADA
                    })
                    .ToListAsync();

                ViewBag.DataInicio = dataInicio;
                ViewBag.DataFim = dataFim;

                return View(notasP);
            }
            catch (Exception ex)
            {
                // Retorne a mensagem de erro para depuração
                return Content($"Erro ao executar a consulta: {ex.Message}\nStack Trace: {ex.StackTrace}");
            }
        }
        public async Task<IActionResult> ExportarExcel(DateTime? dataInicio, DateTime? dataFim)
        {
            try
            {
                var query = _context.V_DEVOLUCOES_ORIGENS.AsQueryable();

                if (dataInicio.HasValue)
                    query = query.Where(V => V.EMISSAO >= dataInicio.Value);

                if (dataFim.HasValue)
                    query = query.Where(V => V.EMISSAO <= dataFim.Value);

                var notas = query
                    .OrderBy(V => V.FILIAL)
                    .Select(V => new DevEntradasModel
                    {
                        CODIGO_FILIAL = V.CODIGO_FILIAL,
                        FILIAL = V.FILIAL,
                        EMISSAO = V.EMISSAO,
                        NF_DEVOLUCAO = V.NF_DEVOLUCAO,
                        ENTRADA_ORIGEM = V.ENTRADA_ORIGEM,
                        CHAVE_DEVOLUCAO = V.CHAVE_DEVOLUCAO,
                        CHAVE_ENTRADA = V.CHAVE_ENTRADA
                    })
                    .ToList();

                if (!notas.Any())
                {
                    return RedirectToAction(nameof(DevEntradas), new { dataInicio, dataFim });
                }

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Notas");

                    // Cabeçalhos
                    worksheet.Cell(1, 1).Value = "CODIGO_FILIAL";
                    worksheet.Cell(1, 2).Value = "FILIAL";
                    worksheet.Cell(1, 3).Value = "EMISSAO";
                    worksheet.Cell(1, 4).Value = "NF_DEVOLUCAO";
                    worksheet.Cell(1, 5).Value = "ENTRADA_ORIGEM";
                    worksheet.Cell(1, 6).Value = "CHAVE_DEVOLUCAO";
                    worksheet.Cell(1, 7).Value = "CHAVE_ENTRADA";

                    // Dados
                    for (int i = 0; i < notas.Count; i++)
                    {
                        worksheet.Cell(i + 2, 1).Value = notas[i].CODIGO_FILIAL;
                        worksheet.Cell(i + 2, 2).Value = notas[i].FILIAL;
                        worksheet.Cell(i + 2, 3).Value = notas[i].EMISSAO;
                        worksheet.Cell(i + 2, 4).Value = notas[i].NF_DEVOLUCAO;
                        worksheet.Cell(i + 2, 5).Value = notas[i].ENTRADA_ORIGEM;
                        worksheet.Cell(i + 2, 6).Value = notas[i].CHAVE_DEVOLUCAO;
                        worksheet.Cell(i + 2, 7).Value = notas[i].CHAVE_ENTRADA;

                    }

                    // Ajustar formato
                    worksheet.Columns().AdjustToContents();
                    worksheet.Row(1).Style.Font.Bold = true;

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        stream.Position = 0;
                        string fileName = $"Relatorio_Notas_{DateTime.Now:yyyyMMdd}.xlsx";
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
