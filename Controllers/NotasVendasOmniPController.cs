using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using RelatoriosRosset.Models;

namespace RelatoriosRosset.Controllers
{
    public class NotasVendasOmniPController : Controller
    {
        private readonly ApplicationDbContext _context;

        public NotasVendasOmniPController(ApplicationDbContext context)
        {
            _context = context;
        }
        public async Task<IActionResult> NotasVendasOmniP(DateTime? dataInicio, DateTime? dataFim, string clienteVarejo, string cpfCgc)
        {
            // Verifica se nenhum filtro foi fornecido
            if (!dataInicio.HasValue && !dataFim.HasValue && string.IsNullOrEmpty(clienteVarejo) && string.IsNullOrEmpty(cpfCgc))
            {
                ViewBag.Mensagem = "Por favor, forneça pelo menos um filtro para realizar a pesquisa.";
                return View(new List<NotasVendasOmniPModel>());
            }

            var query = _context.NOTAS_VENDAS_PROPRIAS.AsQueryable();

            if (dataInicio.HasValue)
                query = query.Where(v => v.EMISSAO >= dataInicio.Value);

            if (dataFim.HasValue)
                query = query.Where(v => v.EMISSAO <= dataFim.Value);

            if (!string.IsNullOrEmpty(clienteVarejo))
                query = query.Where(v => v.CLIENTE_VAREJO.Contains(clienteVarejo));

            if (!string.IsNullOrEmpty(cpfCgc))
                query = query.Where(v => v.CPF_CGC.Contains(cpfCgc));

            var notasP = await query.OrderByDescending(v => v.EMISSAO).Take(10).ToListAsync();

            ViewBag.DataInicio = dataInicio;
            ViewBag.DataFim = dataFim;
            ViewBag.ClienteVarejo = clienteVarejo;
            ViewBag.CpfCgc = cpfCgc;

            return View(notasP);
        }

        public async Task<IActionResult> ExportarExcel(DateTime? dataInicio, DateTime? dataFim, string clienteVarejo, string cpfCgc)
        {
            try
            {
                var query = _context.NOTAS_VENDAS_PROPRIAS.AsQueryable();

                if (dataInicio.HasValue)
                {
                    query = query.Where(v => v.EMISSAO >= dataInicio.Value);
                }

                if (dataFim.HasValue)
                {
                    query = query.Where(v => v.EMISSAO <= dataFim.Value);
                }

                if (!string.IsNullOrEmpty(clienteVarejo))
                {
                    query = query.Where(v => v.CLIENTE_VAREJO.Contains(clienteVarejo));
                }

                if (!string.IsNullOrEmpty(cpfCgc))
                {
                    query = query.Where(v => v.CPF_CGC.Contains(cpfCgc));
                }

                var notasF = await query.OrderBy(v => v.EMISSAO).ToListAsync();

                if (!notasF.Any())
                {
                    return RedirectToAction(nameof(NotasVendasOmniP), new { dataInicio, dataFim, clienteVarejo, cpfCgc });
                }

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Notas");

                    // Adicionar cabeçalhos
                    worksheet.Cell(1, 1).Value = "FILIAL";
                    worksheet.Cell(1, 2).Value = "NUMERO NF";
                    worksheet.Cell(1, 3).Value = "EMISSAO";
                    worksheet.Cell(1, 4).Value = "VALOR TOTAL";
                    worksheet.Cell(1, 5).Value = "NATUREZA OPERACAO";
                    worksheet.Cell(1, 6).Value = "CPF/CNPJ";
                    worksheet.Cell(1, 7).Value = "CLIENTE VAREJO";
                    worksheet.Cell(1, 8).Value = "CHAVE NFE";

                    // Adicionar dados
                    for (int i = 0; i < notasF.Count; i++)
                    {
                        worksheet.Cell(i + 2, 1).Value = notasF[i].FILIAL;
                        worksheet.Cell(i + 2, 2).Value = notasF[i].NF_NUMERO;
                        worksheet.Cell(i + 2, 3).Value = notasF[i].EMISSAO.ToString("dd/MM/yyyy");
                        worksheet.Cell(i + 2, 4).Value = notasF[i].VALOR_TOTAL;
                        worksheet.Cell(i + 2, 5).Value = notasF[i].NATUREZA_OPERACAO_CODIGO;
                        worksheet.Cell(i + 2, 6).Value = notasF[i].CPF_CGC;
                        worksheet.Cell(i + 2, 7).Value = notasF[i].CLIENTE_VAREJO;
                        worksheet.Cell(i + 2, 8).Value = notasF[i].CHAVE_NFE;
                    }

                    // Ajustar formato das colunas
                    worksheet.Columns(1, 8).AdjustToContents();

                    // Configurar cabeçalhos como negrito
                    worksheet.Row(1).Style.Font.Bold = true;

                    // Converter o workbook para um array de bytes
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
