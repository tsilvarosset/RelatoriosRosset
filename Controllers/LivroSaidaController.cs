using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using RelatoriosRosset.Models;

namespace RelatoriosRosset.Controllers
{
    public class LivroSaidaController : Controller
    {
        private readonly ApplicationDbContext _context;
        public LivroSaidaController(ApplicationDbContext context)
        {
            _context = context;
        }
        public async Task<IActionResult> LivroSaida(DateTime? dataInicio, DateTime? dataFim)
        {
            var query = _context.W_LF_REGISTRO_SAIDA_IMPOSTO_ITEM.AsQueryable();

            if (dataInicio.HasValue)
                query = query.Where(v => v.EMISSAO >= dataInicio.Value);

            if (dataFim.HasValue)
                query = query.Where(v => v.EMISSAO <= dataFim.Value);

            var notasP = await query
                .OrderByDescending(v => v.EMISSAO)
                .Take(10)
                .Select(v => new LivroSaidaModel
                {
                    FILIAL = v.FILIAL,
                    NF_SAIDA = v.NF_SAIDA,
                    DESTINO_CLIENTE = v.CLIENTE_VAREJO,
                    DESTINO_FILIAL = v.NOME_CLIFOR,
                    SERIE_NF = v.SERIE_NF,
                    SERIE_NF_OFICIAL = v.SERIE_NF_OFICIAL,
                    IMPOSTO = v.IMPOSTO,
                    VALOR_CONTABIL = v.VALOR_CONTABIL,
                    BASE_IMPOSTO = v.BASE_IMPOSTO,
                    TAXA_IMPOSTO = v.TAXA_IMPOSTO,
                    VALOR_IMPOSTO = v.VALOR_IMPOSTO,
                    VALOR_IMPOSTO_OUTROS = v.VALOR_IMPOSTO_OUTROS,
                    VALOR_IMPOSTO_ISENTO = v.VALOR_IMPOSTO_ISENTO,
                    CODIGO_FISCAL_OPERACAO = v.CODIGO_FISCAL_OPERACAO,
                    DENOMINACAO_CFOP = v.DENOMINACAO_CFOP,
                    EMISSAO = v.EMISSAO ?? default(DateTime), 
                    CODIGO_ITEM = v.CODIGO_ITEM,
                    QTDE_ITEM = v.QTDE_ITEM,
                    PRECO_UNITARIO = v.PRECO_UNITARIO,
                    VALOR_BRUTO_ITEM = v.VALOR_BRUTO_ITEM,
                    DESCRICAO_ITEM = v.DESCRICAO_ITEM,
                    UNIDADE = v.UNIDADE,
                    CLASSIF_FISCAL = v.CLASSIF_FISCAL
                })
                .ToListAsync();

            ViewBag.DataInicio = dataInicio;
            ViewBag.DataFim = dataFim;

            return View(notasP);
        }

        public async Task<IActionResult> ExportarExcel(DateTime? dataInicio, DateTime? dataFim)
        {
            try
            {
                var query = _context.W_LF_REGISTRO_SAIDA_IMPOSTO_ITEM.AsQueryable();

                if (dataInicio.HasValue)
                    query = query.Where(v => v.EMISSAO >= dataInicio.Value);

                if (dataFim.HasValue)
                    query = query.Where(v => v.EMISSAO <= dataFim.Value);

                var notas = query
                    .OrderBy(v => v.FILIAL)
                    .Select(v => new LivroSaidaModel
                    {
                        FILIAL = v.FILIAL,
                        NF_SAIDA = v.NF_SAIDA,
                        DESTINO_CLIENTE = v.CLIENTE_VAREJO,
                        DESTINO_FILIAL = v.NOME_CLIFOR,
                        SERIE_NF = v.SERIE_NF,
                        SERIE_NF_OFICIAL = v.SERIE_NF_OFICIAL,
                        IMPOSTO = v.IMPOSTO,
                        VALOR_CONTABIL = v.VALOR_CONTABIL,
                        BASE_IMPOSTO = v.BASE_IMPOSTO,
                        TAXA_IMPOSTO = v.TAXA_IMPOSTO,
                        VALOR_IMPOSTO = v.VALOR_IMPOSTO,
                        VALOR_IMPOSTO_OUTROS = v.VALOR_IMPOSTO_OUTROS,
                        VALOR_IMPOSTO_ISENTO = v.VALOR_IMPOSTO_ISENTO,
                        CODIGO_FISCAL_OPERACAO = v.CODIGO_FISCAL_OPERACAO,
                        DENOMINACAO_CFOP = v.DENOMINACAO_CFOP,
                        EMISSAO = v.EMISSAO ?? default(DateTime),
                        CODIGO_ITEM = v.CODIGO_ITEM,
                        QTDE_ITEM = v.QTDE_ITEM,
                        PRECO_UNITARIO = v.PRECO_UNITARIO,
                        VALOR_BRUTO_ITEM = v.VALOR_BRUTO_ITEM,
                        DESCRICAO_ITEM = v.DESCRICAO_ITEM,
                        UNIDADE = v.UNIDADE,
                        CLASSIF_FISCAL = v.CLASSIF_FISCAL
                    })
                    .ToList();

                if (!notas.Any())
                {
                    return RedirectToAction(nameof(LivroSaida), new { dataInicio, dataFim });
                }

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Notas");

                    // Cabeçalhos
                    worksheet.Cell(1, 1).Value = "FILIAL";
                    worksheet.Cell(1, 2).Value = "NF SAIDA";
                    worksheet.Cell(1, 3).Value = "DESTINO CLIENTE";
                    worksheet.Cell(1, 4).Value = "DESTINO FILIAL";
                    worksheet.Cell(1, 5).Value = "SERIE NF";
                    worksheet.Cell(1, 6).Value = "SERIE NF OFICIAL";
                    worksheet.Cell(1, 7).Value = "IMPOSTO";
                    worksheet.Cell(1, 8).Value = "VALOR CONTABIL";
                    worksheet.Cell(1, 9).Value = "BASE IMPOSTO";
                    worksheet.Cell(1, 10).Value = "TAXA IMPOSTO";
                    worksheet.Cell(1, 11).Value = "VALOR IMPOSTO";
                    worksheet.Cell(1, 12).Value = "VALOR IMPOSTO OUTROS";
                    worksheet.Cell(1, 13).Value = "VALOR IMPOSTO ISENTO";
                    worksheet.Cell(1, 14).Value = "CODIGO FISCAL OPERACAO";
                    worksheet.Cell(1, 15).Value = "DENOMINACAO CFOP";
                    worksheet.Cell(1, 16).Value = "EMISSAO";
                    worksheet.Cell(1, 17).Value = "CODIGO ITEM";
                    worksheet.Cell(1, 18).Value = "QTDE ITEM";
                    worksheet.Cell(1, 19).Value = "PRECO UNITARIO";
                    worksheet.Cell(1, 21).Value = "VALOR BRUTO ITEM";
                    worksheet.Cell(1, 22).Value = "DESCRICAO ITEM";
                    worksheet.Cell(1, 23).Value = "UNIDADE";
                    worksheet.Cell(1, 24).Value = "CLASSIF FISCAL";

                    // Dados
                    for (int i = 0; i < notas.Count; i++)
                    {
                        worksheet.Cell(i + 2, 1).Value = notas[i].FILIAL;
                        worksheet.Cell(i + 2, 2).Value = notas[i].NF_SAIDA;
                        worksheet.Cell(i + 2, 3).Value = notas[i].DESTINO_CLIENTE;
                        worksheet.Cell(i + 2, 4).Value = notas[i].DESTINO_FILIAL;
                        worksheet.Cell(i + 2, 5).Value = notas[i].SERIE_NF;
                        worksheet.Cell(i + 2, 6).Value = notas[i].SERIE_NF_OFICIAL;
                        worksheet.Cell(i + 2, 7).Value = notas[i].IMPOSTO;
                        worksheet.Cell(i + 2, 8).Value = notas[i].VALOR_CONTABIL;
                        worksheet.Cell(i + 2, 9).Value = notas[i].BASE_IMPOSTO;
                        worksheet.Cell(i + 2, 10).Value = notas[i].TAXA_IMPOSTO;
                        worksheet.Cell(i + 2, 11).Value = notas[i].VALOR_IMPOSTO;
                        worksheet.Cell(i + 2, 12).Value = notas[i].VALOR_IMPOSTO_OUTROS;
                        worksheet.Cell(i + 2, 13).Value = notas[i].VALOR_IMPOSTO_ISENTO;
                        worksheet.Cell(i + 2, 14).Value = notas[i].CODIGO_FISCAL_OPERACAO;
                        worksheet.Cell(i + 2, 15).Value = notas[i].DENOMINACAO_CFOP;
                        worksheet.Cell(i + 2, 16).Value = notas[i].EMISSAO.ToString("dd/MM/yyyy");
                        worksheet.Cell(i + 2, 17).Value = notas[i].CODIGO_ITEM;
                        worksheet.Cell(i + 2, 18).Value = notas[i].QTDE_ITEM;
                        worksheet.Cell(i + 2, 19).Value = notas[i].PRECO_UNITARIO;
                        worksheet.Cell(i + 2, 20).Value = notas[i].VALOR_BRUTO_ITEM;
                        worksheet.Cell(i + 2, 21).Value = notas[i].DESCRICAO_ITEM;
                        worksheet.Cell(i + 2, 22).Value = notas[i].UNIDADE;
                        worksheet.Cell(i + 2, 23).Value = notas[i].CLASSIF_FISCAL;

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
