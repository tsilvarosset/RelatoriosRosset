using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using RelatoriosRosset.Models;
using Microsoft.AspNetCore.Http;
using System.Text.Json;

namespace RelatoriosRosset.Controllers
{
    public static class SessionExtensions
    {
        public static void SetObjectAsJson(this ISession session, string key, object value)
        {
            session.SetString(key, JsonSerializer.Serialize(value));
        }

        public static T GetObjectFromJson<T>(this ISession session, string key)
        {
            var value = session.GetString(key);
            return value == null ? default : JsonSerializer.Deserialize<T>(value);
        }
    }

    public class LojaCaixaCartaoController : Controller
    {
        private readonly ApplicationDbContext _context;

        public LojaCaixaCartaoController(ApplicationDbContext context)
        {
            _context = context;
        }

        public async Task<IActionResult> LojaCaixaCartao(DateTime? dataInicio, DateTime? dataFim)
        {
            var tickets = new List<LojaCaixaCartaoModel>();

            if (dataInicio.HasValue && dataFim.HasValue)
            {
                var query = _context.V_LANCAMENTOS_CAIXA_CARTAO.AsNoTracking().AsQueryable();

                query = query.Where(v => v.DATA_VENDA >= dataInicio.Value && v.DATA_VENDA <= dataFim.Value);

                tickets = await query
                    .OrderByDescending(v => v.DATA_VENDA)
                    .Select(v => new LojaCaixaCartaoModel
                    {
                        CODIGO_FILIAL = v.CODIGO_FILIAL,
                        FILIAL = v.FILIAL,
                        VENDEDOR = v.VENDEDOR,
                        TICKET = v.TICKET,
                        LANCAMENTO_CAIXA = v.LANCAMENTO_CAIXA,
                        TERMINAL = v.TERMINAL,
                        DATA_VENDA = v.DATA_VENDA,
                        CODIGO_CONSUMIDOR = v.CODIGO_CONSUMIDOR,
                        PARCELA = v.PARCELA,
                        NUMERO_CHEQUE_CARTAO = v.NUMERO_CHEQUE_CARTAO,
                        NUMERO_APROVACAO_CARTAO = v.NUMERO_APROVACAO_CARTAO,
                        NUMERO_TITULO = v.NUMERO_TITULO,
                        VALOR_ORIGINAL = v.VALOR_ORIGINAL,
                        TAXA_ADMINISTRACAO = v.TAXA_ADMINISTRACAO,
                        VALOR_A_RECEBER = v.VALOR_A_RECEBER,
                        DATA_HORA_TEF = v.DATA_HORA_TEF,
                        DATA_EMISSAO = v.DATA_EMISSAO,
                        VENCIMENTO_REAL = v.VENCIMENTO_REAL,
                        DESC_TIPO_PGTO = v.DESC_TIPO_PGTO,
                        CMC7_CVCARTAO = v.CMC7_CVCARTAO,
                        LANCAMENTO = v.LANCAMENTO
                    })
                    .ToListAsync();
            }

            ViewBag.DataInicio = dataInicio?.ToString("yyyy-MM-dd");
            ViewBag.DataFim = dataFim?.ToString("yyyy-MM-dd");

            return View(tickets);
        }

        public async Task<IActionResult> ExportarExcel(DateTime? dataInicio, DateTime? dataFim)
        {
            try
            {
                if (!dataInicio.HasValue || !dataFim.HasValue)
                {
                    TempData["Erro"] = "Você precisa informar o período para exportar o relatório.";
                    return RedirectToAction(nameof(LojaCaixaCartao));
                }

                var query = _context.V_LANCAMENTOS_CAIXA_CARTAO.AsNoTracking().AsQueryable();

                query = query.Where(v => v.DATA_VENDA >= dataInicio.Value && v.DATA_VENDA <= dataFim.Value);

                var tickets = await query
                    .OrderBy(v => v.FILIAL)
                    .Select(v => new LojaCaixaCartaoModel
                    {
                        CODIGO_FILIAL = v.CODIGO_FILIAL,
                        FILIAL = v.FILIAL,
                        VENDEDOR = v.VENDEDOR,
                        TICKET = v.TICKET,
                        LANCAMENTO_CAIXA = v.LANCAMENTO_CAIXA,
                        TERMINAL = v.TERMINAL,
                        DATA_VENDA = v.DATA_VENDA,
                        CODIGO_CONSUMIDOR = v.CODIGO_CONSUMIDOR,
                        PARCELA = v.PARCELA,
                        NUMERO_CHEQUE_CARTAO = v.NUMERO_CHEQUE_CARTAO,
                        NUMERO_APROVACAO_CARTAO = v.NUMERO_APROVACAO_CARTAO,
                        NUMERO_TITULO = v.NUMERO_TITULO,
                        VALOR_ORIGINAL = v.VALOR_ORIGINAL,
                        TAXA_ADMINISTRACAO = v.TAXA_ADMINISTRACAO,
                        VALOR_A_RECEBER = v.VALOR_A_RECEBER,
                        DATA_HORA_TEF = v.DATA_HORA_TEF,
                        DATA_EMISSAO = v.DATA_EMISSAO,
                        VENCIMENTO_REAL = v.VENCIMENTO_REAL,
                        DESC_TIPO_PGTO = v.DESC_TIPO_PGTO,
                        CMC7_CVCARTAO = v.CMC7_CVCARTAO,
                        LANCAMENTO = v.LANCAMENTO
                    })
                    .ToListAsync();

                if (!tickets.Any())
                {
                    TempData["Erro"] = "Nenhum dado encontrado para os filtros selecionados.";
                    return RedirectToAction(nameof(LojaCaixaCartao), new { dataInicio, dataFim });
                }

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Vendas");

                    worksheet.Cell(1, 1).Value = "CODIGO FILIAL";
                    worksheet.Cell(1, 2).Value = "FILIAL";
                    worksheet.Cell(1, 3).Value = "VENDEDOR";
                    worksheet.Cell(1, 4).Value = "TICKET";
                    worksheet.Cell(1, 5).Value = "LANCAMENTO CAIXA";
                    worksheet.Cell(1, 6).Value = "TERMINAL";
                    worksheet.Cell(1, 7).Value = "DATA VENDA";
                    worksheet.Cell(1, 8).Value = "CODIGO CONSUMIDOR";
                    worksheet.Cell(1, 9).Value = "PARCELA";
                    worksheet.Cell(1, 10).Value = "NUMERO CHEQUE CARTAO";
                    worksheet.Cell(1, 11).Value = "NUMERO APROVACAO CARTAO";
                    worksheet.Cell(1, 12).Value = "NUMERO TITULO";
                    worksheet.Cell(1, 13).Value = "VALOR ORIGINAL";
                    worksheet.Cell(1, 14).Value = "TAXA ADMINISTRACAO";
                    worksheet.Cell(1, 15).Value = "VALOR A RECEBER";
                    worksheet.Cell(1, 16).Value = "DATA HORA TEF";
                    worksheet.Cell(1, 17).Value = "DATA EMISSAO";
                    worksheet.Cell(1, 18).Value = "VENCIMENTO REAL";
                    worksheet.Cell(1, 19).Value = "DESC TIPO PGTO";
                    worksheet.Cell(1, 20).Value = "CMC7 CVCARTAO";
                    worksheet.Cell(1, 21).Value = "LANCAMENTO";

                    for (int i = 0; i < tickets.Count; i++)
                    {
                        worksheet.Cell(i + 2, 1).Value = tickets[i].CODIGO_FILIAL;
                        worksheet.Cell(i + 2, 2).Value = tickets[i].FILIAL;
                        worksheet.Cell(i + 2, 3).Value = tickets[i].VENDEDOR;
                        worksheet.Cell(i + 2, 4).Value = tickets[i].TICKET;
                        worksheet.Cell(i + 2, 5).Value = tickets[i].LANCAMENTO_CAIXA;
                        worksheet.Cell(i + 2, 6).Value = tickets[i].TERMINAL;
                        worksheet.Cell(i + 2, 7).Value = tickets[i].DATA_VENDA;
                        worksheet.Cell(i + 2, 8).Value = tickets[i].CODIGO_CONSUMIDOR;
                        worksheet.Cell(i + 2, 9).Value = tickets[i].PARCELA;
                        worksheet.Cell(i + 2, 10).Value = tickets[i].NUMERO_CHEQUE_CARTAO;
                        worksheet.Cell(i + 2, 11).Value = tickets[i].NUMERO_APROVACAO_CARTAO;
                        worksheet.Cell(i + 2, 12).Value = tickets[i].NUMERO_TITULO;
                        worksheet.Cell(i + 2, 13).Value = tickets[i].VALOR_ORIGINAL;
                        worksheet.Cell(i + 2, 14).Value = tickets[i].TAXA_ADMINISTRACAO;
                        worksheet.Cell(i + 2, 15).Value = tickets[i].VALOR_A_RECEBER;
                        worksheet.Cell(i + 2, 16).Value = tickets[i].DATA_HORA_TEF;
                        worksheet.Cell(i + 2, 17).Value = tickets[i].DATA_EMISSAO;
                        worksheet.Cell(i + 2, 18).Value = tickets[i].VENCIMENTO_REAL;
                        worksheet.Cell(i + 2, 19).Value = tickets[i].DESC_TIPO_PGTO;
                        worksheet.Cell(i + 2, 20).Value = tickets[i].CMC7_CVCARTAO;
                        worksheet.Cell(i + 2, 21).Value = tickets[i].LANCAMENTO;
                    }

                    worksheet.Columns().AdjustToContents();
                    worksheet.Row(1).Style.Font.Bold = true;

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
                TempData["Erro"] = $"Erro ao gerar o relatório: {ex.Message}";
                return RedirectToAction(nameof(LojaCaixaCartao), new { dataInicio, dataFim });
            }
        }

    }
}