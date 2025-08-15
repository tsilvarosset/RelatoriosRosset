using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using RelatoriosRosset.Models;
using System.Data;

namespace RelatoriosRosset.Controllers
{
    public class EANPorSaidaController : Controller
    {
        private readonly ApplicationDbContext _context;

        public EANPorSaidaController(ApplicationDbContext context)
        {
            _context = context;
        }

        // GET
        public async Task<IActionResult> EANPorSaida()
        {
            var model = new EANPorSaidaModel();
            await CarregarFiliais();
            return View(model);
        }

        private async Task CarregarFiliais()
        {
            try
            {
                var filiais = await _context.FILIAIS
                    .Select(f => new SelectListItem
                    {
                        Value = f.FILIAL,
                        Text = f.FILIAL
                    })
                    .OrderBy(f => f.Text)
                    .ToListAsync();

                ViewBag.FILIAIS = filiais;
            }
            catch (Exception ex)
            {
                ViewBag.FILIAIS = new List<SelectListItem>();
                TempData["Erro"] = $"Erro ao carregar filiais: {ex.Message}";
            }
        }

        [HttpPost]
        public async Task<IActionResult> ExecutarProcedure(DateTime dataInicial, DateTime dataFinal, string filialOrigem, string filialDestino)
        {
            var model = new EANPorSaidaModel
            {
                EMISSAO = DateTime.Now,
                FILIAL_ORIGEM = filialOrigem,
                FILIAL_DESTINO = filialDestino
            };

            try
            {
                using var connection = new SqlConnection(_context.Database.GetConnectionString());
                await connection.OpenAsync();

                using var command = connection.CreateCommand();
                command.CommandTimeout = 120;
                command.CommandText = "EXEC GERA_SAIDA_POR_EAN @DataInicial, @DataFinal, @FilialOrigem, @FilialDestino";
                command.Parameters.Add(new SqlParameter("@DataInicial", SqlDbType.Date) { Value = dataInicial });
                command.Parameters.Add(new SqlParameter("@DataFinal", SqlDbType.Date) { Value = dataFinal });
                command.Parameters.Add(new SqlParameter("@FilialOrigem", SqlDbType.NVarChar) { Value = string.IsNullOrEmpty(filialOrigem) ? (object)DBNull.Value : filialOrigem });
                command.Parameters.Add(new SqlParameter("@FilialDestino", SqlDbType.NVarChar) { Value = string.IsNullOrEmpty(filialDestino) ? (object)DBNull.Value : filialDestino });

                await command.ExecuteNonQueryAsync();

                TempData["Mensagem"] = "Saídas geradas com sucesso!";
            }
            catch (Exception ex)
            {
                TempData["Erro"] = $"Erro ao executar a procedure: {ex.Message}";
            }

            await CarregarFiliais();
            return View("EANPorSaida", model);
        }

        public async Task<IActionResult> ExportarExcel(DateTime? emissao)
        {
            try
            {
                var query = _context.EAN_POR_SAIDA.AsQueryable();
                if (emissao.HasValue)
                {
                    query = query.Where(v => v.EMISSAO <= emissao.Value);
                }

                var saida = await query.OrderBy(v => v.NF_NUMERO).ToListAsync();
                if (!saida.Any())
                {
                    TempData["Erro"] = "Nenhum dado disponível para exportar.";
                    return RedirectToAction(nameof(EANPorSaida), new { emissao });
                }

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Saidas");

                    worksheet.Cell(1, 1).Value = "NF_NUMERO";
                    worksheet.Cell(1, 2).Value = "SERIE";
                    worksheet.Cell(1, 3).Value = "PRODUTO";
                    worksheet.Cell(1, 4).Value = "CODIGO BARRA";
                    worksheet.Cell(1, 5).Value = "EMISSAO";
                    worksheet.Cell(1, 6).Value = "FILIAL ORIGEM";
                    worksheet.Cell(1, 7).Value = "FILIAL DESTINO";
                    worksheet.Cell(1, 8).Value = "VALOR UNITARIO";
                    worksheet.Cell(1, 9).Value = "QTDE ITEM";
                    worksheet.Cell(1, 10).Value = "VALOR TOTAL";
                    worksheet.Cell(1, 11).Value = "VALOR ICMS";
                    worksheet.Cell(1, 12).Value = "CODIGO FISCAL OPERACAO";
                    worksheet.Cell(1, 13).Value = "CHAVE_NFE";

                    for (int i = 0; i < saida.Count; i++)
                    {
                        worksheet.Cell(i + 2, 1).Value = saida[i].NF_NUMERO;
                        worksheet.Cell(i + 2, 2).Value = saida[i].SERIE;
                        worksheet.Cell(i + 2, 3).Value = saida[i].PRODUTO;
                        worksheet.Cell(i + 2, 4).Value = saida[i].CODIGO_BARRA;
                        worksheet.Cell(i + 2, 5).Value = saida[i].EMISSAO.ToString("dd/MM/yyyy");
                        worksheet.Cell(i + 2, 6).Value = saida[i].FILIAL_ORIGEM;
                        worksheet.Cell(i + 2, 7).Value = saida[i].FILIAL_DESTINO;
                        worksheet.Cell(i + 2, 8).Value = saida[i].VALOR_UNITARIO;
                        worksheet.Cell(i + 2, 9).Value = saida[i].QTDE_ITEM;
                        worksheet.Cell(i + 2, 10).Value = saida[i].VALOR_TOTAL;
                        worksheet.Cell(i + 2, 11).Value = saida[i].VALOR_ICMS;
                        worksheet.Cell(i + 2, 12).Value = saida[i].CODIGO_FISCAL_OPERACAO;
                        worksheet.Cell(i + 2, 13).Value = saida[i].CHAVE_NFE;
                    }

                    worksheet.Columns().AdjustToContents();
                    worksheet.Row(1).Style.Font.Bold = true;

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        stream.Position = 0;
                        string fileName = $"Relatorio_Saida_{DateTime.Now:yyyyMMdd}.xlsx";
                        return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                    }
                }
            }
            catch (Exception ex)
            {
                TempData["Erro"] = $"Erro ao gerar o relatório: {ex.Message}";
                return RedirectToAction(nameof(EANPorSaida), new { emissao });
            }
        }
    }
}
