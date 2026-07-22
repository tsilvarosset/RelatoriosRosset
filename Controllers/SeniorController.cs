using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using RelatoriosRosset.Models;
using System.Data;

namespace RelatoriosRosset.Controllers
{
    public class SeniorController : Controller
    {
        private readonly ApplicationDbContext _context;

        public SeniorController(ApplicationDbContext context)
        {
            _context = context;
        }

        // GET: Exibe a view com o formulário
        public async Task<IActionResult> Senior(DateTime? dataInicio, DateTime? dataFim)
        {
            try
            {
                var query = _context.TABELA_SAIDAS_SENIOR.AsQueryable();

                if (dataInicio.HasValue)
                {
                    query = query.Where(v => v.EMISSAO >= dataInicio.Value);
                }

                if (dataFim.HasValue)
                {
                    query = query.Where(v => v.EMISSAO <= dataFim.Value);
                }

                var saidas = await query.OrderByDescending(v => v.EMISSAO).Take(10).ToListAsync();

                ViewBag.DataInicio = dataInicio;
                ViewBag.DataFim = dataFim;

                return View(saidas);
            }
            catch (SqlException ex)
            {
                Console.WriteLine($"SQL Error: {ex.Message}\n{ex.StackTrace}");
                TempData["Erro"] = $"Erro ao acessar TABELA_SAIDAS_SENIOR: {ex.Message}";
                return View(new List<SeniorModel>());
            }
            catch (Exception ex)
            {
                Console.WriteLine($"General Error: {ex.Message}\n{ex.StackTrace}");
                TempData["Erro"] = $"Erro inesperado: {ex.Message}";
                return View(new List<SeniorModel>());
            }
        }

        public async Task<IActionResult> ExecutarProcedureSenior(DateTime dataInicio, DateTime dataFim)
        {
            try
            {
                // Execute the stored procedure
                using var connection = new SqlConnection(_context.Database.GetConnectionString());
                await connection.OpenAsync();
                using var command = connection.CreateCommand();
                command.CommandTimeout = 12000000;
                command.CommandText = "GERA_SAIDAS_INTERM_SENIOR @dataIni, @dataFim";
                command.Parameters.Add(new SqlParameter
                {
                    ParameterName = "@dataIni",
                    SqlDbType = SqlDbType.Date,
                    Value = dataInicio
                });
                command.Parameters.Add(new SqlParameter
                {
                    ParameterName = "@dataFim",
                    SqlDbType = SqlDbType.Date,
                    Value = dataFim
                });

                await command.ExecuteNonQueryAsync();
                Console.WriteLine("Procedure executada com sucesso.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao executar a procedure: {ex.Message}\n{ex.StackTrace}");
                TempData["Erro"] = $"Erro ao executar a procedure: {ex.Message}";
            }

            // Redirect to SaidasF with the date parameters
            return RedirectToAction(nameof(Senior), new { dataInicio, dataFim });
        }

        public async Task<IActionResult> ExportarExcel(DateTime? dataInicio, DateTime? dataFim)
        {
            try
            {
                var query = _context.TABELA_SAIDAS_SENIOR.AsQueryable();

                if (dataInicio.HasValue)
                {
                    query = query.Where(v => v.EMISSAO >= dataInicio.Value);
                }

                if (dataFim.HasValue)
                {
                    query = query.Where(v => v.EMISSAO <= dataFim.Value);
                }

                var saidas = await query.OrderBy(v => v.EMISSAO).ToListAsync();

                if (!saidas.Any())
                {
                    return RedirectToAction(nameof(Senior), new { dataInicio, dataFim });
                }

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Saidas");

                    // Add headers
                    worksheet.Cell(1, 1).Value = "CHAVE NFE";
                    worksheet.Cell(1, 2).Value = "CODIGO FILIAL";
                    worksheet.Cell(1, 3).Value = "NOTA";
                    worksheet.Cell(1, 4).Value = "SERIE";
                    worksheet.Cell(1, 5).Value = "EMISSAO";
                    worksheet.Cell(1, 6).Value = "CFOP";
                    worksheet.Cell(1, 7).Value = "VALOR CONTABIL";
                    worksheet.Cell(1, 8).Value = "BASE ICMS";
                    worksheet.Cell(1, 9).Value = "IMPOSTO ICMS";
                    worksheet.Cell(1, 10).Value = "BASE IPI";
                    worksheet.Cell(1, 11).Value = "IMPOSTO IPI";
                    worksheet.Cell(1, 12).Value = "BASE PIS";
                    worksheet.Cell(1, 13).Value = "IMPOSTO PIS";
                    worksheet.Cell(1, 14).Value = "BASE COFINS";
                    worksheet.Cell(1, 15).Value = "IMPOSTO COFINS";
                    worksheet.Cell(1, 16).Value = "BASE FCP";
                    worksheet.Cell(1, 17).Value = "IMPOSTO FCP";

                    // Add data
                    for (int i = 0; i < saidas.Count; i++)
                    {
                        worksheet.Cell(i + 2, 1).Value = saidas[i].CHAVE_NFE;
                        worksheet.Cell(i + 2, 2).Value = saidas[i].CODIGO_FILIAL;
                        worksheet.Cell(i + 2, 3).Value = saidas[i].NOTA;
                        worksheet.Cell(i + 2, 4).Value = saidas[i].SERIE;
                        worksheet.Cell(i + 2, 5).Value = saidas[i].EMISSAO.ToString("dd/MM/yyyy");
                        worksheet.Cell(i + 2, 6).Value = saidas[i].CFOP;
                        worksheet.Cell(i + 2, 7).Value = saidas[i].VALOR_CONTABIL;
                        worksheet.Cell(i + 2, 8).Value = saidas[i].BASE_ICMS;
                        worksheet.Cell(i + 2, 9).Value = saidas[i].IMPOSTO_ICMS;
                        worksheet.Cell(i + 2, 10).Value = saidas[i].BASE_IPI;
                        worksheet.Cell(i + 2, 11).Value = saidas[i].IMPOSTO_IPI;
                        worksheet.Cell(i + 2, 12).Value = saidas[i].BASE_PIS;
                        worksheet.Cell(i + 2, 13).Value = saidas[i].IMPOSTO_PIS;
                        worksheet.Cell(i + 2, 14).Value = saidas[i].BASE_COFINS;
                        worksheet.Cell(i + 2, 15).Value = saidas[i].IMPOSTO_COFINS;
                        worksheet.Cell(i + 2, 16).Value = saidas[i].BASE_FCP;
                        worksheet.Cell(i + 2, 17).Value = saidas[i].IMPOSTO_FCP;
                    }

                    // Adjust column widths and format headers
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
                return Content($"Erro ao gerar o relatório: {ex.Message}\nStack Trace: {ex.StackTrace}");
            }
        }
    }
}
