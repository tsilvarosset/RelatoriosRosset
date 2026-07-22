using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using RelatoriosRosset.Models;
using System.Data;

namespace RelatoriosRosset.Controllers
{
    public class SeniorEntradasController : Controller
    {
        private readonly ApplicationDbContext _context;

        public SeniorEntradasController(ApplicationDbContext context)
        {
            _context = context;
        }

        // GET: Exibe a view com o formulário
        public async Task<IActionResult> SeniorEntradas(DateTime? dataInicio, DateTime? dataFim)
        {
            try
            {
                var query = _context.TABELA_ENTRADAS_SENIOR.AsQueryable();

                if (dataInicio.HasValue)
                {
                    query = query.Where(v => v.RECEBIMENTO >= dataInicio.Value);
                }

                if (dataFim.HasValue)
                {
                    query = query.Where(v => v.RECEBIMENTO <= dataFim.Value);
                }

                var entradas = await query.OrderByDescending(v => v.RECEBIMENTO).Take(10).ToListAsync();

                ViewBag.DataInicio = dataInicio;
                ViewBag.DataFim = dataFim;

                return View(entradas);
            }
            catch (SqlException ex)
            {
                Console.WriteLine($"SQL Error: {ex.Message}\n{ex.StackTrace}");
                TempData["Erro"] = $"Erro ao acessar TABELA_ENTRADAS_SENIOR: {ex.Message}";
                return View(new List<SeniorEntradasModel>());
            }
            catch (Exception ex)
            {
                Console.WriteLine($"General Error: {ex.Message}\n{ex.StackTrace}");
                TempData["Erro"] = $"Erro inesperado: {ex.Message}";
                return View(new List<SeniorEntradasModel>());
            }
        }

        public async Task<IActionResult> ExecutarProcedureSeniorEntradas(DateTime dataInicio, DateTime dataFim)
        {
            try
            {
                // Execute the stored procedure
                using var connection = new SqlConnection(_context.Database.GetConnectionString());
                await connection.OpenAsync();
                using var command = connection.CreateCommand();
                command.CommandTimeout = 12000000;
                command.CommandText = "GERA_ENTRADAS_INTERM_SENIOR @dataIni, @dataFim";
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

            
            return RedirectToAction(nameof(SeniorEntradas), new { dataInicio, dataFim });
        }

        public async Task<IActionResult> ExportarExcel(DateTime? dataInicio, DateTime? dataFim)
        {
            try
            {
                var query = _context.TABELA_ENTRADAS_SENIOR.AsQueryable();

                if (dataInicio.HasValue)
                {
                    query = query.Where(v => v.RECEBIMENTO >= dataInicio.Value);
                }

                if (dataFim.HasValue)
                {
                    query = query.Where(v => v.RECEBIMENTO <= dataFim.Value);
                }

                var entradas = await query.OrderBy(v => v.RECEBIMENTO).ToListAsync();

                if (!entradas.Any())
                {
                    return RedirectToAction(nameof(SeniorEntradas), new { dataInicio, dataFim });
                }

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Entradas");

                    // Add headers
                    worksheet.Cell(1, 1).Value = "CHAVE NFE";
                    worksheet.Cell(1, 2).Value = "CODIGO FILIAL";
                    worksheet.Cell(1, 3).Value = "NOTA";
                    worksheet.Cell(1, 4).Value = "SERIE";
                    worksheet.Cell(1, 5).Value = "RECEBIMENTO";
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
                    for (int i = 0; i < entradas.Count; i++)
                    {
                        worksheet.Cell(i + 2, 1).Value = entradas[i].CHAVE_NFE;
                        worksheet.Cell(i + 2, 2).Value = entradas[i].CODIGO_FILIAL;
                        worksheet.Cell(i + 2, 3).Value = entradas[i].NOTA;
                        worksheet.Cell(i + 2, 4).Value = entradas[i].SERIE;
                        worksheet.Cell(i + 2, 5).Value = entradas[i].RECEBIMENTO.ToString("dd/MM/yyyy");
                        worksheet.Cell(i + 2, 6).Value = entradas[i].CFOP;
                        worksheet.Cell(i + 2, 7).Value = entradas[i].VALOR_CONTABIL;
                        worksheet.Cell(i + 2, 8).Value = entradas[i].BASE_ICMS;
                        worksheet.Cell(i + 2, 9).Value = entradas[i].IMPOSTO_ICMS;
                        worksheet.Cell(i + 2, 10).Value = entradas[i].BASE_IPI;
                        worksheet.Cell(i + 2, 11).Value = entradas[i].IMPOSTO_IPI;
                        worksheet.Cell(i + 2, 12).Value = entradas[i].BASE_PIS;
                        worksheet.Cell(i + 2, 13).Value = entradas[i].IMPOSTO_PIS;
                        worksheet.Cell(i + 2, 14).Value = entradas[i].BASE_COFINS;
                        worksheet.Cell(i + 2, 15).Value = entradas[i].IMPOSTO_COFINS;
                        worksheet.Cell(i + 2, 16).Value = entradas[i].BASE_FCP;
                        worksheet.Cell(i + 2, 17).Value = entradas[i].IMPOSTO_FCP;
                    }

                    // Adjust column widths and format headers
                    worksheet.Columns().AdjustToContents();
                    worksheet.Row(1).Style.Font.Bold = true;

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        stream.Position = 0;
                        string fileName = $"Relatorio_Entradas_{DateTime.Now:yyyyMMdd}.xlsx";
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
