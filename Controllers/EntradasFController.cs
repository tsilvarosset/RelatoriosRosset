using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using RelatoriosRosset.Models;
using System.Data;

namespace RelatoriosRosset.Controllers
{
    public class EntradasFController : Controller
    {
        private readonly ApplicationDbContext _context;

        public EntradasFController(ApplicationDbContext context)
        {
            _context = context;
        }

        public async Task<IActionResult> EntradasF(DateTime? dataInicio, DateTime? dataFim)
        {
            try
            {
                var query = _context.TABELA_ENTRADAS_F.AsQueryable();

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
                TempData["Erro"] = $"Erro ao acessar TABELA_ENTRADAS_F: {ex.Message}";
                return View(new List<EntradasFModel>());
            }
            catch (Exception ex)
            {
                Console.WriteLine($"General Error: {ex.Message}\n{ex.StackTrace}");
                TempData["Erro"] = $"Erro inesperado: {ex.Message}";
                return View(new List<EntradasFModel>());
            }
        }

        public async Task<IActionResult> ExecutarProcedureEntradasF(DateTime dataInicio, DateTime dataFim)
        {
            try
            {
                // Execute the stored procedure
                using var connection = new SqlConnection(_context.Database.GetConnectionString());
                await connection.OpenAsync();
                using var command = connection.CreateCommand();
                command.CommandTimeout = 12000000;
                command.CommandText = "EXEC GERA_ENTRADAS_F @dataIni, @dataFim";
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
            return RedirectToAction(nameof(EntradasF), new { dataInicio, dataFim });
        }

        public async Task<IActionResult> ExportarExcel(DateTime? dataInicio, DateTime? dataFim)
        {
            try
            {
                var query = _context.TABELA_ENTRADAS_F.AsQueryable();

                if (dataInicio.HasValue)
                {
                    query = query.Where(v => v.RECEBIMENTO >= dataInicio.Value);
                }

                if (dataFim.HasValue)
                {
                    query = query.Where(v => v.RECEBIMENTO <= dataFim.Value);
                }

                var entradas = await query.OrderBy(v => v.NF_ENTRADA).ToListAsync();

                if (!entradas.Any())
                {
                    return RedirectToAction(nameof(EntradasF), new { dataInicio, dataFim });
                }

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Entradas");

                    // Add headers
                    worksheet.Cell(1, 1).Value = "Filial";
                    worksheet.Cell(1, 2).Value = "Origem";
                    worksheet.Cell(1, 3).Value = "NF Entrada";
                    worksheet.Cell(1, 4).Value = "Serie";
                    worksheet.Cell(1, 5).Value = "CFOP";
                    worksheet.Cell(1, 6).Value = "Recebimento";
                    worksheet.Cell(1, 7).Value = "Valor Contábil";
                    worksheet.Cell(1, 8).Value = "Base Imposto";
                    worksheet.Cell(1, 9).Value = "Alíquota";
                    worksheet.Cell(1, 10).Value = "Valor ICMS";
                    worksheet.Cell(1, 11).Value = "Chave NFE";

                    // Add data
                    for (int i = 0; i < entradas.Count; i++)
                    {
                        worksheet.Cell(i + 2, 1).Value = entradas[i].FILIAL;
                        worksheet.Cell(i + 2, 2).Value = entradas[i].ORIGEM;
                        worksheet.Cell(i + 2, 3).Value = entradas[i].NF_ENTRADA;
                        worksheet.Cell(i + 2, 4).Value = entradas[i].SERIE;
                        worksheet.Cell(i + 2, 5).Value = entradas[i].CFOP;
                        worksheet.Cell(i + 2, 6).Value = entradas[i].RECEBIMENTO.ToString("dd/MM/yyyy");
                        worksheet.Cell(i + 2, 7).Value = entradas[i].VALOR_CONTABIL;
                        worksheet.Cell(i + 2, 8).Value = entradas[i].BASE_IMPOSTO;
                        worksheet.Cell(i + 2, 9).Value = entradas[i].ALIQUOTA;
                        worksheet.Cell(i + 2, 10).Value = entradas[i].VALOR_ICMS;
                        worksheet.Cell(i + 2, 11).Value = entradas[i].CHAVE_NFE;
                    }

                    // Adjust column widths and format headers
                    worksheet.Columns().AdjustToContents();
                    worksheet.Row(1).Style.Font.Bold = true;

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        stream.Position = 0;
                        string fileName = $"Relatorio_Entrada_{DateTime.Now:yyyyMMdd}.xlsx";
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
