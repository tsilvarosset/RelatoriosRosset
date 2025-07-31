using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using RelatoriosRosset.Models;
using System.Data;

namespace RelatoriosRosset.Controllers
{
    public class EstoqueEANController : Controller
    {
        private readonly ApplicationDbContext _context;

        public EstoqueEANController(ApplicationDbContext context)
        {
            _context = context;
        }

        public async Task<IActionResult> EstoqueEAN(DateTime? dataSaldo)
        {
            try
            {
                var query = _context.TABELA_ESTOQUE_EAN_CUSTO.AsQueryable();
                if (dataSaldo.HasValue)
                {
                    query = query.Where(v => v.DATA_SALDO <= dataSaldo.Value);
                }

                var estoque = await query.OrderByDescending(v => v.DATA_SALDO).Take(10).ToListAsync();
                ViewBag.DATA_SALDO = dataSaldo;
                return View(estoque);
            }
            catch (SqlException ex)
            {
                Console.WriteLine($"SQL Error: {ex.Message}\n{ex.StackTrace}");
                TempData["Erro"] = $"Erro ao acessar TABELA_ESTOQUE_EAN_CUSTO: {ex.Message}";
                return View(new List<EstoqueEANModel>());
            }
            catch (Exception ex)
            {
                Console.WriteLine($"General Error: {ex.Message}\n{ex.StackTrace}");
                TempData["Erro"] = $"Erro inesperado: {ex.Message}";
                return View(new List<EstoqueEANModel>());
            }
        }

        public async Task<IActionResult> ExecutarProcedureEstoque(DateTime dataFinal)
        {
            try
            {
                // Execute the stored procedure
                using var connection = new SqlConnection(_context.Database.GetConnectionString());
                await connection.OpenAsync();
                using var command = connection.CreateCommand();
                command.CommandTimeout = 12000000;
                command.CommandText = "EXEC ESTOQUE_EAN_CUSTO @DataFinal";
                command.Parameters.Add(new SqlParameter
                {
                    ParameterName = "@DataFinal",
                    SqlDbType = SqlDbType.Date,
                    Value = dataFinal
                });
                await command.ExecuteNonQueryAsync();
                Console.WriteLine("Procedure executada com sucesso.");

                // Query the updated data
                var query = _context.TABELA_ESTOQUE_EAN_CUSTO.AsQueryable();
                if (dataFinal != default)
                {
                    query = query.Where(v => v.DATA_SALDO <= dataFinal);
                }
                var estoque = await query.OrderByDescending(v => v.DATA_SALDO).Take(10).ToListAsync();

                ViewBag.DATA_SALDO = dataFinal;
                return View("EstoqueEAN", estoque); // Pass List<EstoqueEANModel>
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao executar a procedure: {ex.Message}\n{ex.StackTrace}");
                TempData["Erro"] = $"Erro ao executar a procedure: {ex.Message}";
                return RedirectToAction(nameof(EstoqueEAN), new { dataSaldo = dataFinal });
            }
        }

        public async Task<IActionResult> ExportarExcel(DateTime? dataSaldo)
        {
            try
            {
                var query = _context.TABELA_ESTOQUE_EAN_CUSTO.AsQueryable();
                if (dataSaldo.HasValue)
                {
                    query = query.Where(v => v.DATA_SALDO <= dataSaldo.Value);
                }

                var estoque = await query.OrderBy(v => v.DATA_SALDO).ToListAsync();
                if (!estoque.Any())
                {
                    TempData["Erro"] = "Nenhum dado disponível para exportar.";
                    return RedirectToAction(nameof(EstoqueEAN), new { dataSaldo });
                }

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Estoque");

                    // Add headers
                    worksheet.Cell(1, 1).Value = "Filial";
                    worksheet.Cell(1, 2).Value = "Código de Barras";
                    worksheet.Cell(1, 3).Value = "Produto";
                    worksheet.Cell(1, 4).Value = "Estoque";
                    worksheet.Cell(1, 5).Value = "Último Custo";
                    worksheet.Cell(1, 6).Value = "Data Saldo";

                    // Add data
                    for (int i = 0; i < estoque.Count; i++)
                    {
                        worksheet.Cell(i + 2, 1).Value = estoque[i].FILIAL;
                        worksheet.Cell(i + 2, 2).Value = estoque[i].CODIGO_BARRA;
                        worksheet.Cell(i + 2, 3).Value = estoque[i].PRODUTO;
                        worksheet.Cell(i + 2, 4).SetValue(estoque[i].ESTOQUE); // Handle type conversion
                        worksheet.Cell(i + 2, 5).SetValue(estoque[i].ULTIMO_CUSTO); // Handle type conversion
                        worksheet.Cell(i + 2, 6).Value = estoque[i].DATA_SALDO.ToString("dd/MM/yyyy");
                    }

                    // Adjust column widths and format headers
                    worksheet.Columns().AdjustToContents();
                    worksheet.Row(1).Style.Font.Bold = true;

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        stream.Position = 0;
                        string fileName = $"Relatorio_estoque_{DateTime.Now:yyyyMMdd}.xlsx";
                        return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao gerar o relatório: {ex.Message}\n{ex.StackTrace}");
                TempData["Erro"] = $"Erro ao gerar o relatório: {ex.Message}";
                return RedirectToAction(nameof(EstoqueEAN), new { dataSaldo });
            }
        }
    }
}
