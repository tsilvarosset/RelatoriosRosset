using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using RelatoriosRosset.Models;
using System.Data;

namespace RelatoriosRosset.Controllers
{
    public class EstoqueEANLojasController : Controller
    {
        private readonly ApplicationDbContext _context;

        public EstoqueEANLojasController(ApplicationDbContext context)
        {
            _context = context;
        }

        public async Task<IActionResult> EstoqueEANLojas()
        {
            await CarregarFiliais();
            var model = new EstoqueEANLojasModel
            {
                DATA_SALDO = DateTime.Now
            };
            return View(model);
        }

        private async Task CarregarFiliais()
        {
            try
            {
                var filiaisData = await _context.V_FILIAIS_ATIVAS_PROPRIAS
                    .OrderBy(f => f.Filial)
                    .ToListAsync();

                var filiais = filiaisData
                    .Select(f => new SelectListItem
                    {
                        Value = f.Filial,
                        Text = $"{f.Filial} - CNPJ {f.Cgc_Cpf}"
                    })
                    .ToList();

                Console.WriteLine($"Filiais carregadas: {filiais.Count}");
                foreach (var filial in filiais)
                {
                    Console.WriteLine($"Filial: {filial.Text}, Código: {filial.Value}");
                }

                ViewBag.V_FILIAIS_ATIVAS_PROPRIAS = filiais;
            }
            catch (Exception ex)
            {
                ViewBag.V_FILIAIS_ATIVAS_PROPRIAS = new List<SelectListItem>();
                TempData["Erro"] = $"Erro ao carregar filiais: {ex.Message}. StackTrace: {ex.StackTrace}";
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner Exception: {ex.InnerException.Message}\n{ex.InnerException.StackTrace}");
                }
            }
        }

        [HttpPost]
        public async Task<IActionResult> ExecutarProcedure(DateTime dataSaldo, string filial)
        {
            var model = new EstoqueEANLojasModel
            {
                DATA_SALDO = DateTime.Now,
                FILIAL = filial
            };

            try
            {
                using var connection = new SqlConnection(_context.Database.GetConnectionString());
                await connection.OpenAsync();

                using var command = connection.CreateCommand();
                command.CommandTimeout = 120;
                command.CommandText = "EXEC ESTOQUE_EAN_LOJA @DataFim, @Filial";
                command.Parameters.Add(new SqlParameter("@DataFim", SqlDbType.Date) { Value = dataSaldo });
                command.Parameters.Add(new SqlParameter("@Filial", SqlDbType.NVarChar) { Value = string.IsNullOrEmpty(filial) ? (object)DBNull.Value : filial });

                await command.ExecuteNonQueryAsync();

                TempData["Mensagem"] = "Estoque Gerado com Sucesso!";
            }
            catch (Exception ex)
            {
                TempData["Erro"] = $"Erro ao executar a procedure: {ex.Message}";
            }

            await CarregarFiliais();
            return View("EstoqueEANLojas", model);
        }

        public async Task<IActionResult> ExportarExcel(DateTime? dataSaldo)
        {
            try
            {
                var query = _context.TABELA_ESTOQUE_EAN_LOJA.AsQueryable();
                if (dataSaldo.HasValue)
                {
                    query = query.Where(v => v.DATA_SALDO <= dataSaldo.Value);
                }

                var estoque = await query.OrderBy(v => v.PRODUTO).ToListAsync();
                if (!estoque.Any())
                {
                    TempData["Erro"] = "Nenhum dado disponível para exportar.";
                    return RedirectToAction(nameof(EstoqueEANLojas), new { dataSaldo });
                }

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Estoque");

                    worksheet.Cell(1, 1).Value = "FILIAL";
                    worksheet.Cell(1, 2).Value = "CODIGO BARRA";
                    worksheet.Cell(1, 3).Value = "PRODUTO";
                    worksheet.Cell(1, 4).Value = "ESTOQUE";
                    worksheet.Cell(1, 5).Value = "VALOR VENDA";
                    worksheet.Cell(1, 6).Value = "DATA SALDO";

                    for (int i = 0; i < estoque.Count; i++)
                    {
                        worksheet.Cell(i + 2, 1).Value = estoque[i].FILIAL;
                        worksheet.Cell(i + 2, 2).Value = estoque[i].CODIGO_BARRA;
                        worksheet.Cell(i + 2, 3).Value = estoque[i].PRODUTO;
                        worksheet.Cell(i + 2, 4).Value = estoque[i].ESTOQUE;
                        worksheet.Cell(i + 2, 5).Value = estoque[i].VALOR_VENDA;
                        worksheet.Cell(i + 2, 6).Value = estoque[i].DATA_SALDO;
                    }

                    worksheet.Columns().AdjustToContents();
                    worksheet.Row(1).Style.Font.Bold = true;

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        stream.Position = 0;
                        string fileName = $"Relatorio_Estoque_{DateTime.Now:yyyyMMdd}.xlsx";
                        return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                    }
                }
            }
            catch (Exception ex)
            {
                TempData["Erro"] = $"Erro ao gerar o relatório: {ex.Message}";
                return RedirectToAction(nameof(EstoqueEANLojas), new { dataSaldo });
            }
        }
    }
}
