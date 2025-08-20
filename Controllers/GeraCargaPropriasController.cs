using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using RelatoriosRosset.Models;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace RelatoriosRosset.Controllers
{
    public class GeraCargaPropriasController : Controller
    {
        private readonly ApplicationDbContext _context;

        public GeraCargaPropriasController(ApplicationDbContext context)
        {
            _context = context;
        }

        // GET: GeraCargaProprias
        public async Task<IActionResult> GeraCargaProprias()
        {
            var model = new List<GeraCargaPropriasModel>(); // Always return a list
            await CarregarFiliais();
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
        public async Task<IActionResult> ExecutarProcedure(string filial)
        {
            Console.WriteLine($"Executando ExecutarProcedure com Filial: '{filial}' (Tamanho: {filial?.Length})");
            var model = new List<GeraCargaPropriasModel>(); // Initialize as a list

            if (string.IsNullOrEmpty(filial))
            {
                TempData["Mensagem"] = "Por favor, selecione uma filial.";
                await CarregarFiliais();
                return View("GeraCargaProprias", model);
            }

            try
            {
                var filialExists = await _context.V_FILIAIS_ATIVAS_PROPRIAS
                    .AnyAsync(f => f.Filial == filial);
                if (!filialExists)
                {
                    TempData["Mensagem"] = $"Filial '{filial}' não encontrada.";
                    await CarregarFiliais();
                    return View("GeraCargaProprias", model);
                }

                using var connection = new SqlConnection(_context.Database.GetConnectionString());
                await connection.OpenAsync();

                using var command = connection.CreateCommand();
                command.CommandTimeout = 300000000;
                command.CommandText = "GERA_CARGA_INVENT_PROPRIAS @Filial";
                command.Parameters.Add(new SqlParameter
                {
                    ParameterName = "@Filial",
                    SqlDbType = SqlDbType.NVarChar,
                    Value = filial
                });

                await command.ExecuteNonQueryAsync();

                var dados = await _context.TABELA_CARGA_INV_PROPRIAS.ToListAsync();
                Console.WriteLine($"Registros inseridos na tabela: {dados.Count}");
                foreach (var item in dados)
                {
                    Console.WriteLine($"FILIAL: {item.FILIAL}");
                }

                TempData["Mensagem"] = "Saldo Gerado com sucesso!";
                Console.WriteLine("Saldo Gerado com sucesso!");
            }
            catch (Exception ex)
            {
                TempData["Mensagem"] = $"Erro ao executar a procedure: {ex.Message}. StackTrace: {ex.StackTrace}";
                Console.WriteLine($"Erro ao executar a procedure: {ex.Message}\n{ex.StackTrace}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner Exception: {ex.InnerException.Message}\n{ex.InnerException.StackTrace}");
                }
            }

            await CarregarFiliais();
            return View("GeraCargaProprias", model); // Ensure correct view name
        }

        public async Task<IActionResult> ExportarExcel()
        {
            try
            {
                var query = _context.TABELA_CARGA_INV_PROPRIAS.AsQueryable();
                var carga = await query.OrderBy(v => v.FILIAL).ToListAsync();

                //if (!carga.Any())
                //{
                //    TempData["Mensagem"] = "Nenhum dado disponível para exportar.";
                //    return RedirectToAction(nameof(GeraCargaProprias));
                //}

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Carga");

                    // Add headers
                    worksheet.Cell(1, 1).Value = "CODIGO BARRA";
                    worksheet.Cell(1, 2).Value = "DESC PRODUTO";
                    worksheet.Cell(1, 3).Value = "FILIAL";
                    worksheet.Cell(1, 4).Value = "PRODUTO";
                    worksheet.Cell(1, 5).Value = "COR PRODUTO";
                    worksheet.Cell(1, 6).Value = "DESC COR PRODUTO";
                    worksheet.Cell(1, 7).Value = "GRADE";
                    worksheet.Cell(1, 8).Value = "ESTOQUE";
                    worksheet.Cell(1, 9).Value = "CUSTO REPOSICAO";

                    // Add data
                    for (int i = 0; i < carga.Count; i++)
                    {
                        worksheet.Cell(i + 2, 1).Value = carga[i].CODIGO_BARRA;
                        worksheet.Cell(i + 2, 2).Value = carga[i].DESC_PRODUTO;
                        worksheet.Cell(i + 2, 3).Value = carga[i].FILIAL;
                        worksheet.Cell(i + 2, 4).Value = carga[i].PRODUTO;
                        worksheet.Cell(i + 2, 5).Value = carga[i].COR_PRODUTO;
                        worksheet.Cell(i + 2, 6).Value = carga[i].DESC_COR_PRODUTO;
                        worksheet.Cell(i + 2, 7).Value = carga[i].GRADE;
                        worksheet.Cell(i + 2, 8).Value = carga[i].ESTOQUE;
                        worksheet.Cell(i + 2, 9).Value = carga[i].CUSTO_REPOSICAO1;
                    }

                    // Adjust column widths
                    worksheet.Columns().AdjustToContents();

                    // Bold headers
                    worksheet.Row(1).Style.Font.Bold = true;

                    // Convert to byte array
                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        stream.Position = 0;

                        string fileName = $"Relatorio_Carga_{DateTime.Now:yyyyMMdd}.xlsx";
                        return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                    }
                }
            }
            catch (Exception ex)
            {
                TempData["Mensagem"] = $"Erro ao gerar o relatório: {ex.Message}\nStack Trace: {ex.StackTrace}";
                return RedirectToAction(nameof(GeraCargaProprias));
            }
        }
    }
}