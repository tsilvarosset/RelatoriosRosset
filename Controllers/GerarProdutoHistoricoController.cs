using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using RelatoriosRosset.Models;
using System.Data;

namespace RelatoriosRosset.Controllers
{
    public class GerarProdutoHistoricoController : Controller
    {
        private readonly ApplicationDbContext _context;

        public GerarProdutoHistoricoController(ApplicationDbContext context)
        {
            _context = context;
        }

        public async Task<IActionResult> Index()
        {
            await GerarProdutoHistorico();
            return View(new GerarProdutoHistoricoModel());
        }

        private async Task GerarProdutoHistorico()
        {
            try
            {
                var filiais = await _context.V_FILIAIS_ATIVAS_PROPRIAS
                    .Select(f => new SelectListItem
                    {
                        Text = f.Filial
                    })
                    .OrderBy(f => f.Text)
                    .ToListAsync();

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
                Console.WriteLine($"Erro ao carregar filiais: {ex.Message}\n{ex.StackTrace}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner Exception: {ex.InnerException.Message}\n{ex.InnerException.StackTrace}");
                }
            }
        }

        [HttpPost]
        public async Task<IActionResult> ExecutarProcedure(DateTime dataSaldo, string filial)
        {
            var model = new GerarProdutoHistoricoModel
            {
                DATA_SALDO = dataSaldo,
                FILIAL = filial
            };

            try
            {
                using var connection = new SqlConnection(_context.Database.GetConnectionString());
                await connection.OpenAsync();

                using var command = connection.CreateCommand();
                command.CommandTimeout = 120;
                command.CommandText = "LX_GERA_HISTORICO_ESTOQUE_PA_LOJA @DataSaldo, @Filial";
                command.Parameters.Add(new SqlParameter
                {
                    ParameterName = "@DataSaldo",
                    SqlDbType = SqlDbType.Date,
                    Value = dataSaldo
                });
                command.Parameters.Add(new SqlParameter
                {
                    ParameterName = "@Filial",
                    SqlDbType = SqlDbType.NVarChar,
                    Value = string.IsNullOrEmpty(filial) ? (object)DBNull.Value : filial
                });

                await command.ExecuteNonQueryAsync();
                model.Mensagem = "Saldo Gerado com sucesso!";
                Console.WriteLine("Saldo Gerado com sucesso!");
            }
            catch (Exception ex)
            {
                model.Mensagem = $"Erro ao executar a procedure: {ex.Message}. StackTrace: {ex.StackTrace}";
                Console.WriteLine($"Erro ao executar a procedure: {ex.Message}\n{ex.StackTrace}");
            }

            await GerarProdutoHistorico();
            return View("Index", model); // Ajuste para "Index" se a view foi renomeada
        }
    }
}