using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using RelatoriosRosset.Models;
using System.Data;
using System.IO;
using System.Threading.Tasks;
using System.Collections.Generic;
using System;

namespace RelatoriosRosset.Controllers
{
    public class GeraCargaFranquiasController : Controller
    {
        private readonly ApplicationDbContext _context;

        public GeraCargaFranquiasController(ApplicationDbContext context)
        {
            _context = context;
        }

        public async Task<IActionResult> Index()
        {
            Console.WriteLine("Acessando Index...");
            return RedirectToAction("GeraCargaFranquias");
        }

        public async Task<IActionResult> GeraCargaFranquias()
        {
            Console.WriteLine("Acessando GeraCargaFranquias...");
            var model = new GeraCargaFranquiasModel();
            await CarregarFiliais();
            return View(model);
        }

        private async Task CarregarFiliais()
        {
            try
            {
                var filiaisData = await _context.V_FILIAIS_ATIVAS_FRANQUIAS
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

                ViewBag.V_FILIAIS_ATIVAS_FRANQUIAS = filiais;
            }
            catch (Exception ex)
            {
                ViewBag.V_FILIAIS_ATIVAS_FRANQUIAS = new List<SelectListItem>();
                TempData["Erro"] = $"Erro ao carregar filiais: {ex.Message}. StackTrace: {ex.StackTrace}";
                Console.WriteLine($"Erro ao carregar filiais: {ex.Message}\n{ex.StackTrace}");
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
            var model = new GeraCargaFranquiasModel
            {
                Filial = filial?.Trim()
            };

            if (string.IsNullOrEmpty(filial))
            {
                model.Mensagem = "Por favor, selecione uma filial.";
                await CarregarFiliais();
                return View("GeraCargaFranquias", model);
            }

            try
            {
                var filialExists = await _context.V_FILIAIS_ATIVAS_FRANQUIAS
                    .AnyAsync(f => f.Filial == filial);
                if (!filialExists)
                {
                    model.Mensagem = $"Filial '{filial}' não encontrada.";
                    await CarregarFiliais();
                    return View("GeraCargaFranquias", model);
                }

                using var connection = new SqlConnection(_context.Database.GetConnectionString());
                await connection.OpenAsync();
                //using var transaction = await connection.BeginTransactionAsync();

                using var command = connection.CreateCommand();
                //command.Transaction = transaction;
                command.CommandTimeout = 300000;
                command.CommandText = "GERA_CARGA_INVENT_FRANQUIAS @Filial";
                command.Parameters.Add(new SqlParameter
                {
                    ParameterName = "@Filial",
                    SqlDbType = SqlDbType.NVarChar,
                    Value = filial
                });

                await command.ExecuteNonQueryAsync();

                var dados = await _context.TABELA_CARGA_INV_FRANQUIAS.ToListAsync();
                Console.WriteLine($"Registros inseridos na tabela: {dados.Count}");
                foreach (var item in dados)
                {
                    Console.WriteLine($"Produto: {item.Produto}");
                }

                //await transaction.CommitAsync();
                model.Mensagem = "Saldo Gerado com sucesso!";
                Console.WriteLine("Saldo Gerado com sucesso!");
            }

            catch (Exception ex)
            {
                model.Mensagem = $"Erro ao executar a procedure: {ex.Message}. StackTrace: {ex.StackTrace}";
                Console.WriteLine($"Erro ao executar a procedure: {ex.Message}\n{ex.StackTrace}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner Exception: {ex.InnerException.Message}\n{ex.InnerException.StackTrace}");
                }
            }

            await CarregarFiliais();
            return View("GeraCargaFranquias", model);
        }

        [HttpPost]
        public async Task<IActionResult> GerarArquivoTxt()
        {
            Console.WriteLine("Executando GerarArquivoTxt...");
            var model = new GeraCargaFranquiasModel();

            try
            {
                // Consultar os dados da tabela TABELA_CARGA_INV_FRANQUIAS
                var dados = await _context.TABELA_CARGA_INV_FRANQUIAS
                    .Select(f => new { f.Produto })
                    .ToListAsync();
                Console.WriteLine($"Registros encontrados na tabela: {dados.Count}");
                foreach (var item in dados)
                {
                    Console.WriteLine($"Produto: {item.Produto}");
                }

                if (!dados.Any())
                {
                    model.Mensagem = "Nenhum dado encontrado na tabela para gerar o arquivo.";
                    await CarregarFiliais();
                    return View("GeraCargaFranquias", model);
                }

                // Gerar arquivo
                string fileName = $"Inventario_Franquias_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
                string filePath = Path.Combine(Path.GetTempPath(), fileName);
                Console.WriteLine($"Gerando arquivo em: {filePath}");

                using (var writer = new StreamWriter(filePath))
                {
                    await writer.WriteLineAsync("Produto");
                    foreach (var item in dados)
                    {
                        await writer.WriteLineAsync(item.Produto);
                    }
                }

                var fileBytes = await System.IO.File.ReadAllBytesAsync(filePath);
                System.IO.File.Delete(filePath);

                await CarregarFiliais();
                model.Mensagem = "Arquivo TXT gerado com sucesso!";
                return File(fileBytes, "text/plain", fileName);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao gerar TXT: {ex.Message}\n{ex.StackTrace}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner Exception: {ex.InnerException.Message}\n{ex.InnerException.StackTrace}");
                }

                await CarregarFiliais();
                model.Mensagem = $"Erro ao gerar arquivo TXT: {ex.Message}";
                return View("GeraCargaFranquias", model);
            }
        }
    }
}