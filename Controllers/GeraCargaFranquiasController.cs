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
                // Carregar os dados do banco primeiro
                var filiaisData = await _context.V_FILIAIS_ATIVAS_FRANQUIAS
                    .OrderBy(f => f.Filial)
                    .ToListAsync();

                // Formatar os dados no lado do cliente
                var filiais = filiaisData
                    .Select(f => new SelectListItem
                    {
                        Value = f.Cgc_Cpf,
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
            Console.WriteLine($"Executando ExecutarProcedure com Filial: {filial ?? "Nulo"}");
            var model = new FiliaisAtivasFModel
            {
                Filial = filial
            };

            if (string.IsNullOrEmpty(filial))
            {
                model.Mensagem = "Por favor, selecione uma filial.";
                await CarregarFiliais();
                return View("GeraCargaFranquias", model);
            }

            try
            {
                using var connection = new SqlConnection(_context.Database.GetConnectionString());
                await connection.OpenAsync();

                using var command = connection.CreateCommand();
                command.CommandTimeout = 1200000;
                command.CommandText = "GERA_CARGA_INVENT_FRANQUIAS @Filial";
                command.Parameters.Add(new SqlParameter
                {
                    ParameterName = "@Filial",
                    SqlDbType = SqlDbType.NVarChar,
                    Value = filial
                });

                await command.ExecuteNonQueryAsync();
                model.Mensagem = "Saldo Gerado com sucesso!";
                Console.WriteLine("Saldo Gerado com sucesso!");
            }
            catch(Exception ex)
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
        public async Task<IActionResult> GerarArquivoTxt(string filial)
        {
            Console.WriteLine($"Executando GerarArquivoTxt com Filial: {filial ?? "Nulo"}");
            var model = new GeraCargaFranquiasModel
            {
                Filial = filial
            };

            if (string.IsNullOrEmpty(filial))
            {
                model.Mensagem = "Por favor, selecione uma filial.";
                await CarregarFiliais();
                return View("GeraCargaFranquias", model);
            }

            try
            {
                // Executar a stored procedure para garantir dados atualizados
                using var connection = new SqlConnection(_context.Database.GetConnectionString());
                await connection.OpenAsync();

                using var command = connection.CreateCommand();
                command.CommandTimeout = 1200000;
                command.CommandText = "GERA_CARGA_INVENT_FRANQUIAS @Filial";
                command.Parameters.Add(new SqlParameter
                {
                    ParameterName = "@Filial",
                    SqlDbType = SqlDbType.NVarChar,
                    Value = filial
                });

                await command.ExecuteNonQueryAsync();
                Console.WriteLine("Stored procedure executada com sucesso em GerarArquivoTxt");

                // Consultar a tabela TABELA_CARGA_INV_FRANQUIAS
                var dados = await _context.TABELA_CARGA_INV_FRANQUIAS
                    .Select(f => new
                    {
                        f.Produto
                    })
                    .ToListAsync();

                Console.WriteLine($"Registros encontrados: {dados.Count}");

                // Definir caminho e nome do arquivo
                string fileName = $"Inventario_Franquias_{filial}_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
                string filePath = Path.Combine(Path.GetTempPath(), fileName);
                Console.WriteLine($"Gerando arquivo em: {filePath}");

                // Criar e escrever no arquivo TXT
                using (var writer = new StreamWriter(filePath))
                {
                    // Escrever cabeçalho
                    await writer.WriteLineAsync("Produto");

                    // Escrever dados
                    foreach (var item in dados)
                    {
                        string linha = $"{item.Produto}";
                        await writer.WriteLineAsync(linha);
                    }
                }

                // Retornar o arquivo para download
                var fileBytes = await System.IO.File.ReadAllBytesAsync(filePath);
                System.IO.File.Delete(filePath); // Remove o arquivo temporário após leitura

                // Carregar filiais para a view
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