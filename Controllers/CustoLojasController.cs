using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using RelatoriosRosset.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace RelatoriosRosset.Controllers
{
    public class CustoLojasController : Controller
    {
        private readonly ApplicationDbContext _context;

        public CustoLojasController(ApplicationDbContext context)
        {
            _context = context;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        // GET: Exibe a view com o formulário
        public async Task<IActionResult> Index()
        {
            var model = new CustoLojasViewModel();
            await CarregarFiliais(); // Carrega as filiais do banco
            //Console.WriteLine($"Filiais no Index: {((List<SelectListItem>)ViewBag.FILIAIS).Count}");
            return View(model);

        }
        private async Task CarregarFiliais()
        {
            try
            {
                var filiais = await _context.FILIAIS
                    .Select(f => new SelectListItem
                    {
                        //Value = f.COD_FILIAL,
                        Text = f.FILIAL
                    })
                    .OrderBy(f => f.Text)
                    .ToListAsync();

                //Console.WriteLine($"Filiais carregadas: {filiais.Count}");
                foreach (var filial in filiais)
                {
                    Console.WriteLine($"Filial: {filial.Text}, Código: {filial.Value}");
                }

                ViewBag.FILIAIS = filiais;
            }
            catch (Exception ex)
            {
                ViewBag.FILIAIS = new List<SelectListItem>();
                TempData["Erro"] = $"Erro ao carregar filiais: {ex.Message}. StackTrace: {ex.StackTrace}";
                Console.WriteLine($"Erro ao carregar filiais: {ex.Message}\n{ex.StackTrace}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner Exception: {ex.InnerException.Message}\n{ex.InnerException.StackTrace}");
                }
            }
        }

        [HttpPost]
        public async Task<IActionResult> ExecutarProcedure(DateTime dataInicial, DateTime dataFinal, string filial)
        {
            var model = new CustoLojasViewModel
            {
                DataInicial = dataInicial,
                DataFinal = dataFinal,
                Filial = filial
            };

            Console.WriteLine($"Executando procedure com DataInicial: {dataInicial}, DataFinal: {dataFinal}, Filial: {filial ?? "Nulo"}");

            try
            {
                using var connection = new SqlConnection(_context.Database.GetConnectionString());

                await connection.OpenAsync();

                using var command = connection.CreateCommand();
                command.CommandTimeout = 120;
                command.CommandText = "EXEC GERA_CUSTO_LOJAS_EAN @DataInicial, @DataFinal, @Filial";
                command.Parameters.Add(new SqlParameter
                {
                    ParameterName = "@DataInicial",
                    SqlDbType = SqlDbType.Date,
                    Value = dataInicial
                });
                command.Parameters.Add(new SqlParameter
                {
                    ParameterName = "@DataFinal",
                    SqlDbType = SqlDbType.Date,
                    Value = dataFinal
                });
                command.Parameters.Add(new SqlParameter
                {
                    ParameterName = "@Filial",
                    SqlDbType = SqlDbType.NVarChar,
                    Value = string.IsNullOrEmpty(filial) ? (object)DBNull.Value : filial
                });

                await command.ExecuteNonQueryAsync();
                model.Mensagem = "Procedure executada com sucesso!";
                Console.WriteLine("Procedure executada com sucesso.");
            }
            catch (Exception ex)
            {
                model.Mensagem = $"Erro ao executar a procedure: {ex.Message}. StackTrace: {ex.StackTrace}";
                Console.WriteLine($"Erro ao executar a procedure: {ex.Message}\n{ex.StackTrace}");
            }

            // Carrega as filiais
            try
            {
                using var serviceScope = HttpContext.RequestServices.CreateScope();
                using var newContext = serviceScope.ServiceProvider.GetRequiredService<ApplicationDbContext>();
                var filiais = await newContext.FILIAIS
                    .Select(f => new SelectListItem
                    {
                        //Value = f.COD_FILIAL,
                        Text = f.FILIAL
                    })
                    .OrderBy(f => f.Text)
                    .ToListAsync();

                Console.WriteLine($"Filiais carregadas: {filiais.Count}");
                foreach (var f in filiais)
                {
                    Console.WriteLine($"Filial: {f.Text}, Código: {f.Value}");
                }

                ViewBag.FILIAIS = filiais;
            }
            catch (Exception ex)
            {
                ViewBag.FILIAIS = new List<SelectListItem>();
                TempData["Erro"] = $"Erro ao carregar filiais: {ex.Message}. StackTrace: {ex.StackTrace}";
                Console.WriteLine($"Erro ao carregar filiais: {ex.Message}\n{ex.StackTrace}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner Exception: {ex.InnerException.Message}\n{ex.InnerException.StackTrace}");
                }
            }

            return View("Index", model);
        }

        // GET: Gera o relatório em Excel
        [HttpGet]
        public async Task<IActionResult> GerarRelatorio()
        {
            try
            {
                var resultados = new List<CustoLojaResults>();
                using var connection = _context.Database.GetDbConnection();
                await connection.OpenAsync();

                using var command = connection.CreateCommand();
                command.CommandText = @"
                    SELECT 
                        TRY_CAST(ITEM AS INT) AS Item,
                        DESC_ITEM_COMPOSICAO AS DescItemComposicao,
                        TRY_CAST(ITEM_COMPOSICAO AS INT) AS ItemComposicao,
                        PRODUTO AS Produto,
                        COR_PRODUTO AS CorProduto,
                        CODIGO_BARRA AS CodigoBarra,
                        QTDE AS Qtde,
                        VALOR_CUSTO,
                        DOC AS Doc,
                        ROMANEIO_PRODUTO AS RomaneioProduto,
                        SERIE_NF AS SerieNf,
                        CFOP AS Cfop,
                        DESCRICAO_CFOP AS DescricaoCfop,
                        FILIAL AS Filial,
                        RATEIO_FILIAL AS RateioFilial,
                        DATA_MOV AS DataMov,
                        TOTAL_QTDE AS TotalQtde,
                        VALOR_TOTAL AS ValorTotal,
                        VALOR_IMPOSTO_DESTACAR AS ValorImpostoDestacar,
                        VALOR_LIQ AS ValorLiq,
                        VALOR_PRODUCAO AS ValorProducao,
                        VALOR_BRUTO AS ValorBruto,
                        VL_ICMS_OP AS VlIcmsOp
                    FROM TABELA_CUSTO_EAN
                    WHERE TRY_CAST(ITEM AS INT) IS NOT NULL
                        AND TRY_CAST(ITEM_COMPOSICAO AS INT) IS NOT NULL";

                using var reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    resultados.Add(new CustoLojaResults
                    {
                        Item = reader.IsDBNull(0) ? null : reader.GetInt32(0),
                        DescItemComposicao = reader.IsDBNull(1) ? null : reader.GetString(1),
                        ItemComposicao = reader.IsDBNull(2) ? null : reader.GetInt32(2),
                        Produto = reader.IsDBNull(3) ? null : reader.GetString(3),
                        CorProduto = reader.IsDBNull(4) ? null : reader.GetString(4),
                        CodigoBarra = reader.IsDBNull(5) ? null : reader.GetString(5),
                        Qtde = reader.IsDBNull(6) ? 0 : reader.GetInt32(6),
                        ValorCusto = reader.IsDBNull(7) ? null : reader.GetDecimal(7),
                        Doc = reader.IsDBNull(8) ? null : reader.GetString(8),
                        RomaneioProduto = reader.IsDBNull(9) ? null : reader.GetString(9),
                        SerieNf = reader.IsDBNull(10) ? null : reader.GetString(10),
                        Cfop = reader.IsDBNull(11) ? null : reader.GetString(11),
                        DescricaoCfop = reader.IsDBNull(12) ? null : reader.GetString(12),
                        Filial = reader.IsDBNull(13) ? null : reader.GetString(13),
                        RateioFilial = reader.IsDBNull(14) ? null : reader.GetString(14),
                        DataMov = reader.IsDBNull(15) ? DateTime.MinValue : reader.GetDateTime(15),
                        TotalQtde = reader.IsDBNull(16) ? 0 : reader.GetInt32(16),
                        ValorTotal = reader.IsDBNull(17) ? null : reader.GetDecimal(17),
                        ValorImpostoDestacar = reader.IsDBNull(18) ? null : reader.GetDecimal(18),
                        ValorLiq = reader.IsDBNull(19) ? null : reader.GetDecimal(19),
                        ValorProducao = reader.IsDBNull(20) ? null : reader.GetDecimal(20),
                        ValorBruto = reader.IsDBNull(21) ? null : reader.GetDecimal(21),
                        VlIcmsOp = reader.IsDBNull(22) ? null : reader.GetDecimal(22)
                    });
                }

                if (!resultados.Any())
                {
                    var model = new CustoLojasViewModel
                    {
                        Mensagem = "Nenhum registro encontrado na TABELA_CUSTO_EAN."
                    };
                    await CarregarFiliais(); // Recarrega as filiais
                    return View("Index", model);
                }

                using var package = new ExcelPackage();
                var worksheet = package.Workbook.Worksheets.Add("CustoLojas");

                // Headers
                worksheet.Cells[1, 1].Value = "Item";
                worksheet.Cells[1, 2].Value = "Descrição Item Composição";
                worksheet.Cells[1, 3].Value = "Item Composição";
                worksheet.Cells[1, 4].Value = "Produto";
                worksheet.Cells[1, 5].Value = "Cor Produto";
                worksheet.Cells[1, 6].Value = "Código de Barra";
                worksheet.Cells[1, 7].Value = "Quantidade";
                worksheet.Cells[1, 8].Value = "Valor Custo";
                worksheet.Cells[1, 9].Value = "Documento";
                worksheet.Cells[1, 10].Value = "Romaneio Produto";
                worksheet.Cells[1, 11].Value = "Série NF";
                worksheet.Cells[1, 12].Value = "CFOP";
                worksheet.Cells[1, 13].Value = "Descrição CFOP";
                worksheet.Cells[1, 14].Value = "Filial";
                worksheet.Cells[1, 15].Value = "Rateio Filial";
                worksheet.Cells[1, 16].Value = "Data Movimento";
                worksheet.Cells[1, 17].Value = "Total Quantidade";
                worksheet.Cells[1, 18].Value = "Valor Total";
                worksheet.Cells[1, 19].Value = "Valor Imposto Destacar";
                worksheet.Cells[1, 20].Value = "Valor Líquido";
                worksheet.Cells[1, 21].Value = "Valor Produção";
                worksheet.Cells[1, 22].Value = "Valor Bruto";
                worksheet.Cells[1, 23].Value = "Valor ICMS Operação";

                // Data
                for (int i = 0; i < resultados.Count; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = resultados[i].Item;
                    worksheet.Cells[i + 2, 2].Value = resultados[i].DescItemComposicao;
                    worksheet.Cells[i + 2, 3].Value = resultados[i].ItemComposicao;
                    worksheet.Cells[i + 2, 4].Value = resultados[i].Produto;
                    worksheet.Cells[i + 2, 5].Value = resultados[i].CorProduto;
                    worksheet.Cells[i + 2, 6].Value = resultados[i].CodigoBarra;
                    worksheet.Cells[i + 2, 7].Value = resultados[i].Qtde;
                    worksheet.Cells[i + 2, 8].Value = resultados[i].ValorCusto; // Nullable
                    worksheet.Cells[i + 2, 9].Value = resultados[i].Doc;
                    worksheet.Cells[i + 2, 10].Value = resultados[i].RomaneioProduto;
                    worksheet.Cells[i + 2, 11].Value = resultados[i].SerieNf;
                    worksheet.Cells[i + 2, 12].Value = resultados[i].Cfop;
                    worksheet.Cells[i + 2, 13].Value = resultados[i].DescricaoCfop;
                    worksheet.Cells[i + 2, 14].Value = resultados[i].Filial;
                    worksheet.Cells[i + 2, 15].Value = resultados[i].RateioFilial;
                    worksheet.Cells[i + 2, 16].Value = resultados[i].DataMov == DateTime.MinValue ? "" : resultados[i].DataMov.ToString("dd/MM/yyyy");
                    worksheet.Cells[i + 2, 17].Value = resultados[i].TotalQtde;
                    worksheet.Cells[i + 2, 18].Value = resultados[i].ValorTotal; // Nullable
                    worksheet.Cells[i + 2, 19].Value = resultados[i].ValorImpostoDestacar; // Nullable
                    worksheet.Cells[i + 2, 20].Value = resultados[i].ValorLiq; // Nullable
                    worksheet.Cells[i + 2, 21].Value = resultados[i].ValorProducao; // Nullable
                    worksheet.Cells[i + 2, 22].Value = resultados[i].ValorBruto; // Nullable
                    worksheet.Cells[i + 2, 23].Value = resultados[i].VlIcmsOp; // Nullable
                }

                // Apply number formatting
                worksheet.Cells[2, 1, resultados.Count + 1, 1].Style.Numberformat.Format = "0"; // Item
                worksheet.Cells[2, 3, resultados.Count + 1, 3].Style.Numberformat.Format = "0"; // ItemComposicao
                worksheet.Cells[2, 7, resultados.Count + 1, 7].Style.Numberformat.Format = "0"; // Qtde
                worksheet.Cells[2, 17, resultados.Count + 1, 17].Style.Numberformat.Format = "0"; // TotalQtde
                worksheet.Cells[2, 8, resultados.Count + 1, 8].Style.Numberformat.Format = "#,##0.00"; // ValorCusto
                worksheet.Cells[2, 18, resultados.Count + 1, 18].Style.Numberformat.Format = "#,##0.00"; // ValorTotal
                worksheet.Cells[2, 19, resultados.Count + 1, 19].Style.Numberformat.Format = "#,##0.00"; // ValorImpostoDestacar
                worksheet.Cells[2, 20, resultados.Count + 1, 20].Style.Numberformat.Format = "#,##0.00"; // ValorLiq
                worksheet.Cells[2, 21, resultados.Count + 1, 21].Style.Numberformat.Format = "#,##0.00"; // ValorProducao
                worksheet.Cells[2, 22, resultados.Count + 1, 22].Style.Numberformat.Format = "#,##0.00"; // ValorBruto
                worksheet.Cells[2, 23, resultados.Count + 1, 23].Style.Numberformat.Format = "#,##0.00"; // VlIcmsOp

                worksheet.Cells.AutoFitColumns();

                var stream = new MemoryStream();
                package.SaveAs(stream);
                stream.Position = 0;

                return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "CustoLojas.xlsx");
            }
            catch (Exception ex)
            {
                var model = new CustoLojasViewModel
                {
                    Mensagem = $"Erro ao gerar o relatório: {ex.Message}. StackTrace: {ex.StackTrace}"
                };
                await CarregarFiliais(); // Recarrega as filiais
                return View("Index", model);
            }
        }
    }
}