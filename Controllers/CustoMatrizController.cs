using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using RelatoriosRosset.Models;
using System.Data;

namespace RelatoriosRosset.Controllers
{
    public class CustoMatrizController : Controller
    {
        private readonly ApplicationDbContext _context;

        public CustoMatrizController(ApplicationDbContext context)
        {
            _context = context;
        }

        public async Task<IActionResult> CustoMatriz()
        {
            var model = new CustoViewMatrizModel();
            
            return View(model);

        }

        [HttpPost]
        public async Task<IActionResult> ExecutarProcedure(DateTime dataInicial, DateTime dataFinal, string filial)
        {
            var model = new CustoViewMatrizModel
            {
                DataInicial = dataInicial,
                DataFinal = dataFinal
            };

            Console.WriteLine($"Executando procedure com DataInicial: {dataInicial}, DataFinal: {dataFinal} ");

            try
            {
                using var connection = new SqlConnection(_context.Database.GetConnectionString());

                await connection.OpenAsync();

                using var command = connection.CreateCommand();
                command.CommandTimeout = 120000000;
                command.CommandText = "EXEC GERA_CUSTO_LOJAS_EAN_MATRIZ @DataInicial, @DataFinal ";
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
                
                await command.ExecuteNonQueryAsync();
                model.Mensagem = "Saldo Gerado com sucesso!";
                Console.WriteLine("Saldo Gerado com sucesso!");
            }
            catch (Exception ex)
            {
                model.Mensagem = $"Erro ao executar a procedure: {ex.Message}. StackTrace: {ex.StackTrace}";
                Console.WriteLine($"Erro ao executar a procedure: {ex.Message}\n{ex.StackTrace}");
            }

            return View("CustoMatriz", model);
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
                        GRADE AS Grade,
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
                        VALOR_BRUTO AS ValorBrut
                    FROM TABELA_CUSTO_EAN_MATRIZ
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
                        Grade = reader.IsDBNull(5) ? null : reader.GetString(5),
                        CodigoBarra = reader.IsDBNull(6) ? null : reader.GetString(6),
                        Qtde = reader.IsDBNull(7) ? 0 : reader.GetInt32(7),
                        ValorCusto = reader.IsDBNull(8) ? null : reader.GetDecimal(8),
                        Doc = reader.IsDBNull(9) ? null : reader.GetString(9),
                        RomaneioProduto = reader.IsDBNull(10) ? null : reader.GetString(10),
                        SerieNf = reader.IsDBNull(11) ? null : reader.GetString(11),
                        Cfop = reader.IsDBNull(12) ? null : reader.GetString(12),
                        DescricaoCfop = reader.IsDBNull(13) ? null : reader.GetString(13),
                        Filial = reader.IsDBNull(14) ? null : reader.GetString(14),
                        RateioFilial = reader.IsDBNull(15) ? null : reader.GetString(15),
                        DataMov = reader.IsDBNull(16) ? DateTime.MinValue : reader.GetDateTime(16),
                        TotalQtde = reader.IsDBNull(17) ? 0 : reader.GetInt32(17),
                        ValorTotal = reader.IsDBNull(18) ? null : reader.GetDecimal(18),
                        ValorImpostoDestacar = reader.IsDBNull(19) ? null : reader.GetDecimal(19),
                        ValorLiq = reader.IsDBNull(20) ? null : reader.GetDecimal(20),
                        ValorProducao = reader.IsDBNull(21) ? null : reader.GetDecimal(21),
                        ValorBruto = reader.IsDBNull(22) ? null : reader.GetDecimal(22),
                    });
                }

                if (!resultados.Any())
                {
                    var model = new CustoLojasViewModel
                    {
                        Mensagem = "Nenhum registro encontrado na TABELA_CUSTO_EAN."
                    };
                    
                    return View("CustoMatriz", model);
                }

                using var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("CustoLojas");

                // Headers
                worksheet.Cell(1, 1).Value = "Item";
                worksheet.Cell(1, 2).Value = "Descrição Item Composição";
                worksheet.Cell(1, 3).Value = "Item Composição";
                worksheet.Cell(1, 4).Value = "Produto";
                worksheet.Cell(1, 5).Value = "Cor Produto";
                worksheet.Cell(1, 6).Value = "Grade";
                worksheet.Cell(1, 7).Value = "Código de Barra";
                worksheet.Cell(1, 8).Value = "Quantidade";
                worksheet.Cell(1, 9).Value = "Valor Custo";
                worksheet.Cell(1, 10).Value = "Documento";
                worksheet.Cell(1, 11).Value = "Romaneio Produto";
                worksheet.Cell(1, 12).Value = "Série NF";
                worksheet.Cell(1, 13).Value = "CFOP";
                worksheet.Cell(1, 14).Value = "Descrição CFOP";
                worksheet.Cell(1, 15).Value = "Filial";
                worksheet.Cell(1, 16).Value = "Rateio Filial";
                worksheet.Cell(1, 17).Value = "Data Movimento";
                worksheet.Cell(1, 18).Value = "Total Quantidade";
                worksheet.Cell(1, 19).Value = "Valor Total";
                worksheet.Cell(1, 20).Value = "Valor Imposto Destacar";
                worksheet.Cell(1, 21).Value = "Valor Líquido";
                worksheet.Cell(1, 22).Value = "Valor Produção";
                worksheet.Cell(1, 23).Value = "Valor Bruto";

                // Data
                for (int i = 0; i < resultados.Count; i++)
                {
                    worksheet.Cell(i + 2, 1).Value = resultados[i].Item;
                    worksheet.Cell(i + 2, 2).Value = resultados[i].DescItemComposicao;
                    worksheet.Cell(i + 2, 3).Value = resultados[i].ItemComposicao;
                    worksheet.Cell(i + 2, 4).Value = resultados[i].Produto;
                    worksheet.Cell(i + 2, 5).Value = resultados[i].CorProduto;
                    worksheet.Cell(i + 2, 6).Value = resultados[i].Grade;
                    worksheet.Cell(i + 2, 7).Value = resultados[i].CodigoBarra;
                    worksheet.Cell(i + 2, 8).Value = resultados[i].Qtde;
                    worksheet.Cell(i + 2, 9).Value = resultados[i].ValorCusto; // Nullable
                    worksheet.Cell(i + 2, 10).Value = resultados[i].Doc;
                    worksheet.Cell(i + 2, 11).Value = resultados[i].RomaneioProduto;
                    worksheet.Cell(i + 2, 12).Value = resultados[i].SerieNf;
                    worksheet.Cell(i + 2, 13).Value = resultados[i].Cfop;
                    worksheet.Cell(i + 2, 14).Value = resultados[i].DescricaoCfop;
                    worksheet.Cell(i + 2, 15).Value = resultados[i].Filial;
                    worksheet.Cell(i + 2, 16).Value = resultados[i].RateioFilial;
                    worksheet.Cell(i + 2, 17).Value = resultados[i].DataMov == DateTime.MinValue ? "" : resultados[i].DataMov.ToString("dd/MM/yyyy");
                    worksheet.Cell(i + 2, 18).Value = resultados[i].TotalQtde;
                    worksheet.Cell(i + 2, 19).Value = resultados[i].ValorTotal; // Nullable
                    worksheet.Cell(i + 2, 20).Value = resultados[i].ValorImpostoDestacar; // Nullable
                    worksheet.Cell(i + 2, 21).Value = resultados[i].ValorLiq; // Nullable
                    worksheet.Cell(i + 2, 22).Value = resultados[i].ValorProducao; // Nullable
                    worksheet.Cell(i + 2, 23).Value = resultados[i].ValorBruto; // Nullable
                }

                // Apply number formatting
                worksheet.Column(1).Style.NumberFormat.Format = "0"; // Item
                worksheet.Column(3).Style.NumberFormat.Format = "0"; // ItemComposicao
                worksheet.Column(8).Style.NumberFormat.Format = "0"; // Qtde
                worksheet.Column(18).Style.NumberFormat.Format = "0"; // TotalQtde
                worksheet.Column(9).Style.NumberFormat.Format = "#,##0.00"; // ValorCusto
                worksheet.Column(19).Style.NumberFormat.Format = "#,##0.00"; // ValorTotal
                worksheet.Column(20).Style.NumberFormat.Format = "#,##0.00"; // ValorImpostoDestacar
                worksheet.Column(21).Style.NumberFormat.Format = "#,##0.00"; // ValorLiq
                worksheet.Column(22).Style.NumberFormat.Format = "#,##0.00"; // ValorProducao
                worksheet.Column(23).Style.NumberFormat.Format = "#,##0.00"; // ValorBruto
                worksheet.Column(24).Style.NumberFormat.Format = "#,##0.00"; // VlIcmsOp

                // Ajustar formato das colunas
                worksheet.Columns(1, 23).AdjustToContents();

                // Configurar cabeçalhos como negrito
                worksheet.Row(1).Style.Font.Bold = true;

                // Converter o workbook para um array de bytes
                using var stream = new MemoryStream();
                workbook.SaveAs(stream);
                stream.Position = 0;

                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "CustoLojas.xlsx");
            }
            catch (Exception ex)
            {
                var model = new CustoLojasViewModel
                {
                    Mensagem = $"Erro ao gerar o relatório: {ex.Message}. StackTrace: {ex.StackTrace}"
                };
               
                return View("CustoMatriz", model);
            }
        }
    }
}
