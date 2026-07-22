using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using RelatoriosRosset.Models;

namespace RelatoriosRosset.Controllers
{
    public class RegistrosVendasController : Controller
    {
        private readonly ApplicationDbContext _context;

        public RegistrosVendasController(ApplicationDbContext context)
        {
            _context = context;
        }

        // LISTA REGISTRO 54
        public async Task<IActionResult> listaRegistro54(DateTime? dataInicio, DateTime? dataFim)
        {
            if (!dataInicio.HasValue || !dataFim.HasValue)
                return View(new List<Registro54Result>());

            string sql = @"
            SELECT TOP 10
            B.TIPO_REGISTRO,B.CPF_CGC,B.MODELO,B.SERIE_NF,B.NF,B.CODIGO_FISCAL_OPERACAO,B.TRIBUT_ICMS,
            B.ITEM_CFE,B.CODIGO_ITEM,B.CODIGO_BARRA,B.QTDE_ITEM,B.VALOR_ITEM,B.DESCONTO_ITEM,B.BASE_IMPOSTO,
            B.BASE_IMPOSTO_SUBST,B.VALOR_IPI,B.TAXA_IMPOSTO,B.VALOR_CONTABIL,B.VALOR_ICMS,B.FILIAL_CHAVE,
            B.NF_SAIDA_CHAVE,B.SERIE_NF_CHAVE,B.COD_MATRIZ_FISCAL,B.MATRIZ_FISCAL,
            CONVERT(VARCHAR(10),B.EMISSAO,121) as EMISSAO,
            B.DESCRICAO_ITEM,B.UNIDADE,B.RG_IE,B.CLASSIF_FISCAL,B.UF,B.PAIS,B.ITEM_IMPRESSAO,B.COD_FILIAL,
            B.NOTA_CANCELADA,B.SERIE_CFE_SAT,B.CHAVE_NFE 
            FROM VAL_EXPORTA_SPED_CFE_LF_REGISTRO_SAIDA_IMPOSTO  A 
            LEFT JOIN VAL_EXPORTA_SPED_CFE_LF_REGISTRO_SAIDA_IMPOSTO_ITEM_new B 
            ON A.NF=B.NF 
            AND A.COD_MATRIZ_FISCAL=B.COD_MATRIZ_FISCAL 
            AND A.SERIE_NF=B.SERIE_NF 
            AND A.serie_cfe_sat=B.serie_cfe_sat 
            WHERE A.COD_MATRIZ_FISCAL in (
            select distinct CODIGO_FILIAL 
            from LOJA_NOTA_FISCAL 
            where SERIE_NF in ('65','001') 
            and STATUS_NFE = '5' 
            AND EMISSAO BETWEEN {0} AND {1}
            )
            AND A.EMISSAO BETWEEN {0} AND {1}
            AND B.SERIE_NF IS NOT NULL
            ORDER BY A.FILIAL_CHAVE";

            string dtIni = dataInicio.Value.ToString("yyyyMMdd");
            string dtFim = dataFim.Value.ToString("yyyyMMdd");

            var dados = await _context.Registro54
                .FromSqlRaw(sql, dtIni, dtFim)
                .ToListAsync();

            //var dados = await _context.Registro54
            //.FromSqlRaw(sql, dtIni, dtFim)
            //.Take(10)
            //.ToListAsync();

            ViewBag.DataInicio = dataInicio;
            ViewBag.DataFim = dataFim;

            return View(dados);
            //return View("listaRegistro54", dados);
        }

        // EXPORTAR EXCEL DO REGISTRO 54
        public async Task<IActionResult> ExportarExcel(DateTime? dataInicio, DateTime? dataFim)
        {
            if (!dataInicio.HasValue || !dataFim.HasValue)
                return Content("Informe data inicial e final.");

            string sql = @"
                SELECT 
                B.TIPO_REGISTRO,B.CPF_CGC,B.MODELO,B.SERIE_NF,B.NF,B.CODIGO_FISCAL_OPERACAO,B.TRIBUT_ICMS,
                B.ITEM_CFE,B.CODIGO_ITEM,B.CODIGO_BARRA,B.QTDE_ITEM,B.VALOR_ITEM,B.DESCONTO_ITEM,B.BASE_IMPOSTO,
                B.BASE_IMPOSTO_SUBST,B.VALOR_IPI,B.TAXA_IMPOSTO,B.VALOR_CONTABIL,B.VALOR_ICMS,B.FILIAL_CHAVE,
                B.NF_SAIDA_CHAVE,B.SERIE_NF_CHAVE,B.COD_MATRIZ_FISCAL,B.MATRIZ_FISCAL,
                CONVERT(VARCHAR(10),B.EMISSAO,121) as EMISSAO,
                B.DESCRICAO_ITEM,B.UNIDADE,B.RG_IE,B.CLASSIF_FISCAL,B.UF,B.PAIS,B.ITEM_IMPRESSAO,B.COD_FILIAL,
                B.NOTA_CANCELADA,B.SERIE_CFE_SAT,B.CHAVE_NFE 
                FROM VAL_EXPORTA_SPED_CFE_LF_REGISTRO_SAIDA_IMPOSTO  A 
                LEFT JOIN VAL_EXPORTA_SPED_CFE_LF_REGISTRO_SAIDA_IMPOSTO_ITEM_new B 
                    ON A.NF=B.NF 
                    AND A.COD_MATRIZ_FISCAL=B.COD_MATRIZ_FISCAL 
                    AND A.SERIE_NF=B.SERIE_NF 
                    AND A.serie_cfe_sat=B.serie_cfe_sat 
                WHERE A.COD_MATRIZ_FISCAL in (
                    select distinct CODIGO_FILIAL 
                    from LOJA_NOTA_FISCAL 
                    where SERIE_NF in ('65','001') 
                    and STATUS_NFE = '5' 
                    AND EMISSAO BETWEEN {0} AND {1}
                )
                AND A.EMISSAO BETWEEN {0} AND {1}
                AND B.SERIE_NF IS NOT NULL
                ORDER BY A.FILIAL_CHAVE";

            string dtIni = dataInicio.Value.ToString("yyyyMMdd");
            string dtFim = dataFim.Value.ToString("yyyyMMdd");

            var dados = await _context.Registro54
                .FromSqlRaw(sql, dtIni, dtFim)
                .ToListAsync();

            if (!dados.Any())
                return Content("Nenhum registro encontrado no período informado.");

            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Registro54");

                // Cabeçalhos automáticos
                var propriedades = typeof(Registro54Result).GetProperties();
                for (int i = 0; i < propriedades.Length; i++)
                {
                    ws.Cell(1, i + 1).Value = propriedades[i].Name;
                    ws.Cell(1, i + 1).Style.Font.Bold = true;
                }

                // Dados
                for (int row = 0; row < dados.Count; row++)
                {
                    for (int col = 0; col < propriedades.Length; col++)
                    {
                        var valor = propriedades[col].GetValue(dados[row]);

                        // Valor tratado de forma segura para evitar erros de conversão ao definir a célula
                        var cell = ws.Cell(row + 2, col + 1);

                        if (valor == null)
                        {
                            cell.Value = string.Empty;
                            continue;
                        }

                        switch (valor)
                        {
                            case DateTime dt:
                                cell.Value = dt;
                                cell.Style.DateFormat.Format = "dd/MM/yyyy";
                                break;

                            // Se o valor é decimal
                            case decimal dec:
                                cell.Value = Convert.ToDouble(dec);
                                break;

                            // Se o SQL devolveu int mas o modelo espera decimal
                            case int i when propriedades[col].PropertyType == typeof(decimal):
                                cell.Value = Convert.ToDouble(i);
                                break;

                            // Se o SQL devolveu long mas o modelo espera decimal
                            case long l when propriedades[col].PropertyType == typeof(decimal):
                                cell.Value = Convert.ToDouble(l);
                                break;

                            // Outros números
                            case double d:
                                cell.Value = d;
                                break;

                            case float f:
                                cell.Value = Convert.ToDouble(f);
                                break;

                            case int i:
                                cell.Value = i;
                                break;

                            case long l:
                                cell.Value = l;
                                break;

                            default:
                                cell.Value = valor.ToString();
                                break;
                        }


                    }
                }


                ws.Columns().AdjustToContents();

                using var stream = new MemoryStream();
                workbook.SaveAs(stream);
                stream.Position = 0;

                string fileName = $"Registro54_{DateTime.Now:yyyyMMdd}.xlsx";

                return File(stream.ToArray(),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    fileName);
            }
        }
    }
}

