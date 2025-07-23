using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;

namespace RelatoriosRosset.Controllers
{
    public class NotasDevPController : Controller
    {
        private readonly ApplicationDbContext _context;

        public NotasDevPController(ApplicationDbContext context)
        {
            _context = context;
        }
        public async Task<IActionResult> NotasDevP(DateTime? dataInicio, DateTime? dataFim, string clienteVarejo, string cpfCgc)
        {
            var query = _context.NOTAS_DEVOLUCAO_PROPRIAS.AsQueryable();

            if (dataInicio.HasValue)
                query = query.Where(v => v.EMISSAO >= dataInicio.Value);

            if (dataFim.HasValue)
                query = query.Where(v => v.EMISSAO <= dataFim.Value);

            if (!string.IsNullOrEmpty(clienteVarejo))
                query = query.Where(v => v.CLIENTE_VAREJO.Contains(clienteVarejo));

            if (!string.IsNullOrEmpty(cpfCgc))
                query = query.Where(v => v.CPF_CGC.Contains(cpfCgc));

            var notasP = await query.OrderByDescending(v => v.EMISSAO).Take(10).ToListAsync();

            ViewBag.DataInicio = dataInicio;
            ViewBag.DataFim = dataFim;
            ViewBag.ClienteVarejo = clienteVarejo;
            ViewBag.CpfCgc = cpfCgc;

            return View(notasP);
        }

        public async Task<IActionResult> ExportarExcel(DateTime? dataInicio, DateTime? dataFim, string clienteVarejo, string cpfCgc)

        {
            // Configurar a licença do EPPlus para uso não comercial diretamente no código
            //ExcelPackage.LicenseContext = new LicenseInfo { IsCommercial = false };
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var query = _context.NOTAS_DEVOLUCAO_FRANQUIAS.AsQueryable();

            if (dataInicio.HasValue)
                query = query.Where(v => v.EMISSAO >= dataInicio.Value);

            if (dataFim.HasValue)
                query = query.Where(v => v.EMISSAO <= dataFim.Value);

            if (!string.IsNullOrEmpty(clienteVarejo))
                query = query.Where(v => v.CLIENTE_VAREJO.Contains(clienteVarejo));

            if (!string.IsNullOrEmpty(cpfCgc))
                query = query.Where(v => v.CPF_CGC.Contains(cpfCgc));

            var notasF = await query.OrderBy(v => v.EMISSAO).ToListAsync();

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Notas");

                // Adicionar cabeçalhos
                worksheet.Cells[1, 1].Value = "Filial";
                worksheet.Cells[1, 2].Value = "Número NF";
                worksheet.Cells[1, 3].Value = "Emissão";
                worksheet.Cells[1, 4].Value = "Valor Total";
                worksheet.Cells[1, 5].Value = "Natureza Operação";
                worksheet.Cells[1, 6].Value = "CPF/CNPJ";
                worksheet.Cells[1, 7].Value = "Cliente Varejo";
                worksheet.Cells[1, 8].Value = "Chave NFE";

                // Adicionar dados
                for (int i = 0; i < notasF.Count; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = notasF[i].FILIAL;
                    worksheet.Cells[i + 2, 2].Value = notasF[i].NF_NUMERO;
                    worksheet.Cells[i + 2, 3].Value = notasF[i].EMISSAO.ToString("dd/MM/yyyy");
                    worksheet.Cells[i + 2, 4].Value = notasF[i].VALOR_TOTAL;
                    worksheet.Cells[i + 2, 5].Value = notasF[i].NATUREZA_OPERACAO_CODIGO;
                    worksheet.Cells[i + 2, 6].Value = notasF[i].CPF_CGC;
                    worksheet.Cells[i + 2, 7].Value = notasF[i].CLIENTE_VAREJO;
                    worksheet.Cells[i + 2, 8].Value = notasF[i].CHAVE_NFE;
                }

                // Ajustar formato das colunas
                worksheet.Cells[1, 1, notasF.Count + 1, 3].AutoFitColumns();

                // Configurar cabeçalhos como negrito
                worksheet.Row(1).Style.Font.Bold = true;

                // Converter o pacote para um array de bytes
                var stream = new MemoryStream();
                package.SaveAs(stream);
                stream.Position = 0;

                // Nome do arquivo
                string fileName = $"Relatorio_Vendas_{DateTime.Now:yyyyMMdd}.xlsx";

                // Retornar o arquivo Excel
                return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
        }
    }
}
