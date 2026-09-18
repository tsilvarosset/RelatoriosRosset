using Microsoft.AspNetCore.Mvc;
using RelatoriosRosset.Models;

// A action e o model têm o mesmo nome (Manuais). O alias abaixo evita o erro
// CS0118 ("'Manuais' é um método, mas é usado como um tipo") dentro do controller.
using ManuaisModel = RelatoriosRosset.Models.Manuais;

namespace RelatoriosRosset.Controllers
{
    /// <summary>
    /// Página "Manuais Linx": navega pelas pastas configuradas em appsettings.json
    /// (seção "Manuais") e entrega os PDFs para download.
    ///
    /// Estrutura atendida:
    ///   \\rosset39b\Tatiana\Valisere\Manuais\
    ///       VISUAL LINX\PDFs\*.pdf
    ///       Linx POS\PDFs\*.pdf
    ///       (outras pastas da Tatiana ficam de fora, ver "Frentes")
    ///
    /// Configuração:
    ///   Manuais:Pasta            -> pasta raiz (local ou UNC)
    ///   Manuais:Frentes          -> lista de pastas que devem aparecer. Vazia = todas.
    ///   Manuais:SubpastaArquivos -> subpasta com os PDFs que fica invisível (padrão "PDFs")
    ///
    /// Telas:
    ///   /Manuais/Manuais            -> view Views/Manuais/Manuais.cshtml
    ///   /Manuais                    -> redireciona para /Manuais/Manuais
    ///   /Manuais/Download?arquivo=  -> baixa o PDF
    /// </summary>
    public class ManuaisController : Controller
    {
        private readonly IConfiguration _configuration;
        private readonly ILogger<ManuaisController> _logger;

        public ManuaisController(IConfiguration configuration, ILogger<ManuaisController> logger)
        {
            _configuration = configuration;
            _logger = logger;
        }

        /// <summary>Pasta raiz dos manuais (local ou UNC de rede).</summary>
        private string PastaRaiz => _configuration["Manuais:Pasta"] ?? string.Empty;

        /// <summary>
        /// Pastas que devem aparecer na tela inicial (ex.: "VISUAL LINX", "Linx POS").
        /// Qualquer outra pasta da raiz é ignorada. Lista vazia = mostra todas.
        /// </summary>
        private string[] Frentes =>
            (_configuration.GetSection("Manuais:Frentes").Get<string[]>() ?? Array.Empty<string>())
            .Where(nome => !string.IsNullOrWhiteSpace(nome))
            .Select(nome => nome.Trim())
            .ToArray();

        /// <summary>
        /// Nome da subpasta que guarda os PDFs dentro de cada frente e que deve ficar
        /// invisível na navegação. Padrão "PDFs". Deixe em branco para desativar.
        /// </summary>
        private string SubpastaArquivos => _configuration["Manuais:SubpastaArquivos"] ?? "PDFs";

        // GET: /Manuais  -> mantém a rota padrão funcionando
        public IActionResult Index() => RedirectToAction(nameof(Manuais));

        // GET: /Manuais/Manuais
        // GET: /Manuais/Manuais?pasta=VISUAL LINX
        public IActionResult Manuais(string? pasta)
        {
            var model = new ManuaisModel
            {
                CaminhoAtual = NormalizarRelativo(pasta)
            };

            if (string.IsNullOrWhiteSpace(PastaRaiz))
            {
                ViewBag.Erro = "A pasta dos manuais não está configurada. " +
                               "Informe o caminho em appsettings.json na chave \"Manuais:Pasta\".";
                return View(model);
            }

            if (!Directory.Exists(PastaRaiz))
            {
                ViewBag.Erro = $"A pasta dos manuais não foi encontrada ou está inacessível: {PastaRaiz}";
                return View(model);
            }

            // Só permite navegar dentro das frentes configuradas
            if (!CaminhoPermitido(model.CaminhoAtual))
                return NotFound();

            var caminhoCompleto = ResolverCaminho(model.CaminhoAtual);
            if (caminhoCompleto == null)
                return BadRequest();

            if (!Directory.Exists(caminhoCompleto))
                return NotFound();

            model.Trilha = MontarTrilha(model.CaminhoAtual);
            model.CaminhoPai = CaminhoPai(model.CaminhoAtual);

            var naRaiz = string.IsNullOrEmpty(model.CaminhoAtual);
            var frentes = Frentes;

            try
            {
                var subpastas = Directory
                    .EnumerateDirectories(caminhoCompleto)
                    .Select(caminho => new DirectoryInfo(caminho))
                    .Where(dir => (dir.Attributes & FileAttributes.Hidden) == 0)
                    .ToList();

                // Na raiz, mostra apenas as pastas listadas em Manuais:Frentes
                if (naRaiz && frentes.Length > 0)
                {
                    subpastas = subpastas
                        .Where(dir => frentes.Any(frente =>
                            string.Equals(frente, dir.Name, StringComparison.OrdinalIgnoreCase)))
                        .ToList();
                }

                // Subpasta "PDFs" (ou o nome configurado): não aparece na navegação,
                // o conteúdo dela é mostrado junto com o da pasta atual.
                var pastaTransparente = string.IsNullOrWhiteSpace(SubpastaArquivos)
                    ? null
                    : subpastas.FirstOrDefault(dir =>
                        string.Equals(dir.Name, SubpastaArquivos.Trim(), StringComparison.OrdinalIgnoreCase));

                model.Pastas = subpastas
                    .Where(dir => pastaTransparente == null || dir.FullName != pastaTransparente.FullName)
                    .OrderBy(dir => dir.Name, StringComparer.CurrentCultureIgnoreCase)
                    .Select(dir => new PastaItem
                    {
                        Nome = dir.Name,
                        CaminhoRelativo = Combinar(model.CaminhoAtual, dir.Name)
                    })
                    .ToList();

                // Com as frentes configuradas, PDFs soltos na raiz não são exibidos
                var arquivos = naRaiz && frentes.Length > 0
                    ? new List<ManualItem>()
                    : ListarPdfs(caminhoCompleto, model.CaminhoAtual);

                if (pastaTransparente != null)
                {
                    arquivos.AddRange(ListarPdfs(
                        pastaTransparente.FullName,
                        Combinar(model.CaminhoAtual, pastaTransparente.Name)));
                }

                model.Arquivos = arquivos
                    .OrderBy(arquivo => arquivo.Titulo, StringComparer.CurrentCultureIgnoreCase)
                    .ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Erro ao listar os manuais em {Pasta}", caminhoCompleto);
                ViewBag.Erro = "Não foi possível ler a pasta dos manuais. " +
                               "Verifique se o servidor tem permissão de acesso ao caminho configurado.";
            }

            return View(model);
        }

        // GET: /Manuais/Download?arquivo=VISUAL LINX/PDFs/NOME DO MANUAL.pdf
        public IActionResult Download(string? arquivo)
        {
            if (string.IsNullOrWhiteSpace(arquivo) || string.IsNullOrWhiteSpace(PastaRaiz))
                return BadRequest();

            var relativo = NormalizarRelativo(arquivo);

            if (!CaminhoPermitido(relativo))
                return NotFound();

            var caminhoCompleto = ResolverCaminho(relativo);
            if (caminhoCompleto == null)
                return BadRequest();

            if (!caminhoCompleto.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                return BadRequest();

            if (!System.IO.File.Exists(caminhoCompleto))
                return NotFound();

            try
            {
                var stream = new FileStream(
                    caminhoCompleto,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.ReadWrite);

                // Informar o fileDownloadName faz o navegador baixar o arquivo
                // (Content-Disposition: attachment) em vez de abrir na tela.
                return File(stream, "application/pdf", Path.GetFileName(caminhoCompleto));
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Erro ao baixar o manual {Arquivo}", caminhoCompleto);
                return StatusCode(500, "Não foi possível baixar o arquivo solicitado.");
            }
        }

        // ------------------------------------------------------------------
        // Apoio
        // ------------------------------------------------------------------

        /// <summary>
        /// Verifica se o caminho relativo começa por uma das frentes configuradas.
        /// Sem frentes configuradas, tudo dentro da raiz é permitido.
        /// </summary>
        private bool CaminhoPermitido(string relativo)
        {
            var frentes = Frentes;

            if (frentes.Length == 0 || string.IsNullOrEmpty(relativo))
                return true;

            var primeiraParte = relativo.Split('/', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
            if (primeiraParte == null)
                return true;

            return frentes.Any(frente =>
                string.Equals(frente, primeiraParte, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Lista os PDFs de uma pasta física, montando o caminho relativo usado no download.
        /// </summary>
        private static List<ManualItem> ListarPdfs(string caminhoFisico, string caminhoRelativo)
        {
            return Directory
                .EnumerateFiles(caminhoFisico, "*.pdf", SearchOption.TopDirectoryOnly)
                .Select(caminho => new FileInfo(caminho))
                .Select(arquivo => new ManualItem
                {
                    NomeArquivo = arquivo.Name,
                    Titulo = Path.GetFileNameWithoutExtension(arquivo.Name),
                    CaminhoRelativo = Combinar(caminhoRelativo, arquivo.Name)
                })
                .ToList();
        }

        /// <summary>
        /// Deixa o caminho relativo sempre no formato "PASTA/SUBPASTA/ARQUIVO.pdf".
        /// </summary>
        private static string NormalizarRelativo(string? relativo)
        {
            if (string.IsNullOrWhiteSpace(relativo))
                return string.Empty;

            return relativo.Replace('\\', '/').Trim('/');
        }

        private static string Combinar(string pastaAtual, string nome)
        {
            return string.IsNullOrEmpty(pastaAtual) ? nome : $"{pastaAtual}/{nome}";
        }

        private static string? CaminhoPai(string caminhoAtual)
        {
            if (string.IsNullOrEmpty(caminhoAtual))
                return null;

            var posicao = caminhoAtual.LastIndexOf('/');
            return posicao < 0 ? string.Empty : caminhoAtual[..posicao];
        }

        private static List<TrilhaItem> MontarTrilha(string caminhoAtual)
        {
            var trilha = new List<TrilhaItem>
            {
                new() { Nome = "Manuais Linx", CaminhoRelativo = string.Empty }
            };

            if (string.IsNullOrEmpty(caminhoAtual))
                return trilha;

            var acumulado = string.Empty;
            foreach (var parte in caminhoAtual.Split('/', StringSplitOptions.RemoveEmptyEntries))
            {
                acumulado = Combinar(acumulado, parte);
                trilha.Add(new TrilhaItem { Nome = parte, CaminhoRelativo = acumulado });
            }

            return trilha;
        }

        /// <summary>
        /// Converte o caminho relativo em caminho físico, garantindo que ele
        /// continua dentro da pasta raiz. Devolve null quando o caminho é inválido.
        /// </summary>
        private string? ResolverCaminho(string relativo)
        {
            var raiz = Path.GetFullPath(PastaRaiz);

            if (string.IsNullOrEmpty(relativo))
                return raiz;

            var partes = relativo.Split('/', StringSplitOptions.RemoveEmptyEntries);

            // Bloqueia "..", "." e nomes com caracteres inválidos
            if (partes.Any(parte => parte is ".." or "." ||
                                    parte.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0))
                return null;

            var caminhoRelativo = Path.Combine(partes);
            if (Path.IsPathRooted(caminhoRelativo))
                return null;

            var completo = Path.GetFullPath(Path.Combine(raiz, caminhoRelativo));

            var raizComSeparador = raiz.EndsWith(Path.DirectorySeparatorChar)
                ? raiz
                : raiz + Path.DirectorySeparatorChar;

            if (!completo.StartsWith(raizComSeparador, StringComparison.OrdinalIgnoreCase))
                return null;

            return completo;
        }
    }
}
