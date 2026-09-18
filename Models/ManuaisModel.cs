namespace RelatoriosRosset.Models
{
    /// <summary>
    /// Conteúdo da pasta que está sendo exibida na página Manuais Linx.
    /// É o model da view Views/Manuais/Manuais.cshtml.
    /// </summary>
    public class Manuais
    {
        /// <summary>Caminho relativo da pasta atual. Vazio = raiz.</summary>
        public string CaminhoAtual { get; set; } = string.Empty;

        /// <summary>Caminho relativo da pasta anterior. Null quando já está na raiz.</summary>
        public string? CaminhoPai { get; set; }

        /// <summary>Trilha de navegação (Manuais Linx › VISUAL LINX › …).</summary>
        public List<TrilhaItem> Trilha { get; set; } = new();

        /// <summary>Subpastas da pasta atual (VISUAL LINX, LINXPOS, …).</summary>
        public List<PastaItem> Pastas { get; set; } = new();

        /// <summary>Arquivos PDF da pasta atual.</summary>
        public List<ManualItem> Arquivos { get; set; } = new();

        public bool Vazio => Pastas.Count == 0 && Arquivos.Count == 0;
    }

    /// <summary>
    /// Um manual (PDF) disponível para download.
    /// </summary>
    public class ManualItem
    {
        /// <summary>Nome do arquivo com extensão, ex.: "005045 - ENTRADA DE NOTA FISCAL POR XML.pdf".</summary>
        public string NomeArquivo { get; set; } = string.Empty;

        /// <summary>Nome exibido na tela (nome do arquivo sem a extensão).</summary>
        public string Titulo { get; set; } = string.Empty;

        /// <summary>
        /// Caminho relativo à pasta raiz, usado no link de download.
        /// Ex.: "VISUAL LINX/005045 - ENTRADA DE NOTA FISCAL POR XML.pdf".
        /// </summary>
        public string CaminhoRelativo { get; set; } = string.Empty;
    }

    /// <summary>
    /// Uma subpasta dentro da pasta de manuais (ex.: VISUAL LINX, LINXPOS).
    /// </summary>
    public class PastaItem
    {
        /// <summary>Nome da pasta exibido na tela.</summary>
        public string Nome { get; set; } = string.Empty;

        /// <summary>Caminho relativo à pasta raiz, usado no link de navegação.</summary>
        public string CaminhoRelativo { get; set; } = string.Empty;
    }

    /// <summary>
    /// Item da trilha de navegação (breadcrumb).
    /// </summary>
    public class TrilhaItem
    {
        public string Nome { get; set; } = string.Empty;

        /// <summary>Caminho relativo da pasta. Vazio = raiz.</summary>
        public string CaminhoRelativo { get; set; } = string.Empty;
    }
}
