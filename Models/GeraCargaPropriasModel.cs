namespace RelatoriosRosset.Models
{
    public class GeraCargaPropriasModel
    {
        public string? CODIGO_BARRA { get; set; }
        public string? DESC_PRODUTO { get; set; }
        public string? FILIAL { get; set; }
        public string? PRODUTO { get; set; }
        public string? COR_PRODUTO { get; set; }
        public string? DESC_COR_PRODUTO { get; set; }
        public string? GRADE { get; set; }
        public int? ESTOQUE { get; set; }
        public Decimal? CUSTO_REPOSICAO1 { get; set; }
    }
}