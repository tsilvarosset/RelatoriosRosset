namespace RelatoriosRosset.Models
{
    public class LivroSaida
    {
        public string? FILIAL { get; set; }
        public string? NF_SAIDA { get; set; }
        public string? DESTINO_CLIENTE { get; set; }
        public string? DESTINO_FILIAL { get; set; }
        public string? SERIE_NF { get; set; }
        public string? SERIE_NF_OFICIAL { get; set; }
        public string? IMPOSTO { get; set; }
        public Decimal? VALOR_CONTABIL { get; set; }
        public Decimal? BASE_IMPOSTO { get; set; }
        public Decimal? TAXA_IMPOSTO { get; set; }
        public Decimal? VALOR_IMPOSTO { get; set; }
        public Decimal? VALOR_IMPOSTO_OUTROS { get; set; }
        public Decimal? VALOR_IMPOSTO_ISENTO { get; set; }
        public string? CODIGO_FISCAL_OPERACAO { get; set; }
        public string? DENOMINACAO_CFOP { get; set; }
        public DateTime EMISSAO { get; set; }
        public string? CODIGO_ITEM { get; set; }
        public Decimal? QTDE_ITEM { get; set; }
        public Decimal? PRECO_UNITARIO { get; set; }
        public Decimal? VALOR_BRUTO_ITEM { get; set; }
        public string? DESCRICAO_ITEM { get; set; }
        public string? UNIDADE { get; set; }
        public string? CLASSIF_FISCAL { get; set; }
    }
}
