namespace RelatoriosRosset.Models
{
    public class SeniorModel
    {
        public string CHAVE_NFE { get; set; }
        public string CODIGO_FILIAL	{ get; set; }
        public string NOTA { get; set; }
        public string SERIE { get; set; }
        public DateTime EMISSAO { get; set; }
        public string CFOP { get; set; }
        public Decimal? VALOR_CONTABIL { get; set; }
        public Decimal? BASE_ICMS { get; set; }
        public Decimal? IMPOSTO_ICMS { get; set; }
        public Decimal? BASE_IPI { get; set; }
        public Decimal? IMPOSTO_IPI { get; set; }
        public Decimal? BASE_PIS { get; set; }
        public Decimal? IMPOSTO_PIS { get; set; }
        public Decimal? BASE_COFINS { get; set; }
        public Decimal? IMPOSTO_COFINS { get; set; }
        public Decimal? BASE_FCP { get; set; }
        public Decimal? IMPOSTO_FCP { get; set; }
    }
}
