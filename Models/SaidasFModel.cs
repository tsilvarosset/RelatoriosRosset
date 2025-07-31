namespace RelatoriosRosset.Models
{
    public class SaidasFModel
    {
        public string? FILIAL { get; set; }
        public string? DESTINO { get; set; }
        public string? NF_SAIDA { get; set; }
        public string? SERIE { get; set; }
        public string? CFOP { get; set; }
        public DateTime EMISSAO { get; set; }
        public Decimal VALOR_CONTABIL { get; set; }
        public Decimal BASE_IMPOSTO { get; set; }
        public Decimal ALIQUOTA { get; set; }
        public Decimal VALOR_ICMS { get; set; }
        public string? CHAVE_NFE { get; set; }
    }
}
