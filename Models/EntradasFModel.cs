namespace RelatoriosRosset.Models
{
    public class EntradasFModel
    {
        public string? FILIAL {  get; set; }
        public string? ORIGEM { get; set; }
        public string? NF_ENTRADA { get; set; }
        public string? SERIE { get; set; }
        public string? CFOP { get; set; }
        public DateTime RECEBIMENTO { get; set; }
        public Decimal VALOR_CONTABIL { get; set; }
        public Decimal BASE_IMPOSTO { get; set; }
        public Decimal ALIQUOTA { get; set; }
        public Decimal VALOR_ICMS { get; set; }
        public string? CHAVE_NFE { get; set; }
    }
}
