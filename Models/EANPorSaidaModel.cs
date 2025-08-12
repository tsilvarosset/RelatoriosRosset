namespace RelatoriosRosset.Models
{
    public class EANPorSaidaModel
    {
        public string? NF_NUMERO { get; set; }
        public string? SERIE { get; set; }
        public string? PRODUTO { get; set; }
        public string? CODIGO_BARRA { get; set; }
        public DateTime EMISSAO { get; set; }
        public string? FILIAL_ORIGEM { get; set; }
        public string? FILIAL_DESTINO { get; set; }
        public Decimal VALOR_UNITARIO { get; set; }
        public Decimal QTDE_ITEM { get; set; }
        public Decimal VALOR_TOTAL { get; set; }
        public Decimal VALOR_ICMS { get; set; }
        public string? CODIGO_FISCAL_OPERACAO { get; set; }
        public string? CHAVE_NFE { get; set; }
    }
}
