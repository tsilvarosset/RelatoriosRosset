namespace RelatoriosRosset.Models
{
    public class LojaCaixaCartaoModel
    {
        public string? CODIGO_FILIAL { get; set; }
        public string? FILIAL { get; set; }
        public string? VENDEDOR { get; set; }
        public string? TICKET { get; set; }
        public string? LANCAMENTO_CAIXA { get; set; }
        public string? TERMINAL { get; set; }
        public DateTime? DATA_VENDA { get; set; }
        public string? CODIGO_CONSUMIDOR { get; set; }
        public string? PARCELA { get; set; }
        public string? NUMERO_CHEQUE_CARTAO { get; set; }
        public string? NUMERO_APROVACAO_CARTAO { get; set; }
        public string? NUMERO_TITULO { get; set; }
        public Decimal? VALOR_ORIGINAL { get; set; }
        public Decimal? TAXA_ADMINISTRACAO { get; set; }
        public Decimal? VALOR_A_RECEBER { get; set; }
        public DateTime? DATA_HORA_TEF { get; set; }
        public DateTime? DATA_EMISSAO { get; set; }
        public DateTime? VENCIMENTO_REAL { get; set; }
        public string? DESC_TIPO_PGTO { get; set; }
        public string? CMC7_CVCARTAO { get; set; }
        public int? LANCAMENTO { get; set; }
    }
}
