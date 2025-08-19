namespace RelatoriosRosset.Models
{
    public class TicketsNotasModel
    {
        public string FILIAL {  get; set; }
        public DateTime DATA { get; set; }
        public string? TICKET { get; set; }
        public Decimal? VALOR_PAGO { get; set; }
        public string? NUMERO_FISCAL_VENDA { get; set; }
        public string? NUMERO_FISCAL_TROCA { get; set; }
        public string? NUMERO_FISCAL_CANCELAMENTO { get; set; }
    }
}
