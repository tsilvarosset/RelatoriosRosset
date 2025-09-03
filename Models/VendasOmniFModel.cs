namespace RelatoriosRosset.Models
{
    public class VendasOmniFModel
    {
        public string CODIGO_FILIAL { get; set; }
        public string FILIAL { get; set; }
        public DateTime DATA_VENDA { get; set; }
        public Decimal VALOR_PAGO { get; set; }
        public string CLIENTE_VAREJO { get; set; }
        public string CPF_CGC { get; set; }
    }
}
