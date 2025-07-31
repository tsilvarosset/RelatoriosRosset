namespace RelatoriosRosset.Models
{
    public class EstoqueEANModel
    {
        public string FILIAL { get; set; }
        public string CODIGO_BARRA { get; set; }
        public string PRODUTO { get; set; }
        public int ESTOQUE { get; set; }
        public decimal ULTIMO_CUSTO { get; set; }
        public DateTime DATA_SALDO { get; set; }

    }
}
