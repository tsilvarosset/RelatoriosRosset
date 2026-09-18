namespace RelatoriosRosset.Models
{
    public class PedidosGradeModel
    {
        public string GRADE { get; set; }
        public string CLIENTE_ATACADO { get; set; }
        public string PEDIDO { get; set; }
        public string PRODUTO { get; set; }
        public string COR_PRODUTO { get; set; }
        public string TAMANHO { get; set; }
        public string COLECAO { get; set; }
        public string DESC_COLECAO { get; set; }
        public int? QTDE { get; set; }
        public Decimal? VALOR_UNITARIO { get; set; }
        public Decimal? VALOR_PEDIDO { get; set; }
        public DateTime EMISSAO { get; set; }
    }
}
