namespace RelatoriosRosset.Models
{
    public class VendasProdutosModel
    {
        public string CODIGO_FILIAL { get; set; }
        public string FILIAL { get; set; }
        public DateTime DATA_VENDA { get; set; }
        public string PRODUTO { get; set; }
        public string COR_PRODUTO { get; set; }
        public short TAMANHO { get; set; }
        public string CODIGO_BARRA { get; set; }
        public int QTDE_VENDIDA { get; set; }
        public Decimal PRECO_LIQUIDO { get; set; }
        public int QTDE_TICKETS { get; set; }
    }
}
