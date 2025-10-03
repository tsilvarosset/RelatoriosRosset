namespace RelatoriosRosset.Models
{
    public class DescontosPModel
    {
        public string COD_FILIAL { get; set; }
        public string FILIAL { get; set; }
        public string OPERACAO_VENDA { get; set; }
        public Decimal LIMITE_DESCONTO { get; set; }
        public Decimal LIMITE_DESCONTO_GERENTE { get; set; }
        public Decimal LIMITE_DESCONTO_ITEM { get; set; }
        public Decimal LIMITE_DESCONTO_ITEM_GERENTE { get; set; }
    }
}
