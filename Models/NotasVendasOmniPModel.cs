namespace RelatoriosRosset.Models
{
    public class NotasVendasOmniPModel
    {
        public string FILIAL { get; set; }
        public string NF_NUMERO { get; set; }
        public DateTime EMISSAO { get; set; }
        public decimal VALOR_TOTAL { get; set; }
        public string NATUREZA_OPERACAO_CODIGO { get; set; }
        public string CPF_CGC { get; set; }
        public string CLIENTE_VAREJO { get; set; }
        public string CHAVE_NFE { get; set; }
    }
}
