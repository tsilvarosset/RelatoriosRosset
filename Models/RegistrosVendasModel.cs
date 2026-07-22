namespace RelatoriosRosset.Models
{
    public class RegistrosVendasModel
    {
        public string TIPO_REGISTRO { get; set; }
        public string CPF_CGC { get; set; }
        public string MODELO { get; set; }
        public string SERIE_NF { get; set; }
        public string NF { get; set; }
        public string CODIGO_FISCAL_OPERACAO { get; set; }
        public string TRIBUT_ICMS { get; set; }
        public string ITEM_CFE { get; set; }
        public string CODIGO_ITEM { get; set; }
        public string CODIGO_BARRA { get; set; }
        public int QTDE_ITEM { get; set; }
        public Decimal VALOR_ITEM { get; set; }
        public Decimal DESCONTO_ITEM { get; set; }
        public Decimal BASE_IMPOSTO { get; set; }
        public Decimal BASE_IMPOSTO_SUBST { get; set; }
        public Decimal VALOR_IPI { get; set; }
        public Decimal TAXA_IMPOSTO { get; set; }
        public Decimal VALOR_CONTABIL { get; set; }
        public Decimal VALOR_ICMS { get; set; }
        public string FILIAL_CHAVE { get; set; }
        public string NF_SAIDA_CHAVE { get; set; }
        public string SERIE_NF_CHAVE { get; set; }
        public string COD_MATRIZ_FISCAL { get; set; }
        public string MATRIZ_FISCAL { get; set; }
        public DateTime EMISSAO { get; set; }
        public string DESCRICAO_ITEM { get; set; }
        public string UNIDADE { get; set; }
        public string RG_IE { get; set; }
        public string CLASSIF_FISCAL { get; set; }
        public string UF { get; set; }
        public string PAIS { get; set; }
        public string ITEM_IMPRESSAO { get; set; }
        public string COD_FILIAL { get; set; }
        public short NOTA_CANCELADA { get; set; }
        public string SERIE_CFE_SAT { get; set; }
        public string CHAVE_NFE { get; set; }
    }
}
