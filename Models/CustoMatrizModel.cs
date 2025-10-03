namespace RelatoriosRosset.Models
{
    public class CustoMatrizModel
    {
        public int? Item { get; set; }
        public string DescItemComposicao { get; set; }
        public string? ItemComposicao { get; set; }
        public string Produto { get; set; }
        public string CorProduto { get; set; }
        public string Grade { get; set; }
        public string CodigoBarra { get; set; }
        public int Qtde { get; set; }
        public decimal? ValorCusto { get; set; } // Nullable decimal
        public string Doc { get; set; }
        public string RomaneioProduto { get; set; }
        public string SerieNf { get; set; }
        public string Cfop { get; set; }
        public string DescricaoCfop { get; set; }
        public string Filial { get; set; }
        public string RateioFilial { get; set; }
        public DateTime DataMov { get; set; }
        public int TotalQtde { get; set; }
        public decimal? ValorTotal { get; set; } // Nullable decimal
        public decimal? ValorImpostoDestacar { get; set; } // Nullable decimal
        public decimal? ValorLiq { get; set; } // Nullable decimal
        public decimal? ValorProducao { get; set; } // Nullable decimal
        public decimal? ValorBruto { get; set; } // Nullable decimal
    }
}
