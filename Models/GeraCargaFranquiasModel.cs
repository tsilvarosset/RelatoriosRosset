using System.ComponentModel.DataAnnotations.Schema;

namespace RelatoriosRosset.Models
{
    public class GeraCargaFranquiasModel
    {
        public string Produto { get; set; }
        [NotMapped]
        public string Mensagem { get; set; }
        [NotMapped]
        public string Filial { get; set; }
    }
}
