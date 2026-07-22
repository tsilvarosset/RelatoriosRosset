using System.ComponentModel.DataAnnotations.Schema;

namespace RelatoriosRosset.Models
{
    public class FiliaisAtivasFModel
    {
        public string Cgc_Cpf { get; set; }
        public string Codigo_Filial { get; set; }
        public string Filial { get; set; }
        [NotMapped]
        public string Mensagem { get; set; }

    }
}
