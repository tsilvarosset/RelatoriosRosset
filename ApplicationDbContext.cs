using Microsoft.EntityFrameworkCore;
using RelatoriosRosset.Models;
using System.Collections.Generic;
using System.Reflection.Emit;

namespace RelatoriosRosset
{
    public class ApplicationDbContext : DbContext
    {
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<LojaVendaModel>().HasNoKey();
            modelBuilder.Entity<LojaVendaFModel>().HasNoKey();
            modelBuilder.Entity<NotasDevFModel>().HasNoKey();
            modelBuilder.Entity<NotasDevPModel>().HasNoKey();
            modelBuilder.Entity<FilialModel>().HasNoKey();

            modelBuilder.Entity<LojaNotaFiscalModel>()
            .HasNoKey()
            .ToTable("LOJA_NOTA_FISCAL"); 

            modelBuilder.Entity<FiliaisAtivasPModel>().HasNoKey();
            //modelBuilder.Entity<FiliaisAtivasFModel>().HasNoKey();
            modelBuilder.Entity<EstoqueEANModel>().HasNoKey();
            modelBuilder.Entity<EntradasFModel>().HasNoKey();
            modelBuilder.Entity<SaidasFModel>().HasNoKey();
            modelBuilder.Entity<RegistroSaidaModel>().HasNoKey();
            modelBuilder.Entity<RegistroEntradaModel>().HasNoKey();
            modelBuilder.Entity<EANPorSaidaModel>().HasNoKey();
            modelBuilder.Entity<LivroEntradaModel>().HasNoKey();
            modelBuilder.Entity<LivroSaidaModel>().HasNoKey();
            modelBuilder.Entity<FiliaisAtivasFModel>().HasNoKey().ToView("V_FILIAIS_ATIVAS_FRANQUIAS");
            modelBuilder.Entity<GeraCargaFranquiasModel>().HasNoKey().ToTable("TABELA_CARGA_INV_FRANQUIAS");
            modelBuilder.Entity<TicketsNotasModel>().HasNoKey();
            modelBuilder.Entity<FiliaisAtivasPModel>().HasNoKey().ToView("V_FILIAIS_ATIVAS_PROPRIAS");
            modelBuilder.Entity<GeraCargaPropriasModel>().HasNoKey().ToTable("TABELA_CARGA_INV_PROPRIAS");
            modelBuilder.Entity<GerarProdutoHistoricoModel>().HasNoKey();
            modelBuilder.Entity<LojaCaixaCartaoModel>().HasNoKey();
            modelBuilder.Entity<VendasOmniPModel>().HasNoKey();
            modelBuilder.Entity<VendasOmniFModel>().HasNoKey();
        }



        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
            : base(options)
        {
            //Database.SetCommandTimeout(12000000); // Aumenta o timeout para 120 segundos
            Database.SetCommandTimeout(300); // Timeout de 30 segundos
        }

        public DbSet<LojaVendaModel> V_VENDAS_PROPRIAS { get; set; }
        public DbSet<LojaVendaFModel> V_VENDAS_FRANQUIAS { get; set; }
        public DbSet<NotasDevFModel> NOTAS_DEVOLUCAO_FRANQUIAS { get; set; }
        public DbSet<NotasDevPModel> NOTAS_DEVOLUCAO_PROPRIAS { get; set; }
        public DbSet<FilialModel> FILIAIS { get; set; }
        public DbSet<LojaNotaFiscalModel> LOJA_NOTA_FISCAL { get; set; }
        public DbSet<FiliaisAtivasPModel> V_FILIAIS_ATIVAS_PROPRIAS { get; set; }
        public DbSet<FiliaisAtivasFModel> V_FILIAIS_ATIVAS_FRANQUIAS { get; set; }
        public DbSet<EstoqueEANModel> TABELA_ESTOQUE_EAN_CUSTO { get; set; }
        public DbSet<EntradasFModel> TABELA_ENTRADAS_F { get; set; }
        public DbSet<SaidasFModel> TABELA_SAIDAS_F { get; set; }
        public DbSet<RegistroSaidaModel> W_LF_REGISTRO_SAIDA_IMPOSTO_ITEM { get; set; }
        public DbSet<RegistroEntradaModel> W_LF_REGISTRO_ENTRADA_IMPOSTO_ITEM { get; set; }
        public DbSet<EANPorSaidaModel> EAN_POR_SAIDA { get; set; }
        public DbSet<GeraCargaFranquiasModel> TABELA_CARGA_INV_FRANQUIAS { get; set; }
        public DbSet<TicketsNotasModel> V_TICKETS_NOTAS { get; set; }
        public DbSet<GeraCargaPropriasModel> TABELA_CARGA_INV_PROPRIAS { get; set; }
        public DbSet<GerarProdutoHistoricoModel> ESTOQUE_PRODUTOS_HISTORICO { get; set; }
        public DbSet<LojaCaixaCartaoModel> V_LANCAMENTOS_CAIXA_CARTAO { get; set; }
        public DbSet<VendasOmniPModel> V_VENDAS_OMNI_P { get; set; }
        public DbSet<VendasOmniFModel> V_VENDAS_OMNI_F { get; set; }

    }

}
