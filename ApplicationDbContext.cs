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
            .ToTable("LOJA_NOTA_FISCAL"); // Confirme o nome exato da tabela

            modelBuilder.Entity<FiliaisAtivasPModel>().HasNoKey();
            modelBuilder.Entity<FiliaisAtivasFModel>().HasNoKey();
            modelBuilder.Entity<EstoqueEANModel>().HasNoKey();
            modelBuilder.Entity<EntradasFModel>().HasNoKey();


        }



        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
            : base(options)
        {
            Database.SetCommandTimeout(12000000); // Aumenta o timeout para 120 segundos
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


    }

}
