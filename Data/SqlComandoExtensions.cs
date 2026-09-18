using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using System;
using System.Data;
using System.Threading;
using System.Threading.Tasks;

namespace RelatoriosRosset.Data
{
    /// <summary>
    /// Helpers para criar conexões e comandos ADO.NET a partir do DbContext
    /// já com CommandTimeout definido.
    ///
    /// Motivo: comandos criados via connection.CreateCommand() NÃO herdam o
    /// Database.SetCommandTimeout() do EF — eles caem no padrão do SqlClient,
    /// que é 30 segundos. Estes helpers evitam esquecer de definir o timeout.
    /// </summary>
    public static class SqlComandoExtensions
    {
        /// <summary>Timeout padrão, em segundos, quando nenhum é informado.</summary>
        public const int TimeoutPadraoSegundos = 3600; // 1 hora

        /// <summary>
        /// Abre uma SqlConnection NOVA usando a connection string do contexto.
        /// Não usa a conexão interna do EF, então pode ser descartada com "using"
        /// sem quebrar o DbContext.
        /// </summary>
        public static async Task<SqlConnection> AbrirConexaoAsync(
            this DbContext context,
            CancellationToken ct = default)
        {
            var connection = new SqlConnection(context.Database.GetConnectionString());
            await connection.OpenAsync(ct);
            return connection;
        }

        /// <summary>
        /// Cria um SqlCommand já com o CommandTimeout definido.
        /// </summary>
        public static SqlCommand CriarComando(
            this SqlConnection connection,
            string sql,
            int? timeoutSegundos = null)
        {
            var command = connection.CreateCommand();
            command.CommandText = sql;
            command.CommandTimeout = timeoutSegundos ?? TimeoutPadraoSegundos;
            return command;
        }

        /// <summary>
        /// Adiciona um parâmetro. Valor null — ou string vazia/em branco —
        /// vira DBNull, que é o comportamento esperado pelas procedures daqui.
        /// Retorna o próprio comando para permitir encadeamento.
        /// </summary>
        public static SqlCommand ComParametro(
            this SqlCommand command,
            string nome,
            SqlDbType tipo,
            object valor)
        {
            if (valor is string texto && string.IsNullOrWhiteSpace(texto))
                valor = null;

            command.Parameters.Add(new SqlParameter
            {
                ParameterName = nome,
                SqlDbType = tipo,
                Value = valor ?? DBNull.Value
            });

            return command;
        }

        /// <summary>
        /// Atalho para o caso mais comum: abrir conexão, montar os parâmetros,
        /// executar e fechar tudo.
        /// </summary>
        public static async Task ExecutarProcedureAsync(
            this DbContext context,
            string sql,
            Action<SqlCommand> configurarParametros = null,
            int? timeoutSegundos = null,
            CancellationToken ct = default)
        {
            using var connection = await context.AbrirConexaoAsync(ct);
            using var command = connection.CriarComando(sql, timeoutSegundos);

            configurarParametros?.Invoke(command);

            await command.ExecuteNonQueryAsync(ct);
        }
    }
}
