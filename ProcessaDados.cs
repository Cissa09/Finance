using System;
using System.Linq;
using System.Collections.Generic;
using YahooQuotesApi;

public class ProcessaDados
{
    /// <summary>
    /// Obtém os dados históricos de um ativo financeiro e gera um arquivo Excel com a análise mensal e anual.
    /// </summary>
    /// <param name="ticker">Código do ativo (ex: PETR4.SA).</param>
    /// <param name="startDate">Data de início da análise.</param>
    /// <returns>Retorna o caminho do arquivo Excel gerado.</returns>
    public static string GerarExcelHistorico(string ticker, DateTime startDate)
    {
        // Cria uma instância da API do Yahoo Finance para buscar os dados
        var yahooQuotes = new YahooQuotesBuilder().Build();

        // Obtém os dados históricos do ativo fornecido
        var historyResult = yahooQuotes.GetHistoryAsync(ticker).Result;

        // Se não houver dados para o ativo, lança uma exceção
        if (!historyResult.HasValue)
            throw new Exception("Nenhum dado encontrado para esse ativo.");

        // Extrai os dados históricos
        var history = historyResult.Value;

        // Processa os dados históricos para gerar análise mensal e anual
        var monthlyData = ProcessarDados(history, startDate, true);  // Análise mensal       
        var annualData = ProcessarDados(history, startDate.AddYears(-10), false); // Análise anual com histórico de 20 anos.

        // Chama a classe Excel para gerar o arquivo com os dados processados
        string filePath = Excel.Exportar(monthlyData, annualData, ticker);

        // Retorna o caminho do arquivo gerado
        return filePath;
    }

    /// <summary>
    /// Processa os dados históricos e calcula o valor mínimo, máximo e variação percentual no período.
    /// </summary>
    /// <param name="history">Dados históricos do ativo.</param>
    /// <param name="startDate">Data de início da análise.</param>
    /// <param name="isMensal">Define se o processamento será mensal (true) ou anual (false).</param>
    /// <returns>Lista contendo as informações do período processado.</returns>
    private static List<(DateTime Data, double Minimo, double Maximo, double Variacao)> ProcessarDados(History history, DateTime startDate, bool isMensal)
    {
        return history.Ticks
            .Select(t => new { Date = t.Date.ToDateTimeUtc(), t.Low, t.High }) // Converte a data para DateTime e extrai os valores de preço
            .Where(t => t.Date >= startDate) // Filtra apenas os dados a partir da data de início
            .GroupBy(t => isMensal ? $"{t.Date.Year}-{t.Date.Month:D2}" : $"{t.Date.Year}") // Agrupa por mês ou ano
            .Select(g =>
            {
                // Obtém os valores mínimo e máximo no período
                double min = Math.Round(g.Min(t => t.Low), 2);
                double max = Math.Round(g.Max(t => t.High), 2);

                // Calcula a variação percentual (máximo / mínimo - 1), evitando divisão por zero
                double variacao = (min == 0 || double.IsNaN(max) || double.IsNaN(min)) ? 0 : Math.Round((max / min) - 1, 4);

                // Define a data de referência com base no agrupamento (mensal ou anual)
                DateTime dataReferencia = isMensal
                    ? new DateTime(int.Parse(g.Key.Split('-')[0]), int.Parse(g.Key.Split('-')[1]), 1) // Mês/Ano
                    : new DateTime(int.Parse(g.Key), 1, 1); // Apenas Ano

                return (Data: dataReferencia, Minimo: min, Maximo: max, Variacao: variacao);
            })
            .OrderByDescending(d => d.Data) // Ordena os resultados do mais novo para o mais antigo
            .ToList();
    }
}