using System;
using System.Linq;
using System.Collections.Generic;
using YahooQuotesApi;

public class ProcessaDados
{
    public static string GerarExcelHistorico(string ticker, DateTime startDate)
    {
        var yahooQuotes = new YahooQuotesBuilder().Build();
        var historyResult = yahooQuotes.GetHistoryAsync(ticker).Result;

        if (!historyResult.HasValue)
            throw new Exception("Nenhum dado encontrado para esse ativo.");

        var history = historyResult.Value;
        var monthlyData = ProcessarDados(history, startDate, true);
        var annualData = ProcessarDados(history, startDate, false);

        string filePath = Excel.Exportar(monthlyData, annualData, ticker);
        return filePath;
    }

    private static List<(DateTime Data, double Minimo, double Maximo, double Variacao)> ProcessarDados(History history, DateTime startDate, bool isMensal)
    {
        return history.Ticks
            .Select(t => new { Date = t.Date.ToDateTimeUtc(), t.Low, t.High })
            .Where(t => t.Date >= startDate)
            .GroupBy(t => isMensal ? $"{t.Date.Year}-{t.Date.Month:D2}" : $"{t.Date.Year}")
            .Select(g =>
            {
                double min = Math.Round(g.Min(t => t.Low), 2);
                double max = Math.Round(g.Max(t => t.High), 2);
                double variacao = (min == 0 || double.IsNaN(max) || double.IsNaN(min)) ? 0 : Math.Round((max / min) - 1, 4);

                DateTime dataReferencia = isMensal
                    ? new DateTime(int.Parse(g.Key.Split('-')[0]), int.Parse(g.Key.Split('-')[1]), 1)
                    : new DateTime(int.Parse(g.Key), 1, 1);

                return (Data: dataReferencia, Minimo: min, Maximo: max, Variacao: variacao);
            })
            .OrderBy(d => d.Data)
            .ToList();
    }
}
