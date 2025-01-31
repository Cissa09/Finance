using System;
using System.Threading.Tasks;

class Program
{
    static async Task Main(string[] args)
    {
        string ticker;

        if (args.Length == 0)
        {
            Console.Write("Digite o ticker da ação (ex: PETR4.SA): ");
            ticker = Console.ReadLine()?.Trim();

            if (string.IsNullOrEmpty(ticker))
            {
                Console.WriteLine("Ticker inválido. Saindo...");
                return;
            }
        }
        else
        {
            ticker = args[0];
        }

        DateTime startDate = new DateTime(2015, 1, 1);

        try
        {
            string filePath = ProcessaDados.GerarExcelHistorico(ticker, startDate);
            Console.WriteLine($"Arquivo Excel gerado com sucesso: {filePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Erro ao processar dados: {ex.Message}");
        }
    }
}
