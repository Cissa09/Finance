using System;
using System.Threading.Tasks;

class Program
{
    static async Task Main(string[] args)
    {
        string ticker;

        // Verifica se o usuário passou o ticker como argumento na execução
        if (args.Length == 0)
        {
            // Caso o ticker não tenha sido informado, solicita via console
            Console.Write("Digite o ticker da ação (ex: PETR4.SA): ");
            ticker = Console.ReadLine()?.Trim();

            // Se o usuário não informar nada, encerra o programa
            if (string.IsNullOrEmpty(ticker))
            {
                Console.WriteLine("Ticker inválido. Saindo...");
                return;
            }
        }
        else
        {
            // Se foi passado como argumento, usa o valor fornecido
            ticker = args[0];
        }

        // Define a data de início fixa para busca dos dados
        DateTime startDate = new DateTime(2015, 1, 1);

        try
        {
            // Chama o método que busca os dados e gera o Excel
            string filePath = ProcessaDados.GerarExcelHistorico(ticker, startDate);

            // Exibe no console o caminho do arquivo gerado
            Console.WriteLine($"Arquivo Excel gerado com sucesso: {filePath}");
        }
        catch (Exception ex)
        {
            // Em caso de erro, exibe a mensagem para o usuário
            Console.WriteLine($"Erro ao processar dados: {ex.Message}");
        }
    }
}
