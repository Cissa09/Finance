using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;

using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;

public class Excel
{
    /// <summary>
    /// Gera um arquivo Excel contendo duas abas: uma com dados mensais e outra com dados anuais.
    /// </summary>
    /// <param name="monthlyData">Lista de dados processados para análise mensal.</param>
    /// <param name="annualData">Lista de dados processados para análise anual.</param>
    /// <param name="ticker">Código do ativo (ex: PETR4.SA) para nomeação do arquivo.</param>
    /// <returns>Retorna o caminho do arquivo Excel gerado.</returns>
    public static string Exportar(List<(DateTime Data, double Minimo, double Maximo, double Variacao)> monthlyData,
                                  List<(DateTime Data, double Minimo, double Maximo, double Variacao)> annualData,
                                  string ticker)
    {
        // Define o caminho do arquivo Excel na Área de Trabalho do usuário
        string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"Historico_{ticker}.xlsx");

        using (var workbook = new XLWorkbook()) // Cria um novo arquivo Excel
        {
            // Cria duas abas no Excel: uma para os dados mensais e outra para os anuais
            CriarAba(workbook, "Mensal", monthlyData);
            CriarAba(workbook, "Anual", annualData);

            // Salva o arquivo Excel no caminho definido
            workbook.SaveAs(filePath);
        }

        return filePath; // Retorna o caminho do arquivo gerado
    }

    /// <summary>
    /// Cria uma aba no arquivo Excel e preenche com os dados processados.
    /// </summary>
    /// <param name="workbook">Instância do arquivo Excel.</param>
    /// <param name="nomeAba">Nome da aba a ser criada (Mensal ou Anual).</param>
    /// <param name="dados">Lista de dados a serem inseridos na aba.</param>
    private static void CriarAba(XLWorkbook workbook, string nomeAba, List<(DateTime Data, double Minimo, double Maximo, double Variacao)> dados)
    {
        var worksheet = workbook.Worksheets.Add(nomeAba); // Adiciona uma nova aba no Excel

        // Define os cabeçalhos das colunas
        worksheet.Cell(1, 1).Value = "Data";
        worksheet.Cell(1, 2).Value = "Valor Mínimo";
        worksheet.Cell(1, 3).Value = "Valor Máximo";
        worksheet.Cell(1, 4).Value = "Variação (%)";

        int row = 2; // Começa a preencher os dados a partir da segunda linha

        foreach (var item in dados)
        {
            // Insere a data como um objeto DateTime, garantindo que o Excel reconheça como data real
            worksheet.Cell(row, 1).Value = item.Data;
            worksheet.Cell(row, 1).Style.DateFormat.Format = "dd/MM/yyyy"; // Define o formato da data no Excel

            // Insere os valores numéricos diretamente, permitindo cálculos no Excel
            worksheet.Cell(row, 2).Value = item.Minimo;
            worksheet.Cell(row, 3).Value = item.Maximo;
            worksheet.Cell(row, 4).Value = item.Variacao * 100; // Armazena a variação como um número (sem %)

            row++;
        }

        // Ajusta automaticamente a largura das colunas para melhor visualização dos dados
        worksheet.Columns().AdjustToContents();
    }
}

