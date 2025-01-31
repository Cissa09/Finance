using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;

public class Excel
{
    public static string Exportar(List<(DateTime Data, double Minimo, double Maximo, double Variacao)> monthlyData,
                                  List<(DateTime Data, double Minimo, double Maximo, double Variacao)> annualData,
                                  string ticker)
    {
        string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"Historico_{ticker}.xlsx");

        using (var workbook = new XLWorkbook())
        {
            CriarAba(workbook, "Mensal", monthlyData);
            CriarAba(workbook, "Anual", annualData);
            workbook.SaveAs(filePath);
        }

        return filePath;
    }

    private static void CriarAba(XLWorkbook workbook, string nomeAba, List<(DateTime Data, double Minimo, double Maximo, double Variacao)> dados)
    {
        var worksheet = workbook.Worksheets.Add(nomeAba);

        worksheet.Cell(1, 1).Value = "Data";
        worksheet.Cell(1, 2).Value = "Valor Mínimo";
        worksheet.Cell(1, 3).Value = "Valor Máximo";
        worksheet.Cell(1, 4).Value = "Variação (%)";

        int row = 2;
        foreach (var item in dados)
        {
            worksheet.Cell(row, 1).Value = item.Data; // 🔹 Salvar como DateTime real
            worksheet.Cell(row, 1).Style.DateFormat.Format = "dd/MM/yyyy"; // 🔹 Formatar corretamente

            worksheet.Cell(row, 2).Value = item.Minimo;
            worksheet.Cell(row, 3).Value = item.Maximo;
            worksheet.Cell(row, 4).Value = item.Variacao * 100; // 🔹 Não adicionar "%" aqui, o Excel pode formatar depois

            row++;
        }

        worksheet.Columns().AdjustToContents();
    }
}
