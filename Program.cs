using ClosedXML.Excel;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;

class Program
{
    static void Main()
    {
        Console.Write("Enter JSON file path: ");
        string jsonPath = Console.ReadLine();

        if (!File.Exists(jsonPath))
        {
            Console.WriteLine("File not found!");
            return;
        }

        string excelPath = "output.xlsx";
        var json = File.ReadAllText(jsonPath);

        JToken root = JToken.Parse(json);
        JArray data = root is JArray arr ? arr : new JArray(root);

        var rows = new List<Dictionary<string, string>>();
        var headers = new List<string>();

        foreach (JObject item in data)
        {
            var row = new Dictionary<string, string>();
            FlattenJson(item, "", row);

            foreach (var key in row.Keys)
                if (!headers.Contains(key))
                    headers.Add(key);

            rows.Add(row);
        }

        CreateExcel(headers, rows, excelPath);
        Console.WriteLine($"Excel created: {excelPath}");
    }

    static void FlattenJson(JToken token, string prefix, Dictionary<string, string> row)
    {
        if (token is JObject obj)
        {
            foreach (var prop in obj.Properties())
            {
                string newPrefix = string.IsNullOrEmpty(prefix)
                    ? prop.Name
                    : $"{prefix}_{prop.Name}";

                FlattenJson(prop.Value, newPrefix, row);
            }
        }
        else if (token is JArray array)
        {
            if (!string.IsNullOrEmpty(prefix))
                row[prefix] = string.Join(", ", array);
        }
        else
        {
            if (!string.IsNullOrEmpty(prefix))
                row[prefix] = token?.ToString() ?? "";
        }
    }

    static void CreateExcel(
        List<string> headers,
        List<Dictionary<string, string>> rows,
        string path)
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Data");

        for (int i = 0; i < headers.Count; i++)
        {
            ws.Cell(1, i + 1).Value = headers[i];
            ws.Cell(1, i + 1).Style.Font.Bold = true;
        }

        for (int r = 0; r < rows.Count; r++)
        {
            for (int c = 0; c < headers.Count; c++)
            {
                rows[r].TryGetValue(headers[c], out string value);
                ws.Cell(r + 2, c + 1).Value = value ?? "";
            }
        }

        ws.Columns().AdjustToContents();
        wb.SaveAs(path);
    }
}
