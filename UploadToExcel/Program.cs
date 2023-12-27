using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

class Program
{
    static void Main()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // veya LicenseContext.Commercial;


        // Log dosyasının yolu
        string logFilePath = @"C:\cdr\new.txt";

        // Excel dosyasının yolu
        string excelFilePath = @"C:\out\output1.xlsx";

        // Log dosyasını oku ve log türlerine göre ayır
        Dictionary<string, List<string>> logsByType = ReadLogs(logFilePath);

        // Excel'e yazma işlemi
        WriteToExcel(excelFilePath, logsByType);
    }

    static Dictionary<string, List<string>> ReadLogs(string filePath)
    {
        Dictionary<string, List<string>> logsByType = new Dictionary<string, List<string>>
        {
            { "Error", new List<string>() },
            { "Info", new List<string>() },
            { "Warning", new List<string>() }
        };

        try
        {
            // Log dosyasını satır satır oku
            using (StreamReader reader = new StreamReader(filePath))
            {
                while (!reader.EndOfStream)
                {
                    string line = reader.ReadLine();

                    // Log türünü belirle
                    string logType = GetLogType(line);

                    // Belirlenen log türüne göre listeye ekle
                    if (logsByType.ContainsKey(logType))
                    {
                        logsByType[logType].Add(line);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Log dosyası okuma hatası: {ex.Message}");
        }

        return logsByType;
    }

    static string GetLogType(string log)
    {
        // Log satırını analiz ederek log türünü belirle
        if (log.Contains("ERROR"))
        {
            return "Error";
        }
        else if (log.Contains("INFO"))
        {
            return "Info";
        }
        else if (log.Contains("WARNING"))
        {
            return "Warning";
        }

        // Belirtilen türler dışındaki loglar için varsayılan "Info" döndür
        return "Info";
    }

    static void WriteToExcel(string filePath, Dictionary<string, List<string>> logsByType)
    {
        try
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                foreach (var logType in logsByType.Keys)
                {
                    var worksheet = package.Workbook.Worksheets.Add(logType);

                    // Log türüne göre logları yaz
                    for (int i = 0; i < logsByType[logType].Count; i++)
                    {
                        worksheet.Cells[i + 1, 1].Value = logsByType[logType][i];
                    }
                }

                // Excel dosyasını kaydet
                package.Save();
            }

            Console.WriteLine("Loglar başarıyla Excel dosyasına yazıldı.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Excel dosyasına yazma hatası: {ex.Message}");
        }
    }
}

