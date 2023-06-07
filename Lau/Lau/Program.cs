using HtmlAgilityPack;
using OfficeOpenXml;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;

class Program
{
    static void Main()
    {
        /////////////////////////////////////////
        /////////////////////////////////////////
        /////////////////IMPORTANTE//////////////
        /////////////////////////////////////////
        /////////////////////////////////////////
        // Ruta del archivo .xlsx
        string xlsxFilePath = "C:\\Users\\cejit\\OneDrive\\Escritorio\\dataset_fran_uwu.xlsx";

        /////////////////////////////////////////
        /////////////////////////////////////////
        /////////////////IMPORTANTE//////////////
        /////////////////////////////////////////
        /////////////////////////////////////////
        // Directorio de salida para los archivos PDF generados
        string outputDirectory = "C:\\Users\\cejit\\OneDrive\\Escritorio\\Testing";

        // Crear el directorio de salida si no existe
        Directory.CreateDirectory(outputDirectory);

        // Establecer el contexto de licencia de EPPlus
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // Lista de palabras clave
        List<string> palabrasClave = new List<string>()
        {
            "archivo", "museo", "bibliotecario", "acceso", "ciencia", "tecnología", "información",
            "biblioteca", "digital", "repositorio", "audiovisual", "difusión", "educación",
            "comunitaria", "radio", "televisión", "telefonía", "internet", "capacitación",
            "accesibilidad", "recursos", "bibliográficos", "documentales", "documentación",
            "investigadores", "encuestas", "censos", "universidad"
        };

        // Leer los enlaces del archivo .xlsx y procesar cada página
        using (ExcelPackage package = new ExcelPackage(new FileInfo(xlsxFilePath)))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault(); // Obtener la primera hoja de trabajo

            if (worksheet != null)
            {
                int rowCount = worksheet.Dimension.Rows;
                int columnCount = worksheet.Dimension.Columns;

                for (int row = 2; row <= rowCount; row++) // Empezar desde la segunda fila para omitir la fila de encabezados
                {
                    string url = worksheet.Cells[row, 14].Value?.ToString(); // Columna N que contiene los enlaces
                    string numeroLey = worksheet.Cells[row, 2].Value?.ToString(); // Columna B que contiene los números de ley
                    string titulo = worksheet.Cells[row, 3].Value?.ToString(); // Columna C que contiene los títulos

                    if (!string.IsNullOrEmpty(url))
                    {
                        // Abrir el enlace en el navegador web predeterminado
                        OpenLinkInBrowser(url);

                        // Descargar el contenido HTML de la página web
                        string htmlContent = DownloadHtmlContent(url);

                        // Extraer el texto visible de la página web, excluyendo encabezados, pies de página, scripts y estilos
                        string extractedText = ExtractVisibleTextFromHtml(htmlContent);

                        // Verificar si la página contiene al menos una palabra clave
                        if (PageContainsKeywords(extractedText, palabrasClave))
                        {
                            // Generar un nuevo archivo PDF con el texto extraído
                            char[] invalidChars = Path.GetInvalidFileNameChars();
                            titulo = string.Join("_", titulo.Split(invalidChars, StringSplitOptions.RemoveEmptyEntries));
                            string outputFilePath = Path.Combine(outputDirectory, $"numero_norma_{titulo}.pdf");
                            SaveTextAsPdf(extractedText, outputFilePath);

                            Console.WriteLine($"Se ha guardado el texto visible de la URL {url} en el archivo PDF: {outputFilePath}");
                        }
                    }
                }
            }
        }

        Console.WriteLine("Proceso completado.");
        Console.ReadLine();
    }

    static string DownloadHtmlContent(string url)
    {
        using (HttpClient client = new HttpClient())
        {
            return client.GetStringAsync(url).Result;
        }
    }

    static string ExtractVisibleTextFromHtml(string htmlContent)
    {
        HtmlDocument doc = new HtmlDocument();
        doc.LoadHtml(htmlContent);

        // Obtener todos los nodos de texto visibles en el documento, excluyendo encabezados, pies de página, scripts y estilos
        var textNodes = doc.DocumentNode.DescendantsAndSelf()
            .Where(n => n.NodeType == HtmlNodeType.Text &&
                        !IsInHeaderOrFooter(n) &&
                        !IsInScriptOrStyle(n))
            .Select(n => n.InnerText.Trim());

        // Unir los nodos de texto en una sola cadena
        string extractedText = string.Join(" ", textNodes);

        return extractedText;
    }

    static bool IsInHeaderOrFooter(HtmlNode node)
    {
        // Verificar si el nodo está contenido en un encabezado o pie de página
        while (node.ParentNode != null)
        {
            if (node.ParentNode.Name.Equals("header", StringComparison.OrdinalIgnoreCase) ||
                node.ParentNode.Name.Equals("footer", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            node = node.ParentNode;
        }

        return false;
    }

    static bool IsInScriptOrStyle(HtmlNode node)
    {
        // Verificar si el nodo está contenido en un elemento <script> o <style>
        while (node.ParentNode != null)
        {
            if (node.ParentNode.Name.Equals("script", StringComparison.OrdinalIgnoreCase) ||
                node.ParentNode.Name.Equals("style", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            node = node.ParentNode;
        }

        return false;
    }

    static void SaveTextAsPdf(string text, string outputFilePath)
    {
        using (FileStream fs = new FileStream(outputFilePath, FileMode.Create))
        {
            Document document = new Document();
            PdfWriter writer = PdfWriter.GetInstance(document, fs);

            document.Open();
            document.Add(new Paragraph(text));
            document.Close();
        }
    }

    static bool PageContainsKeywords(string pageText, List<string> keywords)
    {
        foreach (string keyword in keywords)
        {
            if (pageText.Contains(keyword, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }
        return false;
    }

    static void OpenLinkInBrowser(string url)
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = url,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            Console.WriteLine($"No se pudo abrir el enlace {url} en el navegador web predeterminado: {ex.Message}");
        }
    }
}
