using System.Collections.Generic;

namespace Logic
{
    public static class CalendarParser
    {
        public Dictionary<int, Dictionary<string, int>> ReadPdfTable(string filePath)
{
    PdfDocument pdfDoc = new PdfDocument(new PdfReader(filePath));
    Dictionary<int, Dictionary<string, int>> tableData = new Dictionary<int, Dictionary<string, int>>();

    for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
    {
        ICollection<SimpleTextExtractionStrategy> textStrategies = new List<SimpleTextExtractionStrategy>();
        LocationTextExtractionStrategy strategy = new LocationTextExtractionStrategy();

        PdfPage page = pdfDoc.GetPage(i);

        string pageText = PdfTextExtractor.GetTextFromPage(page, strategy);

        string[] lines = pageText.Split('\n');

        Dictionary<string, int> rowData = new Dictionary<string, int>();
        foreach (string line in lines)
        {
            string[] cells = line.Split('\t'); // Предполагаем, что данные в таблице разделены табуляцией

            int emptyCellCount = 0;
            int textCellCount = 0;

            foreach (string cell in cells)
            {
                if (string.IsNullOrWhiteSpace(cell))
                {
                    emptyCellCount++;
                }
                else
                {
                    textCellCount++;
                }

                // Добавляем содержимое ячейки в словарь rowData
                // Здесь вы должны определить, какая колонка таблицы соответствует текущей ячейке
                // Например, если вы знаете порядок колонок, то можно использовать индекс ячейки для этого
                // Если данные в таблице имеют определенный формат, можете использовать регулярные выражения для извлечения нужных данных

                // Например, если первая ячейка содержит название, и вторая - количество, то
                // rowData.Add("Название", cell);
                // rowData.Add("Количество", Int32.Parse(cell)); // предполагая, что количество - целое число
            }

            // После обработки строки добавляем информацию о пустых и заполненных ячейках в rowData
            rowData.Add("EmptyCellCount", emptyCellCount);
            rowData.Add("TextCellCount", textCellCount);

            // Добавляем словарь rowData в общий словарь tableData
            tableData.Add(i, rowData);
        }
        
    }

    pdfDoc.Close();
    return tableData;
}
        public static List<Group> Parse(string path)
        {
            throw new System.NotImplementedException();
        }
    }
}
