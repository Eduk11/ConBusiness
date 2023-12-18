using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;
using Xceed.Words.NET;
using Xceed.Document.NET;
using DataTable = System.Data.DataTable;

namespace Business
{
    public class Report
    {
        public Report() { }
        public byte[] ExportarExcel()
        {
            DataTable data = new DataTable();
            data.Columns.Add("Nombre");
            data.Columns.Add("Telefono");

            data.Rows.Add("Juan", "44556622");
            data.Rows.Add("Pedro", "995553366");

            using (var libro = new XLWorkbook())
            {
                var hoja = libro.Worksheets.Add(data, "Clientes");

                var celdaCabecera = hoja.Row(1).CellsUsed();
                celdaCabecera.Style.Fill.BackgroundColor = XLColor.FromHtml("#184C78");

                using (var memoria = new MemoryStream())
                {
                    libro.SaveAs(memoria);
                    return memoria.ToArray();
                }
            }
        }

        public byte[] ExportarWord()
        {
            DataTable data = new DataTable();
            data.Columns.Add("Nombre");
            data.Columns.Add("Telefono");

            data.Rows.Add("Juan", "44556622");
            data.Rows.Add("Pedro", "995553366");
            using (var stream = new MemoryStream())
            {
                using (var doc = DocX.Create(stream))
                {
                    // Establecer márgenes más estrechos (0.5 pulgadas en todas las direcciones)
                    doc.MarginLeft = doc.MarginRight = doc.MarginTop = doc.MarginBottom = 36.0F; // Agregué el sufijo 'F' aquí

                    var formatoEncabezado = new Xceed.Document.NET.Formatting
                    {
                        Bold = true,
                        Size = 12,
                        FontColor = System.Drawing.Color.White
                    };

                    var colorFondoEncabezado = System.Drawing.ColorTranslator.FromHtml("#184C78");

                    // Crear una tabla
                    var table = doc.AddTable(data.Rows.Count + 1, data.Columns.Count);

                    // Configurar el formato del encabezado y establecer el color de fondo
                    for (int col = 0; col < data.Columns.Count; col++)
                    {
                        var celdaEncabezado = table.Rows[0].Cells[col];
                        celdaEncabezado.Paragraphs[0].InsertText(data.Columns[col].ColumnName, false, formatoEncabezado);
                        celdaEncabezado.FillColor = colorFondoEncabezado;
                    }

                    // Llenar las celdas de datos
                    for (int row = 0; row < data.Rows.Count; row++)
                    {
                        for (int col = 0; col < data.Columns.Count; col++)
                        {
                            table.Rows[row + 1].Cells[col].Paragraphs[0].InsertText(data.Rows[row][col].ToString());
                        }
                    }

                    // Insertar tabla en el documento
                    doc.InsertTable(table);

                    // Establecer la orientación de la página en horizontal
                    doc.PageLayout.Orientation = Xceed.Document.NET.Orientation.Landscape;

                    doc.Save();
                }

                stream.Position = 0;

                return stream.ToArray();
            }
        }
    }
}
