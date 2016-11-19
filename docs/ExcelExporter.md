#Excel Exporter (PoC)

Permite exportar un IEnumerable a una hoja excel. Aplicando formato a las filas en base a valores de cada uno de los elementos del IEnumerable.

Puede generar las columnas del excel leyendo los valores del atributo Display (Name, Order...)

     var items = new List<TestListItem>() { ... };
     var sheetTitle = "Hoja 1";

     var exporter = new Exporter<TestListItem>()
         .SetData(sheetTitle, items)                
         .AddFormat(p => p.Propiedad == 3, IntNovAction.Utils.ExcelExporter.Utils.FontFormat.Bold);