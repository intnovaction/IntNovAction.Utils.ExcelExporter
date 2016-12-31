# IntNovAction Utils

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Proyectos con utilidades sencillas para hacer cosas varias
Welcome to the IntNovAction-Utils wiki!

## [Excel Exporter (PoC)](./docs/ExcelExporter.md)
[![NuGet](https://img.shields.io/nuget/v/Intnovaction.Utils.ExcelExporter.svg)](https://www.nuget.org/packages/IntNovAction.Utils.ExcelExporter/)

Crea ficheros excel a partir de IEnumerables, permite:
- Trabajar con multiples hojas dentro de un libro.
- Aplicar atributos Display de los items del IEnumerable para obtener los titulos de las columnas o establecerlos manualmente.
- Especificar las columnas a mostrar, dandoles formato.
- Aplicar reglas de formato sencillas usando expresiones lambda, basándose en los valores de los items.
- Realizar transformaciones a los datos para darles formato.

```c#
var exporter = new Exporter()
    .AddSheet<TestListItem>(sheet => sheet.SetData(dataToExport)
        .Name("Sheet Name")
        .AddFormatRule(data => data.Profit < 0, format => format.Color(255, 0, 0))
        .Columns(cols => {
            cols.Clear();
            cols.AddColumn(prop => prop.Price);
            cols.AddColumn(prop => prop.Discount);
            cols.AddColumn(prop => prop.Profit).Format(f => f.Bold().Italic()).Title("Total");
        })
    );

byte[] excelBytes = exporter.Export();
```

 