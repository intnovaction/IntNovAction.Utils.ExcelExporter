# Excel Exporter 
[![NuGet](https://img.shields.io/nuget/v/Intnovaction.Utils.ExcelExporter.svg)](https://www.nuget.org/packages/IntNovAction.Utils.ExcelExporter/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Permite exportar un IEnumerable a una o varias hoja excel. Aplicando formato a las filas en base a valores de cada uno de los elementos del IEnumerable.

Puede generar las columnas del excel leyendo los valores del atributo Display (Name, Order...)

Podemos generar una clase con los datos que vamos a mostrar
```c#
    public class TestListItem
    {
        [Display(Name = "PropA Disp Name", Order = 2)]
        public string PropA { get; set; }

        [Display(Order = 3)]
        public string PropB { get; set; }

        public int PropC { get; set; }
    }
```

Creamos un IEnumerable con items de esa clase...

```c#
    var dataToExport = new List<TestListItem>();
    for (int i = 0; i < 5; i++)
    {
        dataToExport.Add(new TestListItem()
        {
            PropA = $"PropA - {i}",
            PropB = $"PropB - {i}",
            PropC = i,
        });
    }
```

Configuramos el excel, creando una hoja, con un nombre e indicando los datos a pintar.
La librería leerá los metadatos de la clase, creando las columnas con el nombre indicado en los atributos.

```c#
    var exporter = new Exporter()
		.AddSheet<TestListItem>(c => c.Name("Sheet Name").SetData(dataToExport));
```

Por ultimo exportamos el excel

    var result = exporter.Export();

## Crear múltiples hojas

Para crear múltiples hojas con múltiples sets de datos llamamos varias veces a AddSheet:

```c#
    var exporter = new Exporter()
		.AddSheet<TestListItem>(c => c.SetData(dataToExport).Name("Sheet Name"))
        .AddSheet<TestListItem>(c => c.SetData(dataToExport).Name("Sheet 2 Name"));
```

## No pintar la cabecera con los nombres de columnas de la tabla

Para no pintar en la tabla los nombres de las columnas y que empiecen los datos en la coordenadas iniciales.

```c#
    var exporter = new Exporter()
		.AddSheet<TestListItem>(c => c.SetData(dataToExport).Name("Sheet Name").HideColumnHeaders());
        
```

## Establecer coordenadas para pintar la tabla

Podemos especificar que no se empiecen a pintar las filas en A1, sino donde queramos

```c#
var exporter = new Exporter()
    .AddSheet<TestListItem>(c => c.SetData(dataToExport)
        .Name("Sheet Name")
        .SetCoordinates(3, 2)
     );
        
```

## Formatear las filas en base a los valores

Podemos establecer condiciones (reglas de formato a nivel de fila) en forma de expresiones, que apliquen formatos:
- Bold
- Italic
- Underline
- Color
- FontSize

Si el valor de PropC es 3, entonces haz la fila en negrita y cursiva:

```c#
    var exporter = new Exporter()
		.AddSheet<TestListItem>(c => c.SetData(dataToExport)
		  .Name("Sheet Name")
		  .AddFormatRule(p => p.PropC == 3, format => format.Bold().Italic())
		);
```

## Establecer un titulo en la hoja

Podemos establecer un titulo por defecto en la parte superior

```c#
    var exporter = new Exporter()
    .AddSheet<TestListItem>(c => c.SetData(dataToExport).Name("Sheet Name")
            .Title()
    );
```

O personalizarlo con un texto, formato, etc.

```c#
    var exporter = new Exporter()
	.AddSheet<TestListItem>(c => c.SetData(dataToExport).Name("Sheet Name")
            .Title(t => t.Text("Title Text").Format(f => f.Bold()))
    );
```
## Especificar las columnas a mostrar
Mediante la colección columns podemos añadir, ocultar o quitar columnas. Los metadatos del atributo Display de la propiedad (si se utiliza) se tendrán en cuenta.


```c#
    var exporter = new Exporter()
    .AddSheet<TestListItem>(c => c.SetData(dataToExport).Name("Sheet Name")
		.Columns(cols =>
	    {
            cols.Clear(); // Limpiamos todas las columnas autogeneradas
            cols.AddColumn(prop => prop.PropA);            
            cols.AddColumn(prop => prop.PropA).Title("Prop A (2)");
        })
    );
```

Con HideColumn podemos quitar alguna de las columnas autogeneradas. Muy util cuando pasamos un DTO que tiene IDs que no queremos mostrar.

```c#
    var exporter = new Exporter()
    .AddSheet<TestListItem>(c => c.SetData(dataToExport).Name("Sheet Name")
		.Columns(cols =>
	    {
            cols.HideColumn(prop => prop.PropA);
        })
    );
```

Podemos dar formato a una columna, si entra en conflicto con un formato condicional de fila, el de fila tiene preferencia.
Además de los formatos de fila, se puede establecer el ancho de la columna.

```c#
    var exporter = new Exporter()
    .AddSheet<TestListItem>(c => c.SetData(dataToExport).Name("Sheet Name")
		.Columns(cols =>
	    {
            cols.Clear(); // Limpiamos todas las columnas autogeneradas
            cols.AddColumn(prop => prop.PropA);            
            cols.AddColumn(prop => prop.PropB).Format(f => f.Bold().Width(50));           
        })
    );
```

## Modificar los datos a mostrar mediante expresiones
Podemos especificar en lugar del nombre de la columna, una expresión para realizar transformaciones a los datos. Hay que especificar el titulo de la columna.

```c#
    var exporter = new Exporter()
    .AddSheet<TestListItem>(c => c.SetData(dataToExport).Name("Sheet Name")
		.Columns(cols =>
	    {
            cols.Clear(); // Limpiamos todoas las columnas autogeneradas
            cols.AddColumn(prop => prop.PropA); // Mostramos la columna PropA
            cols.AddColumnExpr(prop => prop.PropC + 1, "PropC plus 1"); // Mostramos el contenido de PropC sumándole 1 y lo llamamos PropC plus 1
        })
    );
```

## Imprimir todo en una hoja
Podemos indicar que se impriman todos los datos en una sola hoja:

```c#
var exporter = new Exporter()
    .AddSheet<TestListItem>(c => c.SetData(dataToExport)
        .Name("Sheet Name")
        .PrintInOnePage()
     );
        
```