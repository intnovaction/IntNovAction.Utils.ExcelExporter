#Excel Exporter (PoC)

**Fase preliminar de desarrollo**

Permite exportar un IEnumerable a una hoja excel. Aplicando formato a las filas en base a valores de cada uno de los elementos del IEnumerable.

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

````c#
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

Configuramos el excel, creando una hoja, con un nombre.

```c#
    var exporter = new Exporter()
		.AddSheet<TestListItem>(c => c.SetData(dataToExport).Name("Sheet Name"));
```

Por ultimo exportamos el excel

    var result = exporter.Export();

##Crear múltiples hojas##

Para crear múltiples hojas con múltiples sets de datos llamamos varias veces a AddSheet:

```c#
    var exporter = new Exporter()
		.AddSheet<TestListItem>(c => c.SetData(dataToExport).Name("Sheet Name"))
        .AddSheet<TestListItem>(c => c.SetData(dataToExport).Name("Sheet 2 Name"));
```

##No pintar la cabecera con los nombres de columnas de la tabla##

Para no pintar en la tabla los nombres de las columnas y que empiecen los datos en la coordenadas iniciales.

```c#
    var exporter = new Exporter()
		.AddSheet<TestListItem>(c => c.SetData(dataToExport).Name("Sheet Name").HideHeaders());
        
```

##Establecer coordenadas para pintar la tabla##

Podemos especificar que no se empiecen a pintar las filas en A1, sino donde queramos

```c#
    var exporter = new Exporter()
		.AddSheet<TestListItem>(c => c.SetData(dataToExport).Name("Sheet Name").SetCoordinates(3, 2));
        
```

##Formatear las filas en base a los valores##

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