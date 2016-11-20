#Excel Exporter (PoC)

**Fase preliminar de desarrollo**

Permite exportar un IEnumerable a una hoja excel. Aplicando formato a las filas en base a valores de cada uno de los elementos del IEnumerable.

Puede generar las columnas del excel leyendo los valores del atributo Display (Name, Order...)

Podemos generar una clase con los datos que vamos a mostrar

    public class TestListItem
    {
        [Display(Name = "PropA Disp Name", Order = 2)]
        public string PropA { get; set; }

        [Display(Order = 3)]
        public string PropB { get; set; }

        public int PropC { get; set; }
    }

Creamos un IEnumerable con items de esa clase...

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

Configuramos el excel, creando una hoja, con un nombre y un formato condicional en base a los valores de PropC de cada item.

    var exporter = new Exporter()
		.AddSheet<TestListItem>(c =>
			c.SetData(dataToExport)
			  .Name("Sheet Name")
			  .AddFormat(p => p.PropC == 3, IntNovAction.Utils.ExcelExporter.Utils.FontFormat.Bold)
		);

Por ultimo exportamos el excel

    var result = exporter.Export();
