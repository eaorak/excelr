# excelr
Basic utility in Java that reads both XLSX, XLS and CSV files and converts Excel rows to Java Objects. It uses Apache POI and OpenCSV for reading Excel files.

Usage example:

		RowConverter<Country> converter = (Object[] row) -> new Country((String)row[0], (String)row[1]);
		
		ExcelReader<Country> reader = ExcelReader.builder(Country.class)
		     .converter(converter)
		     .withHeader()
		     .csvDelimiter(';')
		     .sheets(1)
		     .build();
		
		List<Country> list;
		list = reader.read("src/test/resources/CountryCodes.xlsx");
		list = reader.read("src/test/resources/CountryCodes.xls");
		list = reader.read("src/test/resources/CountryCodes.csv");
		

Excel File: 
		
		Code	Country
		ad	Andorra
		ae	United Arab Emirates
		af	Afghanistan
		ag	Antigua and Barbuda
		...
		

Country Class:
		
	public static class Country {
		public String shortCode;
		public String name;

		public Country(String shortCode, String name) {
			this.shortCode = shortCode;
			this.name = name;
		}
	}
