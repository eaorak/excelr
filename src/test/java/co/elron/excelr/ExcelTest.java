package co.elron.excelr;

import co.elron.utils.excelr.ExcelReader;
import co.elron.utils.excelr.ExcelReader.RowConverter;
import org.junit.Before;
import org.junit.Test;

import java.util.List;

import static org.junit.Assert.assertEquals;

public class ExcelTest {

	public static class Country {
		public String shortCode;
		public String name;

		public Country(String shortCode, String name) {
			this.shortCode = shortCode;
			this.name = name;
		}
	}

	interface Run {
		void run() throws Exception;
	}

	private ExcelReader<Country> reader;

	@Before
	public void setUp() {
		RowConverter<Country> converter = (row) -> new Country((String) row[0], (String) row[1]);
		reader = ExcelReader.builder(Country.class)
				            .converter(converter)
				            .withHeader()
				            .csvDelimiter(';')
				            .sheets(1)
				            .build();
	}

	@Test
	public void shouldParseCorrectly_GivenXlsxFile() throws Exception {
		List<Country> list;
		list = reader.read("src/test/resources/CountryCodes.xlsx");
		checkList(list);
	}

	@Test
	public void shouldParseCorrectly_GivenXlsFile() throws Exception {
		List<Country> list;
		list = reader.read("src/test/resources/CountryCodes.xls");
		checkList(list);
	}

	@Test
	public void shouldParseCorrectly_GivenCsvFile() throws Exception {
		List<Country> list;
		list = reader.read("src/test/resources/CountryCodes.csv");
		checkList(list);
	}

	@Test
	public void shouldHandleNullValues_GivenANullCell() throws Exception {
		List<Country> list;
		list = reader.read("src/test/resources/CountryCodes.xls");
		Country country = list.get(1);
		assertEquals(null, country.name);
		assertEquals("ae", country.shortCode);
	}

	private void checkList(List<Country> list) {
		assertEquals(252, list.size());
		assertEquals(list.get(0).shortCode, "ad");
		assertEquals(list.get(0).name, "Andorra");
		//
		assertEquals(list.get(56).shortCode, "dj");
		assertEquals(list.get(56).name, "Djibouti");
		//
		assertEquals(list.get(243).shortCode, "wf");
		assertEquals(list.get(243).name, "Wallis and Futuna Islands");
	}

	@Test
	public void benchmark() throws Exception {
		RowConverter<Country> converter = (row) -> new Country((String) row[0], (String) row[1]);
		//
		ExcelReader<Country> reader = ExcelReader.builder(Country.class).converter(converter).csvDelimiter(';').sheets(1).build();
		//
		delta(() -> reader.read("src/test/resources/CountryCodes.xlsx"));
		delta(() -> reader.read("src/test/resources/CountryCodes.xls"));
		delta(() -> reader.read("src/test/resources/CountryCodes.csv"));
	}

	public void delta(Run c) throws Exception {
		long start = System.currentTimeMillis();
		c.run();
		System.out.println("Delta: " + (System.currentTimeMillis() - start));
	}
}
