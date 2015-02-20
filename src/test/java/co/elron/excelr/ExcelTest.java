package co.elron.excelr;

import java.util.List;

import org.junit.Assert;
import org.junit.Test;

import co.elron.utils.excelr.ExcelReader;
import co.elron.utils.excelr.ExcelReader.RowConverter;

public class ExcelTest {

	public static class Country {
		public String shortCode;
		public String name;

		public Country(String shortCode, String name) {
			this.shortCode = shortCode;
			this.name = name;
		}
	}

	static interface Run {
		void run() throws Exception;
	}

	@Test
	public void test() throws Exception {
		RowConverter<Country> converter = (row) -> new Country(row[0], row[1]);
		//
		ExcelReader<Country> reader = ExcelReader.builder(Country.class).converter(converter).withHeader().csvDelimiter(';').sheets(1).build();
		//
		List<Country> list;
		list = reader.read("src/test/resources/CountryCodes.xlsx");
		checkFirst(list.get(0));
		list = reader.read("src/test/resources/CountryCodes.xls");
		checkFirst(list.get(0));
		list = reader.read("src/test/resources/CountryCodes.csv");
		checkFirst(list.get(0));
	}

	private void checkFirst(Country c) {
		Assert.assertEquals(c.shortCode, "ad");
		Assert.assertEquals(c.name, "Andorra");
	}

	@Test
	public void benchmark() throws Exception {
		RowConverter<Country> converter = (row) -> new Country(row[0], row[1]);
		//
		ExcelReader<Country> reader = ExcelReader.builder(Country.class).converter(converter).csvDelimiter(';').sheets(1).build();
		//
		delta(() -> reader.read("src/test/resources/CountryCodes.xlsx"));
		delta(() -> reader.read("src/test/resources/CountryCodes.xls"));
		delta(() -> reader.read("src/test/resources/CountryCodes.csv"));
		//
	}

	public void delta(Run c) throws Exception {
		long start = System.currentTimeMillis();
		c.run();
		System.out.println("Delta: " + (System.currentTimeMillis() - start));
	}
}
