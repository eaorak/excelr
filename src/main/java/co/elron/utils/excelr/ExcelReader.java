package co.elron.utils.excelr;

import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import au.com.bytecode.opencsv.CSVReader;

public class ExcelReader<T> {

	private static final int MAX_SHEET = 1;
	private static final char CSV_DELIM = ';';

	public static interface RowConverter<T> {
		public T convert(String[] row);
	}

	public static class Builder<T> {
		private boolean hasHeader;
		private RowConverter<T> converter;
		private int sheets;
		private char delimiter = CSV_DELIM;

		public Builder() {
		}

		public Builder<T> converter(RowConverter<T> converter) {
			this.converter = converter;
			return this;
		}

		public Builder<T> withHeader() {
			this.hasHeader = true;
			return this;
		}

		public Builder<T> sheets(int sheetCount) {
			this.sheets = sheetCount;
			return this;
		}

		public Builder<T> csvDelimiter(char delimiter) {
			this.delimiter = delimiter;
			return this;
		}

		public ExcelReader<T> build() {
			return new ExcelReader<T>(this);
		}

	}

	private Builder<T> info;

	public static <T> Builder<T> builder(Class<T> cls) {
		return new Builder<T>();
	}

	private ExcelReader(Builder<T> info) {
		this.info = info;
	}

	public List<T> read(String fileName) throws IOException, InvalidFormatException {
		try (FileInputStream is = new FileInputStream(fileName)) {
			return read(is);
		}
	}

	public List<T> read(InputStream is) throws IOException, InvalidFormatException {
		List<T> objList = null;
		try (BufferedInputStream buf = new BufferedInputStream(is)) {
			if (isExcel(buf)) { // XLSX, XLS
				objList = readExcel(buf);
			} else { // CSV
				objList = readCsv(buf);
			}
		}
		return objList;
	}

	private List<T> readExcel(InputStream is) throws InvalidFormatException, IOException {
		Workbook workbook = WorkbookFactory.create(is);
		int sheetCount = Math.min(workbook.getNumberOfSheets(), MAX_SHEET);
		List<T> objList = new ArrayList<>();
		sheetCount = (info.sheets == 0) ? sheetCount : info.sheets;
		for (int i = 0; i < sheetCount; i++) {
			Sheet sheet = workbook.getSheetAt(i);
			extractSheet(sheet, objList);
		}
		return objList;
	}

	private List<T> readCsv(InputStream in) throws IOException {
		List<T> objList = new ArrayList<>();
		InputStreamReader isr = new InputStreamReader(in);
		try (CSVReader cvsr = new CSVReader(isr, info.delimiter)) {
			List<String[]> allRows = cvsr.readAll();
			int start = info.hasHeader ? 1 : 0;
			for (int i = start; i < allRows.size(); i++) {
				T obj = info.converter.convert(allRows.get(i));
				objList.add(obj);
			}
		}
		return objList;
	}

	private void extractSheet(Sheet sheet, List<T> objList) {
		Iterator<Row> rowIterator = sheet.iterator();
		if (rowIterator.hasNext() && info.hasHeader) {
			rowIterator.next();
		}
		while (rowIterator.hasNext()) {
			T obj = extractObject(rowIterator);
			objList.add(obj);
		}
	}

	private T extractObject(Iterator<Row> rowIterator) {
		Row row = rowIterator.next();
		Iterator<Cell> cellIterator = row.cellIterator();
		String[] rowVals = new String[row.getLastCellNum()];
		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			rowVals[cell.getColumnIndex()] = getValue(cell);
		}
		return info.converter.convert(rowVals);
	}

	private boolean isExcel(InputStream is) throws IOException {
		return POIXMLDocument.hasOOXMLHeader(is) /* .xlsx */
				|| POIFSFileSystem.hasPOIFSHeader(is); /* .xls */
	}

	private String getValue(Cell cell) {
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_STRING:
			return cell.getStringCellValue();
		case Cell.CELL_TYPE_NUMERIC:
			return String.valueOf(cell.getNumericCellValue());
		}
		return null;
	}

}
