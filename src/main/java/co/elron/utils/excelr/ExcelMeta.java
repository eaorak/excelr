package co.elron.utils.excelr;

import java.util.HashMap;
import java.util.Map;

/**
 * Created by ender on 04/02/15.
 */
public class ExcelMeta {

	public enum FileType {
		EXCEL, TEXT, VCARD;
	}

	// Ex: NAME:0, MSISDN:1, EMAIL:2, ADDRESS:3 ...
	public Map<String, Integer> fieldMap = new HashMap<>();
	public FileType type;
	public boolean hasHeader;

	public static ExcelMeta defaultMeta() {
		ExcelMeta meta = new ExcelMeta();
		meta.fieldMap.put("NAME", 0);
		meta.fieldMap.put("MSISDN", 1);
		meta.fieldMap.put("EMAIL", 2);
		meta.fieldMap.put("ADDRESS", 3);
		meta.type = FileType.EXCEL;
		return meta;
	}

}
