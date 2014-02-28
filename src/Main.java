import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Main {
	public static void main(String[] args) {

		try {
			String excelPath = "data.xlsx";
			if(args.length > 0)
				excelPath = args[0];
			FileInputStream excelFIS = new FileInputStream(excelPath);
			Workbook excelBook = WorkbookFactory.create(excelFIS);
			// ---------
			System.out
					.println("<?xml version=\"1.0\" encoding=\"UTF-8\"?> \n"
							+ "<!DOCTYPE plist PUBLIC \"-//Apple//DTD PLIST 1.0//EN\" \"http://www.apple.com/DTDs/PropertyList-1.0.dtd\">\n"
							+ "<plist version=\"1.0\">\n" + "<dict>");
			int sheetNum = excelBook.getNumberOfSheets();
			for (int i = 0; i < sheetNum; i++) {
				Sheet sheet = excelBook.getSheetAt(i);
				String sheetName = sheet.getSheetName();

				System.out
						.println("\t<key>" + sheetName + "</key>\n" + "\t<array>");
				//
				int startRow = 0;
				Row row = sheet.getRow(startRow);

				if (row != null) {
					int endRow = sheet.getLastRowNum();

					for (int j = 1; j <= endRow; ++j) {
						HashMap<String, HashMap<String, HashMap<String, String>>> subArray = new HashMap<String, HashMap<String, HashMap<String, String>>>();
						Row row1 = sheet.getRow(j);
						System.out.println("\t\t<dict>");
						for (int k = 0; k < row1.getPhysicalNumberOfCells(); ++k) {
							String key = getCellContent(row.getCell(k));

							if (key.indexOf('-') < 0) {
								System.out.print("\t\t\t<key>"
										+ getCellContent(row.getCell(k))
										+ "</key>\n" + "\t\t\t<string>"
										+ getCellContent(row1.getCell(k))
										+ "</string>\n");
							} else {
								String subValue = getCellContent(row1
										.getCell(k));
								parseSubArray(key, subValue, subArray);
							}
						}

						genSubArray(subArray);


						System.out.println("\t\t</dict>");
					}

					// System.out.println(endRow);
					// System.out.println(endCol);
				}
				System.out.println("\t</array>");
			}
			System.out.println("</dict>");
			System.out.println("</plist>");
			// ---------
			excelFIS.close();
		} catch (Exception e) {
			System.out.println(e);
		}
	}

	public static String getCellContent(Cell cell) {
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_BOOLEAN:
			return cell.getBooleanCellValue() ? "true" : "false";
		case Cell.CELL_TYPE_FORMULA:
			return null;
		case Cell.CELL_TYPE_NUMERIC:
			return "" + cell.getNumericCellValue();
		case Cell.CELL_TYPE_STRING:
			return cell.getStringCellValue();
		default:
			return null;
		}
	}

	public static void parseSubArray(String key, String subValue, HashMap<String, HashMap<String, HashMap<String, String>>> subArray)
	{

		String subArrayKey = key.substring(0,
				key.indexOf('-'));
		String index = key.substring(
						key.indexOf('-'),
						key.lastIndexOf('-'));
		String subKey = key.substring(key
				.lastIndexOf('-') + 1);
		if (subArray.containsKey(subArrayKey)) {
			if(!subArray.get(subArrayKey).containsKey(index))
			{
				HashMap<String, String> value = new HashMap<String, String>();

				subArray.get(subArrayKey).put(index, value);
			}
			subArray.get(subArrayKey).get(index)
					.put(subKey, subValue);
		} else {
			HashMap<String, String> pair = new HashMap<String, String>();
			pair.put(subKey, subValue);
			
			HashMap<String, HashMap<String, String>> dicts = new HashMap<String, HashMap<String, String>>();
			dicts.put(index, pair);

			subArray.put(subArrayKey, dicts);
		}
	}
	
	public static void genSubArray(HashMap<String, HashMap<String, HashMap<String, String>>> subArray)
	{

		for (Iterator<String> iterator = subArray.keySet()
				.iterator(); iterator.hasNext();) {
			String key = (String) iterator.next();

			// rewards
			System.out.println("\t\t\t<key>" + key + "</key>");
			System.out.println("\t\t\t<array>");
			
			HashMap<String, HashMap<String, String>> subDicts = subArray.get(key);
			for (Iterator<String> iterator2 = subDicts.keySet().iterator(); iterator2.hasNext();) {
				String type = (String) iterator2.next();
				HashMap<String, String> subDict = subDicts.get(type);
				
				// items 0....
				System.out.println("\t\t\t\t<dict>");

				for (Iterator<String> iterator3 = subDict.keySet()
						.iterator(); iterator3.hasNext();) {
					String subKey = (String) iterator3.next();
					String subValue = subDict.get(subKey);

					System.out.println("\t\t\t\t\t<key>" + subKey
							+ "</key>\n" + "\t\t\t\t\t<string>"
							+ subValue + "</string>");
				}

				System.out.println("\t\t\t\t</dict>");
			}

			System.out.println("\t\t\t</array>");
		}
	}
}

class Pair {
	String key;
	String value;
}