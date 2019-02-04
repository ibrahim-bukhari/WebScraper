import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.Writer;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Connection;
import org.jsoup.Jsoup;
import org.jsoup.nodes.DataNode;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

public class ExtractData {
	
	//"https://www.fishersci.com/us/en/catalog/search/products?keyword=NC9070936&nav="
	private static final String BASE_URL = "https://www.fishersci.com";
	private static final String SEARCH_URL = "https://www.fishersci.com/us/en/catalog/search/products";
	private static final String SEARCH_KEYWORD = "?keyword=";
	private static final String SEARCH_NAV = "&nav=";
	private static final String INPUT_FILENAME = "Fisher-Batch1.xlsx";
	private static final String OUTPUT_FILENAME = "Fisher_Product_Details.xlsx";
	private static final String NOTFOUND_FILENAME = "Unidentified.xlsx";
	
	private static final String INPUT_COUNTER_FILENAME = "input_counter.txt";
	private static final String OUTPUT_COUNTER_FILENAME = "output_counter.txt";
	private static final String NOTFOUND_COUNTER_FILENAME = "notfound_counter.txt";
	
	private static int inputCounter = 0;
	private static int outputCounter = 0;
	private static int notFoundCounter = 0;
	
	private ArrayList<String> aliasIds = null;
	private ArrayList<String> catalogNos = null;
	ArrayList<Map<String, String>> notFoundList = null;
	ArrayList<Map<String, String>> foundList = null;
	
	private static final String COMPANY_NAME = "Fisher Scientific";
	private static final String COMPANY_NAME_KEY = "Company Name";
	private static final String ALIAS_ID_KEY = "Alias_Id";
	private static final String CATALOG_NUM_KEY = "Catalog_Number";
	private static final String PRODUCT_NAME_KEY = "Product Name";
	private static final String PRODUCT_NUM_KEY = "Product Number";
	private static final String SHORT_DESCRIPTION_KEY = "Short Description";
	private static final String LONG_DESCRIPTION_KEY = "Long Description";
	private static final String MANUFACTURER_INFO_KEY = "Manufacturer Info";
	private static final String MANUFACTURER_NAME_KEY = "Manufacturer Name";
	private static final String MANUFACTURER_NUM_KEY = "Manufacturer Number";
	private static final String PRODUCT_SPECS_KEY = "Product Specifications";
	private static final String PRODUCT_FEATURES_KEY = "Product Features";
	private static final String STATUS_KEY = "Status";
	private static final String STATUS_NOTFOUND = "Not Found";
	private static final String PACKAGING_KEY = "Packaging";
	private static final String NOT_APPLICABLE = "N/A";
	
	
	private void readFileData(String filename) throws IOException {
		FileInputStream file = new FileInputStream(new File(filename));

		//Create Workbook instance holding reference to .xlsx file
		XSSFWorkbook workbook = new XSSFWorkbook(file);

		//Get first/desired sheet from the workbook
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		//Iterate through each row one by one
		Iterator<Row> rowIterator = sheet.iterator();
		int aliasIndex = 0;
		int catalogIndex = 0;
		
		while (rowIterator.hasNext()) 
		{
			Row row = rowIterator.next();
			//For each row, iterate through all the columns
			Iterator<Cell> cellIterator = row.cellIterator();
			if (row.getRowNum() == 0) {
				while (cellIterator.hasNext()) 
				{
					Cell cell = cellIterator.next();				
					if (cell.getStringCellValue().compareToIgnoreCase(ALIAS_ID_KEY) == 0) {
						aliasIndex = cell.getColumnIndex();
					}
					else if (cell.getStringCellValue().compareToIgnoreCase(CATALOG_NUM_KEY) == 0) {
						catalogIndex = cell.getColumnIndex();
					}
				}
			}
			else {
				break;
			}
		}
		aliasIds = getColumnData(aliasIndex, sheet);
		catalogNos = getColumnData(catalogIndex, sheet);
		System.out.println("Catalog Number extraction done!");
	}
	
	
	
	
	private void parseItemList(String catalogNumber, int currentIndex) throws IOException {
		ArrayList<String> links = searchItem(catalogNumber);
		if(links == null) {
			//item not found
			//System.out.println(catalogNumber + " not found!");
			addToNotFound(aliasIds.get(currentIndex), catalogNumber, STATUS_NOTFOUND);
		}
		else {
			//Intermediate page
			int listSize = links.size();
			String productURL = "";
			if(listSize > 0) {
				boolean productFound = false;
				for (int i = 0; i < listSize; i++) {
					try {
						productURL = BASE_URL + links.get(i);
						Document doc = Jsoup.connect(productURL).timeout(10000).get();
						String pId = getProductId(doc);
						String pIdSpaces = pId.replace("-", " ");
						String pIdNoSpaces = pId.replace("-", "");
						if (pId.compareToIgnoreCase(catalogNumber) == 0 || pIdSpaces.compareToIgnoreCase(catalogNumber) == 0 || pIdNoSpaces.compareToIgnoreCase(catalogNumber) == 0) {
							//this is the product we want. capture information for this. ignore others links
							productFound = true;
							String name_desc = getProductName(doc);
							String manufacturerInfo = getManufacturerInfo(doc);
							String manufacturerName = getManufacturerName(manufacturerInfo);
							String manufacturerNumber = getManufacturerNumber(manufacturerInfo);
							addToFound(COMPANY_NAME, aliasIds.get(currentIndex), catalogNumber, name_desc, name_desc, getProductLongDescription(doc), getProductPackaging(doc), manufacturerInfo, manufacturerName, manufacturerNumber, getProductSpecifications(doc), getProductFeatures(doc));
							break;
						}
					} catch (org.jsoup.HttpStatusException e) {
						System.out.println("received error for URL : " + productURL);
					}
					
				}
				if (!productFound) {
					System.out.println("Product not found in list of links");
					addToNotFound(aliasIds.get(currentIndex), catalogNumber, STATUS_NOTFOUND);
				}
			}
			else {
				//direct product page
				productURL = SEARCH_URL + SEARCH_KEYWORD + catalogNumber + SEARCH_NAV;
				try {
					Document doc = Jsoup.connect(productURL).timeout(10000).get();
					String name_desc = getProductName(doc);
					String manufacturerInfo = getManufacturerInfo(doc);
					String manufacturerName = getManufacturerName(manufacturerInfo);
					String manufacturerNumber = getManufacturerNumber(manufacturerInfo);
					
					addToFound(COMPANY_NAME, aliasIds.get(currentIndex), catalogNumber, name_desc, name_desc, getProductLongDescription(doc), getProductPackaging(doc), manufacturerInfo, manufacturerName, manufacturerNumber, getProductSpecifications(doc), getProductFeatures(doc));
				} catch(org.jsoup.HttpStatusException e) {
					System.out.println("received error for URL : " + productURL);
				}
				
			}
		}
	}
	
	private ArrayList<String> getColumnData(int index, XSSFSheet sheet) {
		ArrayList<String> ids = new ArrayList<String>();
		//Iterate through each row one by one
		Iterator<Row> rowIterator = sheet.iterator();
		while (rowIterator.hasNext()) 
		{
			Row row = rowIterator.next();
			//For each row, iterate through all the columns
			if (row.getRowNum() > 0) {
				Cell cell = row.getCell(index);
				String str = null;
				if (cell != null) {
		    		switch (cell.getCellType()) 
					{
						case Cell.CELL_TYPE_NUMERIC:
							str = fmt(cell.getNumericCellValue());
							break;
						case Cell.CELL_TYPE_STRING:
							str = cell.getStringCellValue();
							break;
					}
		    		ids.add(str);
				}
			}
		}
		return ids;
	}
	
	private String getProductSpecifications (Document doc) {
		String str = NOT_APPLICABLE;
		Elements specs = null;
		Element specTable = null;
		specs = doc.select("div[id=spec_and_desc]");
		if(specs != null) {
			String tempSpecs = "";
			specTable = specs.select("table[class=specs_table_full]").first();
			if (specTable != null) {
				Elements items = specTable.select("tr");
				for (int i = 0; i < items.size(); i++) {
					Element e = items.get(i);
					tempSpecs += e.text();
					tempSpecs += " ";
				}
				str = tempSpecs;
			}
		}
		str = str.trim();	
		return str;
	}
	
	private String getProductFeatures (Document doc) {
		String str = NOT_APPLICABLE;
		Elements specs = null;
		Elements lFeatures = null;
		Elements pFeatures = null;
		specs = doc.select("div[id=spec_and_desc]");
		if(specs != null) {
			String tempSpecs = "";
			lFeatures = specs.select("li");
			pFeatures = specs.select("p");
			if (lFeatures != null) {
				tempSpecs += lFeatures.text();
			}
			tempSpecs += "  ";
			if (pFeatures != null) {
				tempSpecs += pFeatures.text();
			}
			str = tempSpecs;
		}
		str = str.trim();
		return str;
	}
	
	private String getProductId(Document doc) {
		String str = NOT_APPLICABLE;
		Elements idDetails = doc.select("div[id=SKUHighlightContainer]");
		Element productId = idDetails.select("span").first();
		if (productId != null) {
			str = productId.text();
		}
		return str;
	}
	
	private String getProductPackaging(Document doc) {
		String str = NOT_APPLICABLE;
		Elements priceDetails = doc.select("label[class=price]");
		Element packaging = priceDetails.select("span").first();
		if (packaging == null) {
			Elements promoDetails = doc.select("div[class=promo_price]");
			Element promoPrice = promoDetails.select("span").first();
			if (promoPrice != null) {
				str = promoPrice.text();
				str = str.replace("/", "");
				str = str.trim();
			}
			else {
				System.out.println("PROMO PRICE NOT FOUND!");
			}
		}
		else {
			str = packaging.text();
			str = str.replace("/", "");
			str = str.trim();
		}
		return str;
	}
	
	private String getProductName(Document doc) {
		String str = NOT_APPLICABLE;
		Elements productDetails = doc.select("div[id=ProductDescriptionContainer]");
		Element productName = productDetails.select("h1").first();
		if (productName != null) {
			str = productName.text();
			str = str.trim();
		}
		return str;
	}
	
	private String getManufacturerInfo(Document doc) {
		String str = NOT_APPLICABLE;
		Elements productDetails = doc.select("div[id=ProductDescriptionContainer]");
		Elements details = productDetails.select("div[class=subhead]");
		Elements pElements = details.select("p");
		for (int i = 0; i < pElements.size(); i++) {
			Element e = pElements.get(i);
			str = e.text();
			if(str.contains("Manufacturer:")) {
				// Remove "Manufacturer: " from start of string
				str = str.replace("Manufacturer:", "");
				break;
			}
		}
		str = str.trim();
		return str;
	}
	private String getManufacturerName(String info) {
		String str = NOT_APPLICABLE;		
		String[] data = info.split("\\s+");
		String tempManuName = "";
		for (int j = 0; j < data.length - 1; j++) {
			tempManuName += data[j];
			tempManuName += " ";
		}
		tempManuName = tempManuName.trim();
		str = tempManuName;
		str = str.trim();
		return str;
	}

	private String getManufacturerNumber(String info) {
		String str = NOT_APPLICABLE;
		String[] data = info.split("\\s+");
		String tempManuNum = data[data.length-1];
		str = tempManuNum;
		str = str.trim();
		return str;
	}
	
	
	private String getProductLongDescription(Document doc) {
		String str = NOT_APPLICABLE;
		Elements productDetails = doc.select("div[id=ProductDescriptionContainer]");
		Elements subHeadings = productDetails.select("div[class=subhead]");
		Elements blockHeadings = productDetails.select("div[class=block_head]");
		Elements warnings = productDetails.select("div[class=warning-msg-container]");
		Elements pElements = productDetails.select("p");
		
		String fullDetails = "";
		for (int i = 0; i < pElements.size(); i++) {
			Element e = pElements.get(i);
			fullDetails += e.text();
			fullDetails += " ";
		}
		fullDetails = fullDetails.trim();
		
		ArrayList <String> stringsToRemove = new ArrayList<String>();
		for (int i = 0; i < warnings.size(); i++) {
			Element e = warnings.get(i);
			stringsToRemove.add(e.text());
		}
		for (int i = 0; i < blockHeadings.size(); i++) {
			Element e = blockHeadings.get(i);
			stringsToRemove.add(e.text());
		}
		for (int i = 0; i < subHeadings.size(); i++) {
			Element e = subHeadings.get(i);
			stringsToRemove.add(e.text());
		}
		
		for (int i = 0; i < stringsToRemove.size(); i++) {
			fullDetails = fullDetails.replace(stringsToRemove.get(i), "");
		}
		str = fullDetails;
		str = str.trim();
		
		return str;
	}
	
	
	private ArrayList<String> searchItem(String catalogNumber) throws IOException {
		String searchURL = SEARCH_URL + SEARCH_KEYWORD + catalogNumber + SEARCH_NAV;
		ArrayList<String> sublinks = null;
		try {
			Document doc = Jsoup.connect(searchURL).timeout(10000).get();
			Elements errorElements = doc.select("div[class=search_results_error_message]");
			if (errorElements.size() > 0) {
				return null;
			}
			Elements scriptElements = doc.select("script");
			sublinks = new ArrayList<String>();
			for (int i = 0; i < scriptElements.size(); i++) {
				Element element = scriptElements.get(i);
				for (int j = 0; j < element.dataNodes().size(); j++) {
					DataNode node = element.dataNodes().get(j);
					String str = node.getWholeData();
					Pattern pattern = Pattern.compile("productUrl(.*?)promoUrl");
					Matcher matcher = pattern.matcher(str);
					while (matcher.find()) {
						String link = matcher.group(1);
						link = link.substring(3);
						link = link.substring(0, link.length() - 3);
						//System.out.println(link);
						sublinks.add(link);
					}
				}
			}
		} catch (org.jsoup.HttpStatusException e) {
			System.out.println("received error for URL : " + searchURL);
			return sublinks;
		}
		return sublinks;
	}
	
	private void initializeLists() throws IOException {
		notFoundCounter = getCounter(NOTFOUND_COUNTER_FILENAME);
		inputCounter = getCounter(INPUT_COUNTER_FILENAME);
		outputCounter = getCounter(OUTPUT_COUNTER_FILENAME);
		
		
		notFoundList = new ArrayList<Map<String, String>>();
		foundList = new ArrayList<Map<String, String>>();
		
		File varTmpDir = new File(NOTFOUND_FILENAME);
		boolean exists = varTmpDir.exists();
		if (!exists) {
			addToNotFound(ALIAS_ID_KEY, PRODUCT_NUM_KEY, STATUS_KEY);
		}
		
		File varTmpDir_1 = new File(OUTPUT_FILENAME);
		boolean exists_1 = varTmpDir_1.exists();
		if (!exists_1) {
			addToFound(COMPANY_NAME, ALIAS_ID_KEY, PRODUCT_NUM_KEY, PRODUCT_NAME_KEY, SHORT_DESCRIPTION_KEY, LONG_DESCRIPTION_KEY, PACKAGING_KEY, MANUFACTURER_INFO_KEY, MANUFACTURER_NAME_KEY, MANUFACTURER_NUM_KEY, PRODUCT_SPECS_KEY, PRODUCT_FEATURES_KEY);
		}
	}
	
	private int getCounter(String filename) throws IOException {
		int index = 0;
		
		// This will reference one line at a time
        String line = null;
        
		// FileReader reads text files in the default encoding.
        FileReader fileReader = new FileReader(filename);

        // Always wrap FileReader in BufferedReader.
        BufferedReader bufferedReader = new BufferedReader(fileReader);

        while((line = bufferedReader.readLine()) != null) {
            index = Integer.valueOf(line);
        }   

        // Always close files.
        bufferedReader.close();
	
		return index;
	}
	
	private void setCounter(int index, String filename) throws IOException {
		Writer wr = new FileWriter(filename);
		index = index + 1;
		wr.write(Integer.toString(index++));
		wr.close();
	}
	
	private void addToNotFound(String aliasId, String productNo, String status) throws IOException {
		writeNotFoundItems(aliasId, productNo, status);
	}
	
	private void addToFound(String companyName, String aliasId, String productNo, String productName, String sdescription, String ldescription, String packaging, String manufacturerInfo, String manufacturerName, String manufacturerNo, String productSpecs, String productFeatures) throws IOException {
		writeFoundItems(companyName, aliasId, productNo, productName, sdescription, ldescription, packaging, manufacturerInfo, manufacturerName, manufacturerNo, productSpecs, productFeatures);
	}
	
	private void writeNotFoundItems(String aliasId, String productNo, String status) throws IOException {
		
		//Workbook
		XSSFWorkbook workbook = null;
		
		//Worksheet
		XSSFSheet sheet = null;
		
		//This data will be written (Object[])
		Map<String, Object[]> excelData = new TreeMap<String, Object[]>();
		
		notFoundCounter = getCounter(NOTFOUND_COUNTER_FILENAME);
		
		File varTmpDir = new File(NOTFOUND_FILENAME);
		boolean exists = varTmpDir.exists();
		if (!exists) {
			workbook = new XSSFWorkbook();
			//Create a blank sheet
			sheet = workbook.createSheet("Not Found");
			
			//Setup headings for columns
			excelData.put("1", new Object[] {aliasId, productNo, status});
		}
		else {
			//Read the spreadsheet that needs to be updated
			FileInputStream fsIP= new FileInputStream(new File(NOTFOUND_FILENAME));
			
			//Access the workbook
			workbook = new XSSFWorkbook(fsIP); 
			
			//Get existing sheet
			sheet = workbook.getSheet("Not Found");
			
			excelData.put(Integer.toString(notFoundCounter+1), new Object[] {aliasId, productNo, status});
		}
		
		//Iterate over data and write to sheet
		int rownum = notFoundCounter;
		Set<String> keyset = excelData.keySet();
		for (String key : keyset)
		{
			Row row = sheet.createRow(rownum++);
			Object [] objArr = excelData.get(key);
			int cellnum = 0;
			for (Object obj : objArr)
			{
				Cell cell = row.createCell(cellnum++);
				if(obj instanceof String)
					cell.setCellValue((String)obj);
				else if(obj instanceof Integer)
					cell.setCellValue((Integer)obj);
			}
		}
		setCounter(notFoundCounter, NOTFOUND_COUNTER_FILENAME);
		
		FileOutputStream out = new FileOutputStream(new File(NOTFOUND_FILENAME));
		workbook.write(out);
		out.close();

		//System.out.println(NOTFOUND_FILENAME + " written successfully on disk.");
	}
	
	private void writeFoundItems(String companyName, String aliasId, String productNo, String productName, String sdescription, String ldescription, String packaging, String manufacturerInfo, String manufacturerName, String manufacturerNo, String productSpecs, String productFeatures) throws IOException {
		//Workbook
		XSSFWorkbook workbook = null;

		//Worksheet
		XSSFSheet sheet = null;

		//This data will be written (Object[])
		Map<String, Object[]> excelData = new TreeMap<String, Object[]>();

		outputCounter = getCounter(OUTPUT_COUNTER_FILENAME);

		File varTmpDir = new File(OUTPUT_FILENAME);
		boolean exists = varTmpDir.exists();
		if (!exists) {
			workbook = new XSSFWorkbook();
			//Create a blank sheet
			sheet = workbook.createSheet("Product Data");

			//Setup headings for columns
			excelData.put("1", new Object[] {companyName, aliasId, productNo, productName, sdescription, ldescription, packaging, manufacturerInfo, manufacturerName, manufacturerNo, productSpecs, productFeatures});
		}
		else {
			//Read the spreadsheet that needs to be updated
			FileInputStream fsIP= new FileInputStream(new File(OUTPUT_FILENAME));

			//Access the workbook
			workbook = new XSSFWorkbook(fsIP); 

			//Get existing sheet
			sheet = workbook.getSheet("Product Data");

			excelData.put(Integer.toString(outputCounter+1), new Object[] {companyName, aliasId, productNo, productName, sdescription, ldescription, packaging, manufacturerInfo, manufacturerName, manufacturerNo, productSpecs, productFeatures});
		}



		//Iterate over data and write to sheet
		int rownum = outputCounter;
		Set<String> keyset = excelData.keySet();
		for (String key : keyset)
		{
			Row row = sheet.createRow(rownum++);
			Object [] objArr = excelData.get(key);
			int cellnum = 0;
			for (Object obj : objArr)
			{
				Cell cell = row.createCell(cellnum++);
				if(obj instanceof String)
					cell.setCellValue((String)obj);
				else if(obj instanceof Integer)
					cell.setCellValue((Integer)obj);
			}
		}
		setCounter(outputCounter, OUTPUT_COUNTER_FILENAME);

		FileOutputStream out = new FileOutputStream(new File(OUTPUT_FILENAME));
		workbook.write(out);
		out.close();

		//System.out.println(OUTPUT_FILENAME + " written successfully on disk.");
	}
	
	public static String fmt(double d)
	{
	    if(d == (long) d)
	        return String.format("%d",(long)d);
	    else
	        return String.format("%s",d);
	}
	
	public void run() throws IOException {
		
		long startTime = System.currentTimeMillis();

		
		
		readFileData(INPUT_FILENAME);
		initializeLists();
		
//		String catalogNumber = "501003312"; //intermediate links
		//22505005  50175219  50175046  	//No price
//		String catalogNumber = "S20681";
//		String catalogNumber = "BDB553998";  //Features and specs found
//		String catalogNumber = "2939";  	//Too many links
//		String catalogNumber = "NC0760647"; //Direct product 
//		String catalogNumber = "14-387-967" //Manufacturer test
		
		//10 451 208PR 	//No pid
		
		inputCounter = getCounter(INPUT_COUNTER_FILENAME);
		for (int i = inputCounter; i < catalogNos.size(); i++) {
			System.out.println("Parsing item # " + i + " with product # " + catalogNos.get(i));
			parseItemList(catalogNos.get(i), i);
			setCounter(i, INPUT_COUNTER_FILENAME);
		}
		
//		parseItemList(catalogNumber, 73);
		System.out.println("All Items Parsed");
		long endTime = System.currentTimeMillis();
		long seconds = (endTime - startTime)/1000;
		System.out.println("That took " + seconds + " seconds or " + seconds/60 + "minutes");
	}
}
