package util;

import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.OptionalDouble;
import java.util.Set;
import java.util.TreeSet;
import java.util.function.Function;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import data.SpeciesProperty;

public class Utils {

    private static final int COLUMN_INDEX_SPECIES_NAME = 0;

    private static XSSFWorkbook writeWorkbook = new XSSFWorkbook();

    private static XSSFCellStyle emptyCellStyle;
    private static XSSFCellStyle propertyNameCellStyle;
    private static XSSFCellStyle refNameCellStyle;
    private static XSSFCellStyle refCellStyle;

    public static void main(String[] args) throws Exception {

	List<SpeciesProperty> speciesProperties = new ArrayList<>();
	File folder = new File("tables");
	checkTableFolder(folder);


	for (File file : folder.listFiles()) {
	    speciesProperties.addAll(readSpeciesProperties(file));
	}
	//
	Map<String, Map<String, List<SpeciesProperty>>> summary = createSummary(speciesProperties);

	String fileName = "gesamt.xlsx";

	initWorkbookStyle();
	writeSummary(summary, fileName);

	Desktop.getDesktop().open(new File(fileName));

    }

    private static void checkTableFolder(File folder) {
    	if(folder.exists()) return;
    	else{
    		throw new RuntimeException("Ordner tables nicht gefunden. Ordner tables muss zuvor angelegt werden.");
    	}
		
	}

	private static void initWorkbookStyle() {
	emptyCellStyle = writeWorkbook.createCellStyle();
	emptyCellStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(234, 234, 234)));
	emptyCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

	propertyNameCellStyle = writeWorkbook.createCellStyle();
	propertyNameCellStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(180, 198, 231)));
	propertyNameCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	XSSFFont font = writeWorkbook.createFont();
	font.setBold(true);
	propertyNameCellStyle.setFont(font);

	refNameCellStyle = writeWorkbook.createCellStyle();
	refNameCellStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(180, 198, 231)));
	refNameCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	refNameCellStyle.setAlignment(HorizontalAlignment.CENTER);
	XSSFFont fontRefName = writeWorkbook.createFont();
	fontRefName.setFontHeightInPoints((short) 9);
	fontRefName.setItalic(true);
	refNameCellStyle.setFont(fontRefName);

	refCellStyle = writeWorkbook.createCellStyle();
	XSSFFont fontRef = writeWorkbook.createFont();
	fontRef.setFontHeightInPoints((short) 9);
	fontRef.setItalic(true);
	refCellStyle.setAlignment(HorizontalAlignment.CENTER);
	refCellStyle.setFont(fontRef);

    }

    private static void writeSummary(Map<String, Map<String, List<SpeciesProperty>>> summary, String fileName)
	    throws Exception {

	Sheet sheet = writeWorkbook.createSheet("Zusammenfassung");

	TreeSet<String> propertyNames = getAllProperties(summary);
	writeHeader(sheet, propertyNames);
	writeBody(sheet, summary, propertyNames);

	FileOutputStream fileOut = new FileOutputStream(fileName);
	writeWorkbook.write(fileOut);
	fileOut.close();
    }

    private static void writeBody(Sheet sheet, Map<String, Map<String, List<SpeciesProperty>>> summary,
	    TreeSet<String> propertyNames) throws Exception {

	List<String> speciesNameList = new ArrayList<>(new TreeSet<>(summary.keySet()));
	List<String> propertyList = new ArrayList<>(propertyNames);

	for (int i = 0; i < speciesNameList.size(); i++) {
	    Row createRow = sheet.createRow(i + 1);
	    Cell speciesCell = createRow.createCell(0);
	    String keySpecies = speciesNameList.get(i);
	    speciesCell.setCellValue(keySpecies);

	    for (int j = 1, k = 0; k < propertyNames.size(); j += 2, k++) {
		Cell valueCell = createRow.createCell(j);
		Cell refCell = createRow.createCell(j + 1);
		// get all properties of the current species (row)
		Map<String, List<SpeciesProperty>> map = summary.get(keySpecies);
		// get the only the properties for the current property (column)
		String key = propertyList.get(k);
		List<SpeciesProperty> list = map.get(key);
		if (list == null || list.isEmpty()) {
		    handleEmptyCell(valueCell, refCell);
		} else {
		    refCell.setCellStyle(refCellStyle);
		    setMergedProperties(valueCell, refCell, list);
		}
	    }
	}

	short lastCellNum = sheet.getRow(0).getLastCellNum();
	for (int i = 0; i < lastCellNum; i++) {
	    sheet.autoSizeColumn(i);
	}
    }

    private static void handleEmptyCell(Cell valueCell, Cell refCell) {
	valueCell.setCellStyle(emptyCellStyle);
	refCell.setCellStyle(emptyCellStyle);
	valueCell.setCellValue("");
	refCell.setCellValue("");
    }

    private static void setMergedProperties(Cell valueCell, Cell refCell, List<SpeciesProperty> list) {

	LinkedHashSet<String> references = new LinkedHashSet<>();

	Set<Double> usedDoubles = new TreeSet<>();
	Set<String> usedStrings = new TreeSet<>();

	List<SpeciesProperty> myList = new ArrayList<>(list);
	myList.sort((s1, s2) -> s1.reference.compareTo(s2.reference));

	for (int i = 0; i < myList.size(); i++) {
	    if (myList.get(i).doubleValue != null) {
		usedDoubles.add(myList.get(i).doubleValue);
	    } else if (myList.get(i).stringValue != null) {
		usedStrings.add(myList.get(i).stringValue);
	    }
	    references.add(myList.get(i).reference.toString());
	}

	if (!usedDoubles.isEmpty() && !usedStrings.isEmpty()) {
	    throw new RuntimeException("found string and double values ");
	}

	if (!usedDoubles.isEmpty()) {
	    OptionalDouble average = usedDoubles.stream().mapToDouble(d -> d.doubleValue()).average();
	    valueCell.setCellValue(round(average.getAsDouble(), 3));
	} else {
	    String stringValues = usedStrings.stream().collect(Collectors.joining(","));
	    valueCell.setCellValue(stringValues);
	}

	String refValues = references.stream().collect(Collectors.joining(","));
	refCell.setCellValue(refValues);
    }

    private static void writeHeader(Sheet sheet, Set<String> propertyNames) {

	List<String> propertyList = new ArrayList<>(propertyNames);

	Row row = sheet.createRow(0);
	for (int i = 0, columnIndex = 1; i < propertyList.size(); i++, columnIndex += 2) {
	    Cell propertyColumn = row.createCell(columnIndex, CellType.STRING);
	    Cell referenceColumn = row.createCell(columnIndex + 1, CellType.STRING);
	    propertyColumn.setCellValue(propertyList.get(i));
	    propertyColumn.setCellStyle(propertyNameCellStyle);
	    referenceColumn.setCellValue("ref");
	    referenceColumn.setCellStyle(refNameCellStyle);

	}

    }

    private static TreeSet<String> getAllProperties(Map<String, Map<String, List<SpeciesProperty>>> summary) {

	TreeSet<String> properties = new TreeSet<>();
	for (String speciesName : summary.keySet()) {
	    Set<String> propertiesOfSpecies = summary.get(speciesName).keySet();
	    properties.addAll(propertiesOfSpecies);
	}
	return properties;
    }

    private static Map<String, Map<String, List<SpeciesProperty>>> createSummary(
	    List<SpeciesProperty> speciesProperties) {

	Map<String, Map<String, List<SpeciesProperty>>> result = new HashMap<>();

	Map<String, List<SpeciesProperty>> speciesMap = speciesProperties.stream()
		.collect(Collectors.groupingBy(p -> p.species));

	for (String speciesName : speciesMap.keySet()) {

	    List<SpeciesProperty> speciesList = speciesMap.get(speciesName);
	    Map<String, List<SpeciesProperty>> propertyMapForSpecies = speciesList.stream()
		    .collect(Collectors.groupingBy(p -> p.property));
	    result.put(speciesName, propertyMapForSpecies);

	    // for(String property : propertyMapForSpecies.keySet()){
	    // Map<String,List<SpeciesProperty>> propertyMap = new HashMap<>();
	    // propertyMap.put(property,propertyMapForSpecies.get(property));
	    //
	    // //Check if species already exists in the result map
	    // Map<String, List<SpeciesProperty>> map = result.get(speciesName);
	    // if(map==null){
	    // //species does not exist, so we put the new map in there
	    // result.put(speciesName,propertyMap);
	    // }else{
	    // //species exists, so we put only a key-value to the existing map
	    // map.put(property,propertyMapForSpecies.get(property));
	    // }
	    // }
	}

	return result;

    }

    public static List<SpeciesProperty> readSpeciesProperties(File file) throws Exception {

	Integer reference = getReference(file.getName());

	try (FileInputStream fin = new FileInputStream(file); Workbook wb = new XSSFWorkbook(fin)) {

	    final Sheet sheet = wb.getSheetAt(0);
	    final List<SpeciesProperty> propertyListInFile = new ArrayList<>();
	    List<String> columnNames = getColumnsInTopRow(sheet);
	    columnNames = columnNames.stream().map(s -> s.toLowerCase()).collect(Collectors.toList());
	    columnNames = columnNames.stream().filter(s -> !s.contains("reference")).collect(Collectors.toList());

	    final int physicalRows = sheet.getPhysicalNumberOfRows();

	    for (int i = 1; i < physicalRows; i++) {
		final Row row = sheet.getRow(i);
		if (row == null) {
		    continue;
		}
		if (isRowEmpty(row))
		    continue;

		List<SpeciesProperty> propertiesRow = readSpeciesPropertyRow(row, columnNames, reference);
		propertyListInFile.addAll(propertiesRow);
	    }

	    return propertyListInFile;

	} catch (IOException e) {
	    e.printStackTrace();
	    throw new Exception("Datei " + file.getName() + " konnte nicht gelesen werden.");
	}
    }

    private static Integer getReference(String fileName) throws Exception {

	try {
	    String strippedRef = fileName.substring(0, fileName.indexOf(" ")).replace("Ref", "");
	    return Integer.parseInt(strippedRef);
	} catch (Exception e) {
	    throw new Exception("Reference konnte nicht aus Dateiname gelesen werden.");
	}
    }

    public static List<String> getColumnsInTopRow(final Sheet sheet) {

	final List<String> values = getValues(CellType.STRING, 0, sheet);
	values.removeIf(Objects::isNull);
	return values;
    }

    public static List<SpeciesProperty> readSpeciesPropertyRow(final Row row, final List<String> propertyNames,
	    Integer reference) {

	List<SpeciesProperty> propertyList = new ArrayList<>();

	final String speciesName = row.getCell(COLUMN_INDEX_SPECIES_NAME).getStringCellValue().trim();

	for (int i = 1; i < propertyNames.size(); i++) {
	    String stringValue = null;
	    Double doubleValue = null;
	    final String propertyName = propertyNames.get(i);
	    final Cell cell = row.getCell(i);
	    if (cell == null)
		continue;
	    switch (cell.getCellTypeEnum()) {
	    case NUMERIC: {
		doubleValue = cell.getNumericCellValue();
		if (doubleValue < 0) {
		    System.out.println("Wert negativ in Ref. " + reference + " [" + propertyNames.get(i) + "] ="
			    + doubleValue + ", Zeile " + row.getRowNum());
		}
		break;
	    }
	    case STRING: {
		String intermediateValue = cell.getStringCellValue();
		if (!hasValue(intermediateValue))
		    break;
		OptionalDouble maybe = tryDoubleConversion(intermediateValue);
		if (maybe.isPresent())
		    doubleValue = maybe.getAsDouble();
		else
		    stringValue = intermediateValue;
		break;
	    }
	    default:
		continue;
	    }
	    if (doubleValue != null || stringValue != null) {
		propertyList.add(new SpeciesProperty(speciesName, propertyName, doubleValue, stringValue, reference));
	    }
	}
	return propertyList;
    }

    private static boolean hasValue(String intermediateValue) {
	return !(intermediateValue == null || intermediateValue.trim().isEmpty() || intermediateValue.equals("NA"));
    }

    private static OptionalDouble tryDoubleConversion(String stringValue) {

	OptionalDouble maybe = OptionalDouble.empty();
	stringValue = applyCleaning(stringValue);

	try {

	    Double value = Double.valueOf(stringValue);
	    maybe = OptionalDouble.of(value);
	} catch (NumberFormatException e) {
	    // ignore
	}
	return maybe;
    }

    private static String applyCleaning(String stringValue) {

	Function<String, String> removePlusMinus = (s -> {
	    int indexOf = s.indexOf("+/-");
	    if (indexOf != -1) {
		return s.substring(0, indexOf);
	    }
	    return s;
	});

	Function<String, String> removeThousandSeparators = (s -> {
	    if (s.split(",").length >= 3) {
		return s.replace(",", "");
	    }
	    return s;
	});

	return removePlusMinus.andThen(removeThousandSeparators).apply(stringValue);

    }

    /**
     * Checks if the row is empty. If the first column is empty it's assumed to
     * be empty.
     * 
     * @param row
     *            the row to check
     * @return
     */
    private static boolean isRowEmpty(final Row row) {
	final Cell cell = row.getCell(0);
	if (cell == null)
	    return true;
	return cell.getStringCellValue().trim().isEmpty();
    }

    /**
     * Read all String values in the specified row. If a cell does not contain a
     * String, a NULL is added instead.
     * 
     * @param row
     *            the index of the specified row
     * @return
     * 
     */
    public static List<String> getValues(final CellType cellType, final int rowIndex, final Sheet sheet) {

	if (rowIndex > sheet.getPhysicalNumberOfRows() - 1)
	    throw new IllegalArgumentException("no physical row for requested index");

	final Row row = sheet.getRow(rowIndex);
	if (row == null)
	    return Collections.emptyList();

	final List<String> values = new ArrayList<>();

	final short minColIndex = row.getFirstCellNum();
	final short maxColIndex = row.getLastCellNum();

	for (short colIx = minColIndex; colIx < maxColIndex; colIx++) {
	    final Cell cell = row.getCell(colIx);
	    if (cell != null) {
		try {
		    final CellType currentCellType = cell.getCellTypeEnum();

		    String value = null;
		    if (currentCellType.equals(cellType)) {
			value = getCellValue(cellType, cell);

		    }
		    values.add(value);
		} catch (final Exception e) {
		    values.add(null);
		}
	    }
	}
	return values;
    }

    public static String getCellValue(final CellType type, final Cell cell) {
	String value;
	switch (type) {

	case NUMERIC:
	    value = new Double(cell.getNumericCellValue()).toString();
	    break;

	case STRING:
	    value = cell.getStringCellValue().trim();
	    break;

	case BLANK:
	    value = "";
	    break;

	case FORMULA:
	    value = cell.getCellFormula();

	default:
	    value = null;

	}
	return value;
    }

    public static double round(double value, int places) {
	if (places < 0)
	    throw new IllegalArgumentException();

	BigDecimal bd = new BigDecimal(value);
	bd = bd.setScale(places, RoundingMode.HALF_UP);
	return bd.doubleValue();
    }

}
