import java.io.File;
import java.io.FileWriter;
import java.io.PrintWriter;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReducer {

    public void processFile(String filePath) throws Exception {
        File file = new File(filePath);
        XSSFWorkbook workbook = new XSSFWorkbook(file);

        // Detailed rows: one row per entity/source
        ArrayList<String[]> detailedRows = new ArrayList<String[]>();

        // Grouped rows: one row per person, with all entities combined
        LinkedHashMap<String, LinkedHashSet<String>> groupedMap = new LinkedHashMap<String, LinkedHashSet<String>>();

        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            XSSFSheet sheet = workbook.getSheetAt(i);
            String sheetName = sheet.getSheetName();

            if (!sheetName.equals("Schedule A1")
                    && !sheetName.equals("Schedule A-2")
                    && !sheetName.equals("Schedule C - Income Section")) {
                continue;
            }

            int headerRow = getHeaderRow(sheetName);
            int dataStartRow = headerRow + 2;

            Map<String, Integer> columns = buildColumnMap(sheet, headerRow);

            System.out.println("=================================");
            System.out.println("Sheet: " + sheetName);
            System.out.println("=================================");

            for (int rowIndex = dataStartRow; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                String lastName = getCellSafe(sheet, rowIndex, columns.get("last"));
                String firstName = getCellSafe(sheet, rowIndex, columns.get("first"));
                String agency = getCellSafe(sheet, rowIndex, columns.get("agency"));
                String filedDate = getCellSafe(sheet, rowIndex, columns.get("filed"));

                String entityOrSource = "";

                if (sheetName.equals("Schedule A1")) {
                    entityOrSource = getCellSafe(sheet, rowIndex, columns.get("entity"));
                } else if (sheetName.equals("Schedule A-2")) {
                    entityOrSource = getCellSafe(sheet, rowIndex, columns.get("entityTrust"));
                } else if (sheetName.equals("Schedule C - Income Section")) {
                    entityOrSource = getCellSafe(sheet, rowIndex, columns.get("source"));
                }

                if (!lastName.equals("") && !entityOrSource.equals("")) {
                    // Print to terminal
                    System.out.println(
                        lastName + " | " +
                        firstName + " | " +
                        agency + " | " +
                        entityOrSource + " | " +
                        filedDate + " | " +
                        sheetName
                    );

                    // Save detailed row
                    detailedRows.add(new String[] {
                        lastName,
                        firstName,
                        agency,
                        entityOrSource,
                        filedDate,
                        sheetName
                    });

                    // Save grouped row
                    String groupedKey =
                        lastName + "|" +
                        firstName + "|" +
                        agency + "|" +
                        filedDate + "|" +
                        sheetName;

                    if (!groupedMap.containsKey(groupedKey)) {
                        groupedMap.put(groupedKey, new LinkedHashSet<String>());
                    }

                    groupedMap.get(groupedKey).add(entityOrSource);
                }
            }

            System.out.println();
        }

        workbook.close();

        writeDetailedCsv("data/output/sonoma_detailed.csv", detailedRows);
        writeGroupedCsv("data/output/sonoma_grouped.csv", groupedMap);

        System.out.println("Detailed CSV written to data/output/sonoma_detailed.csv");
        System.out.println("Grouped CSV written to data/output/sonoma_grouped.csv");
    }

    private void writeDetailedCsv(String outputPath, ArrayList<String[]> rows) throws Exception {
        PrintWriter writer = new PrintWriter(new FileWriter(outputPath));

        writer.println("last_name,first_name,agency,entity_name,filed_date,schedule");

        for (int i = 0; i < rows.size(); i++) {
            String[] row = rows.get(i);

            writer.println(
                csv(row[0]) + "," +
                csv(row[1]) + "," +
                csv(row[2]) + "," +
                csv(row[3]) + "," +
                csv(row[4]) + "," +
                csv(row[5])
            );
        }

        writer.close();
    }

    private void writeGroupedCsv(String outputPath, LinkedHashMap<String, LinkedHashSet<String>> groupedMap) throws Exception {
        PrintWriter writer = new PrintWriter(new FileWriter(outputPath));

        writer.println("last_name,first_name,agency,filed_date,schedule,entities");

        for (Map.Entry<String, LinkedHashSet<String>> entry : groupedMap.entrySet()) {
            String key = entry.getKey();
            LinkedHashSet<String> entities = entry.getValue();

            String[] parts = key.split("\\|", -1);

            String lastName = parts[0];
            String firstName = parts[1];
            String agency = parts[2];
            String filedDate = parts[3];
            String schedule = parts[4];

            StringBuilder entityList = new StringBuilder();
            int count = 0;

            for (String entity : entities) {
                if (count > 0) {
                    entityList.append("; ");
                }
                entityList.append(entity);
                count++;
            }

            writer.println(
                csv(lastName) + "," +
                csv(firstName) + "," +
                csv(agency) + "," +
                csv(filedDate) + "," +
                csv(schedule) + "," +
                csv(entityList.toString())
            );
        }

        writer.close();
    }

    private String csv(String value) {
        if (value == null) {
            return "\"\"";
        }

        String escaped = value.replace("\"", "\"\"");
        return "\"" + escaped + "\"";
    }

    private int getHeaderRow(String sheetName) {
        if (sheetName.equals("Schedule A1")) {
            return 0;
        }
        if (sheetName.equals("Schedule A-2")) {
            return 1;
        }
        if (sheetName.equals("Schedule C - Income Section")) {
            return 1;
        }
        return 0;
    }

    private Map<String, Integer> buildColumnMap(XSSFSheet sheet, int headerRow) {
        LinkedHashMap<String, Integer> map = new LinkedHashMap<String, Integer>();

        if (sheet.getRow(headerRow) == null) {
            return map;
        }

        short lastCell = sheet.getRow(headerRow).getLastCellNum();

        for (int col = 0; col < lastCell; col++) {
            String header = getCell(sheet, headerRow, col).toLowerCase();

            if (header.contains("last name")) {
                map.put("last", col);
            }
            if (header.contains("first name")) {
                map.put("first", col);
            }
            if (header.equals("agency") || header.contains("agency")) {
                map.put("agency", col);
            }
            if (header.contains("filed date")) {
                map.put("filed", col);
            }
            if (header.contains("name of business entity")
                    && !header.contains("trust")) {
                map.put("entity", col);
            }
            if (header.contains("name of business entity or trust")) {
                map.put("entityTrust", col);
            }
            if (header.contains("name of source")) {
                map.put("source", col);
            }
        }

        return map;
    }

    private String getCell(XSSFSheet sheet, int row, int col) {
        if (sheet.getRow(row) == null) return "";
        Cell cell = sheet.getRow(row).getCell(col);
        if (cell == null) return "";

        if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
            return new SimpleDateFormat("MM/dd/yyyy").format(cell.getDateCellValue());
        }

        return cell.toString().trim();
    }

    private String getCellSafe(XSSFSheet sheet, int row, Integer col) {
        if (col == null) return "";
        return getCell(sheet, row, col);
    }
}