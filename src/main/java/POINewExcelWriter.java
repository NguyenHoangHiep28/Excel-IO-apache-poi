import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

public class POINewExcelWriter extends POIExcelWriter {
    private List<ExcelWriterStructural<?>> excelWriterStructurals;

    public POINewExcelWriter(String fileName, List<ExcelWriterStructural<?>> structurals) {
        this.fileName = fileName;
        this.excelWriterStructurals = structurals;
    }

    private void writeHeader(Map<Integer, Integer> sheetColumnsWidth, String... headers) {
        Row row = sheet.createRow(0);
        int[] headerStyle = {BORDERED, TEXT_BOLD};
        CellStyle style = getCellStyle(headerStyle);
        setSheetColumnsWidth(sheetColumnsWidth);
        Font font = workbook.createFont();
        font.setFontHeightInPoints((short) 16);
        style.setFont(font);
        int columnCount = 0;
        for (String header : headers) {
            fillValueToCell(row, columnCount++, header, style);
        }
    }

    public void generateExcelFile() throws IOException {
        File newFile = new File(fileName + ".xlsx");
        if (!newFile.exists()) {
            newFile.createNewFile();
        }
        this.fileOutputStream = new FileOutputStream(newFile);
        workbook = new XSSFWorkbook();
        this.excelWriterStructurals.forEach(t -> {
            this.sheet = workbook.createSheet(t.getSheetName());
            writeHeader(t.getCellsWidth(), t.getHeaders());
            AtomicInteger rowCount = new AtomicInteger(1);
            CellStyle style = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setFontHeightInPoints((short) 14);
            style.setFont(font);
            t.getListItem().forEach(record -> {
                Class<?> klass = record.getClass();
                Row row = sheet.createRow(rowCount.getAndIncrement());
                try {
                    for (Map.Entry<Integer, CellMapping> entry : t.getObjectMappingRules().entrySet()) {
                        if (entry.getValue() != null && entry.getValue().getFieldName() != null) {
                            Field field = klass.getDeclaredField(entry.getValue().getFieldName());
                            Method getter = klass.getDeclaredMethod("get" + field.getName().substring(0, 1).toUpperCase() + field.getName().substring(1));
                            CellStyle cellStyle = getCellStyle(entry.getValue().getCellStyles());
                            fillValueToCell(row, entry.getKey(), getter.invoke(record), cellStyle);
                        }
                    }
                } catch (Exception ex) {
                    ex.printStackTrace();
                }
            });
        });
        workbook.write(this.fileOutputStream);
        workbook.close();
        fileOutputStream.close();
    }

    private void setSheetColumnsWidth(Map<Integer, Integer> columnsWidth) {
        for (Map.Entry<Integer, Integer> entry : columnsWidth.entrySet()) {
            sheet.setColumnWidth(entry.getKey(), entry.getValue() * 256);
        }
    }
}
