import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

public class POIExcelWriter {
    private String templatePath;
    private String fileName;
    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private FileOutputStream fileOutputStream;
    private List<ExcelTemplateWriterStructural<?>> templateWriterStructurals;
    private List<AdvancedExcelStructural<?>> advancedExcelStructurals;
    // Styling static variables

    public static final int FONT_ALIGNMENT_CENTER = 0;
    public static final int FONT_ALIGNMENT_LEFT = 1;
    public static final int FONT_ALIGNMENT_RIGHT = 2;
    public static final int TEXT_BOLD = 3;
    public static final int BORDERED = 4;

    public POIExcelWriter(List<ExcelTemplateWriterStructural<?>> structurals, String templatePath) {
        this.templatePath = templatePath;
        this.templateWriterStructurals = structurals;
    }
    public POIExcelWriter(String fileName, List<AdvancedExcelStructural<?>> structurals) {
        this.fileName = fileName;
        this.advancedExcelStructurals = structurals;
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

    private void writeFromTemplate() throws IOException {
        File f = new File(templatePath);
        FileInputStream fileInputStream = new FileInputStream(f);
        File newFile = new File(f.getName());
        FileUtils.copyFile(f, newFile);
        workbook = new XSSFWorkbook(fileInputStream);
        this.fileOutputStream = new FileOutputStream(newFile);
        this.templateWriterStructurals.forEach(t -> {
            this.sheet = workbook.getSheetAt(t.getSheetAt());
            AtomicInteger rowCount = new AtomicInteger(t.getStartRow());
            CellStyle style = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setFontHeightInPoints((short) 14);
            font.setBold(true);
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
    }

    private void write() throws IOException {
        workbook = new XSSFWorkbook();
        this.advancedExcelStructurals.forEach(t -> {
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
    }

    private CellStyle getCellStyle(int[] cellStyles) {
        CellStyle cellStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        for (Integer style :
                cellStyles) {
            switch (style) {
                case FONT_ALIGNMENT_CENTER:
                    cellStyle.setAlignment(HorizontalAlignment.CENTER);
                    break;
                case FONT_ALIGNMENT_LEFT:
                    cellStyle.setAlignment(HorizontalAlignment.LEFT);
                    break;
                case FONT_ALIGNMENT_RIGHT:
                    cellStyle.setAlignment(HorizontalAlignment.RIGHT);
                    break;
                case TEXT_BOLD:
                    font.setBold(true);
                    break;
                case BORDERED:
                    cellStyle.setBorderBottom(BorderStyle.THIN);
                    cellStyle.setBorderTop(BorderStyle.THIN);
                    cellStyle.setBorderRight(BorderStyle.THIN);
                    cellStyle.setBorderLeft(BorderStyle.THIN);
                    break;
            }
        }
        cellStyle.setFont(font);
        return cellStyle;
    }

    private void fillValueToCell(Row row, int columnCount, Object valueOfCell, CellStyle cellStyle) {
        Cell cell = row.createCell(columnCount);
        try {
            if (valueOfCell instanceof Integer) {
                cell.setCellValue((Integer) valueOfCell);
            } else if (valueOfCell instanceof Long) {
                cell.setCellValue((Long) valueOfCell);
            } else if (valueOfCell instanceof String) {
                cell.setCellValue((String) valueOfCell);
            } else if (valueOfCell instanceof Double) {
                DataFormat format = workbook.createDataFormat();
                cellStyle.setDataFormat(format.getFormat("0.00"));
                cell.setCellValue((Double) valueOfCell);
            } else if (valueOfCell instanceof Date) {
                DataFormat format = workbook.createDataFormat();
                cellStyle.setDataFormat(format.getFormat("dd/MM/yyyy"));
                cell.setCellValue((Date) valueOfCell);
            } else {
                cell.setCellValue((Boolean) valueOfCell);
            }
            cell.setCellStyle(cellStyle);
        } catch (NullPointerException e) {
        }
    }
    public void generateExcelFile() throws IOException {
        if (this.templateWriterStructurals != null && !this.templateWriterStructurals.isEmpty()) {
            writeFromTemplate();
        } else {
            File newFile = new File(fileName + ".xlsx");
            if (!newFile.exists()) {
                newFile.createNewFile();
            }
            this.fileOutputStream = new FileOutputStream(newFile);
            write();
        }
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
