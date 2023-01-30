import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

public class POITemplateExcelWriter extends POIExcelWriter {
    private String templatePath;
    private List<ExcelTemplateWriterStructural<?>> templateWriterStructurals;
    public POITemplateExcelWriter(String templatePath, String newFileName, List<ExcelTemplateWriterStructural<?>> structurals) {
        this.templatePath = templatePath;
        this.fileName = newFileName;
        this.templateWriterStructurals = structurals;
    }

    @Override
    public void generateExcelFile() throws IOException {
        File f = new File(templatePath);
        FileInputStream fileInputStream = new FileInputStream(f);
        File newFile = new File(this.fileName + ".xlsx");
        if (!newFile.exists()) {
            newFile.createNewFile();
        }
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
        workbook.write(this.fileOutputStream);
        workbook.close();
        fileOutputStream.close();
    }
}
