import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
public abstract class POIExcelWriter {
    protected XSSFWorkbook workbook;
    protected String fileName;
    protected XSSFSheet sheet;
    protected FileOutputStream fileOutputStream;
    // Styling static variables
    public static final int FONT_ALIGNMENT_CENTER = 0;
    public static final int FONT_ALIGNMENT_LEFT = 1;
    public static final int FONT_ALIGNMENT_RIGHT = 2;
    public static final int TEXT_BOLD = 3;
    public static final int BORDERED = 4;
    protected CellStyle getCellStyle(int[] cellStyles) {
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
    protected void fillValueToCell(Row row, int columnCount, Object valueOfCell, CellStyle cellStyle) {
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
    public abstract void generateExcelFile() throws IOException;

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }
}
