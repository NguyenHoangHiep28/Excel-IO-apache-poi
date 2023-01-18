import org.apache.commons.collections4.IteratorUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class POIExcelReader<T> {
    private static Logger logger = LoggerFactory.getLogger(POIExcelReader.class);
    private Class<T> clazz;
    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private Map<Integer, String> objectMappingRules;
    public POIExcelReader(File file, Map<Integer, String> objectMappingRules, Class<T> clazz) throws IOException {
        FileInputStream fileInputStream = new FileInputStream(file);
        this.workbook = new XSSFWorkbook(fileInputStream);
        this.clazz = clazz;
        this.objectMappingRules = objectMappingRules;
    }

    public List<T> getResult(int startRow, int endRow) throws Exception {
        if (startRow < 0) {
            throw new Exception("Start row must be greater than Zero");
        }
        List<T> result = new ArrayList<>();
        this.sheet = workbook.getSheetAt(0);
        //Iterate through each row one by one
        List<Row> rowList = IteratorUtils.toList(sheet.iterator());
        int length = rowList.size();
        if (length > 0) {
            for (int i = 0; i < length; i++) {
                if (i < startRow) {
                    continue;
                } else if (endRow > startRow && i > endRow) {
                    break;
                } else {
                    List<Cell> cellList = IteratorUtils.toList(rowList.get(i).cellIterator());
                    int cellLength = cellList.size();
                    T o = null;
                    try {
                        o = this.clazz.newInstance();
                    } catch (InstantiationException e) {
                        logger.error(e.getMessage());
                        logger.error("Required a no arguments constructor of generic type!");
                        break;
                    } catch (IllegalAccessException e) {
                        logger.error(e.getMessage());
                        logger.error("Constructors must be public!");
                        break;
                    }
                    for (int j = 0; j < cellLength; j++) {
                        Cell cell = cellList.get(j);
                        String property = this.objectMappingRules.get(j);
                        if (property != null && !property.isEmpty()) {
                            try {
                                Field field = this.clazz.getDeclaredField(property);
                                String name = "set" + field.getName().substring(0, 1).toUpperCase() + field.getName().substring(1);
                                Method setter = this.clazz.getDeclaredMethod(name, field.getType());
                                invokeSetter(cell, field, setter, o);
                            } catch (NoSuchMethodException | NoSuchFieldException e) {
                                logger.error(e.getMessage());
                                logger.error("Field or method does not exist!" );
                            }
                        }
                    }
                    result.add(o);
                }
            }
        }
        return result;
    }
    private void invokeSetter(Cell cell, Field field, Method setter, T object) {
        try {
            switch (cell.getCellType())
            {
                case NUMERIC:
                    double value = cell.getNumericCellValue();
                    String typeName = field.getType().getTypeName();
                    switch (typeName) {
                        case "java.lang.Integer":
                            Integer integerValue = (int) value;
                            setter.invoke(object, integerValue);
                            break;
                        case "java.lang.Double":
                            setter.invoke(object, value);
                            break;
                        case "java.lang.Long":
                            Long longValue = (long) value;
                            setter.invoke(object, longValue);
                            break;
                    }
                    break;
                case STRING:
                    setter.invoke(object, cell.getStringCellValue());
                    break;
                default:
                    break;
            }
        } catch (IllegalAccessException e) {
            logger.error(e.getMessage());
            logger.error("Setters of generic type must be public!");
        } catch (InvocationTargetException e) {
            throw new RuntimeException(e);
        }
    }
}
