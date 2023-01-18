import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

public class App {
    public static void main(String[] args) throws Exception {
        File f = new File("./src/Book1.xlsx");
        Map<Integer, String> objectMappingRules = new HashMap<>();
        objectMappingRules.put(0, "id");
        objectMappingRules.put(1, "name");
        objectMappingRules.put(2, "grade");
        POIExcelReader<Student> accountPOIExcelReader = new POIExcelReader<>(f, objectMappingRules, Student.class);
        List<Student> a = accountPOIExcelReader.getResult(1, 2);
        if (!a.isEmpty()) {
            a.forEach(i -> {
                System.out.println(i.toString());
            });
        }
    }
}
