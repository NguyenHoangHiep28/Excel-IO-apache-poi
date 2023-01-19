import org.apache.commons.io.FileUtils;

import java.io.File;
import java.io.FileOutputStream;
import java.util.*;

public class App {
    public static void main(String[] args) throws Exception {
//        File f = new File("./src/Book1.xlsx");
//        Map<Integer, String> objectMappingRules = new HashMap<>();
//        objectMappingRules.put(0, "id");
//        objectMappingRules.put(1, "name");
//        objectMappingRules.put(2, "grade");
//        POIExcelReader<Student> accountPOIExcelReader = new POIExcelReader<>(f, objectMappingRules, Student.class);
//        List<Student> a = accountPOIExcelReader.getResult(1, 2);
//        if (!a.isEmpty()) {
//            a.forEach(i -> {
//                System.out.println(i.toString());
//            });
//        }
//        String templatePath = "./src/template/Book1.xlsx";
        List<Account> accounts = new ArrayList<>();
        accounts.add(new Account(1, "Hoang Van Test", 1000.0, new Date()));
        accounts.add(new Account(2, "Nguyen Van Bug", 500.0, new Date()));
        Map<Integer, CellMapping> accountObjectMappingRules = new HashMap<>();
        accountObjectMappingRules.put(0, new CellMapping("id", POIExcelWriter.FONT_ALIGNMENT_CENTER, POIExcelWriter.BORDERED));
        accountObjectMappingRules.put(1, new CellMapping("name", POIExcelWriter.FONT_ALIGNMENT_LEFT, POIExcelWriter.BORDERED));
        accountObjectMappingRules.put(2, new CellMapping("balance", POIExcelWriter.FONT_ALIGNMENT_RIGHT, POIExcelWriter.BORDERED));
        accountObjectMappingRules.put(3, new CellMapping("effectiveStartDate", POIExcelWriter.FONT_ALIGNMENT_CENTER, POIExcelWriter.BORDERED));
//        List<Student> students = new ArrayList<>();
//        students.add(new Student(4, "Nguyen Van Test", "G"));
//        students.add(new Student(5, "Nguyen Van Bug", "F"));
//        Map<Integer, String> objectMappingRules = new HashMap<>();
//        objectMappingRules.put(0, "id");
//        objectMappingRules.put(1, "name");
//        objectMappingRules.put(2, "grade");
//        ExcelTemplateWriterStructural<Student> structural = new ExcelTemplateWriterStructural<>(students, 0, 4, objectMappingRules);
//        ExcelTemplateWriterStructural<Account> structural1 = new ExcelTemplateWriterStructural<>(accounts, 1, 2, accountObjectMappingRules);
//        List<ExcelTemplateWriterStructural<?>> structurals = new ArrayList<>();
//        structurals.add(structural);
//        structurals.add(structural1);
//        POIExcelWriter studentPOIExcelWriter = new POIExcelWriter(structurals, templatePath);
//        studentPOIExcelWriter.generateExcelFile();
        String[] headers = {"Id", "Name", "Balace", "Effective Start Date"};
        Map<Integer, Integer> cellWidths = new HashMap<>();
        cellWidths.put(0, 20);
        cellWidths.put(1, 50);
        cellWidths.put(2, 20);
        cellWidths.put(3, 50);
        AdvancedExcelStructural<Account> advancedExcelStructural = new AdvancedExcelStructural<Account>("Account", headers, cellWidths, accounts, accountObjectMappingRules);
        List<AdvancedExcelStructural<?>> advancedExcelStructuralList = new ArrayList<>();
        advancedExcelStructuralList.add(advancedExcelStructural);
        POIExcelWriter accountWriter = new POIExcelWriter("account", advancedExcelStructuralList);
        accountWriter.generateExcelFile();

    }
}
