import java.io.File;
import java.util.*;

public class App {
    public static void main(String[] args) throws Exception {
        File f = new File("./src/template/Book1.xlsx");
        Map<Integer, String> objectMappingRules = new HashMap<>();
        objectMappingRules.put(0, "id");
        objectMappingRules.put(1, "name");
        objectMappingRules.put(2, "grade");
        POIExcelReader<Student> accountPOIExcelReader = new POIExcelReader<>(f, objectMappingRules, Student.class);
        List<Student> a = accountPOIExcelReader.getResult(1, -1);
        if (!a.isEmpty()) {
            a.forEach(i -> {
                System.out.println(i.toString());
            });
        }
//        String templatePath = "./src/template/Book1.xlsx";
//        List<Account> accounts = new ArrayList<>();
//        accounts.add(new Account(1, "Hoang Van Test1", 1200.0, new Date()));
//        accounts.add(new Account(2, "Nguyen Van Bug2", 510.0, new Date()));
//        Map<Integer, CellMapping> accountObjectMappingRules = new HashMap<>();
//        accountObjectMappingRules.put(0, new CellMapping("id", POINewExcelWriter.FONT_ALIGNMENT_CENTER, POINewExcelWriter.BORDERED));
//        accountObjectMappingRules.put(1, new CellMapping("name", POINewExcelWriter.FONT_ALIGNMENT_LEFT, POINewExcelWriter.BORDERED));
//        accountObjectMappingRules.put(2, new CellMapping("balance", POINewExcelWriter.FONT_ALIGNMENT_RIGHT, POINewExcelWriter.BORDERED));
//        accountObjectMappingRules.put(3, new CellMapping("effectiveStartDate", POINewExcelWriter.FONT_ALIGNMENT_CENTER, POINewExcelWriter.BORDERED));
//        List<Student> students = new ArrayList<>();
//        students.add(new Student(4, "Nguyen Van Test", "G"));
//        students.add(new Student(5, "Nguyen Van Bug", "F"));
//        Map<Integer, CellMapping> objectMappingRules = new HashMap<>();
//        objectMappingRules.put(0, new CellMapping("id", POINewExcelWriter.FONT_ALIGNMENT_CENTER, POINewExcelWriter.BORDERED, POIExcelWriter.TEXT_BOLD));
//        objectMappingRules.put(1, new CellMapping("name", POINewExcelWriter.FONT_ALIGNMENT_CENTER, POINewExcelWriter.BORDERED, POIExcelWriter.TEXT_BOLD));
//        objectMappingRules.put(2, new CellMapping("grade", POINewExcelWriter.FONT_ALIGNMENT_CENTER, POINewExcelWriter.BORDERED, POIExcelWriter.TEXT_BOLD));
//        ExcelTemplateWriterStructural<Student> structural = new ExcelTemplateWriterStructural<>(students, 0, 4, objectMappingRules);
//        ExcelTemplateWriterStructural<Account> structural1 = new ExcelTemplateWriterStructural<>(accounts, 1, 2, accountObjectMappingRules);
//        List<ExcelTemplateWriterStructural<?>> structurals = new ArrayList<>();
//        structurals.add(structural);
//        structurals.add(structural1);
//        POIExcelWriter studentPOIExcelWriter = new POITemplateExcelWriter(templatePath, "baoCao", structurals);
//        studentPOIExcelWriter.generateExcelFile();
//        String[] headers = {"Id", "Name", "Balace", "Effective Start Date"};
//        Map<Integer, Integer> cellWidths = new HashMap<>();
//        cellWidths.put(0, 20);
//        cellWidths.put(1, 50);
//        cellWidths.put(2, 20);
//        cellWidths.put(3, 50);
//        AdvancedExcelStructural<Account> advancedExcelStructural = new AdvancedExcelStructural<Account>("Account", headers, cellWidths, accounts, accountObjectMappingRules);
//        List<AdvancedExcelStructural<?>> advancedExcelStructuralList = new ArrayList<>();
//        advancedExcelStructuralList.add(advancedExcelStructural);
//        POIExcelWriter accountWriter = new POINewExcelWriter("account", advancedExcelStructuralList);
//        accountWriter.generateExcelFile();

    }
}
