import java.util.List;
import java.util.Map;

public class ExcelTemplateWriterStructural<T> {
    private List<T> listItem;
    private int sheetAt;
    private int startRow;
    private Map<Integer, CellMapping> objectMappingRules;

    public ExcelTemplateWriterStructural(List<T> listItem, int sheetAt, int startRow, Map<Integer, CellMapping> objectMappingRules) throws Exception {
        if (listItem == null || listItem.isEmpty() || sheetAt < 0 || objectMappingRules == null) {
            throw new Exception("Invalid instatiate payload");
        }
        this.listItem = listItem;
        this.sheetAt = sheetAt;
        this.startRow = startRow;
        this.objectMappingRules = objectMappingRules;
    }

    public List<T> getListItem() {
        return listItem;
    }

    public int getSheetAt() {
        return sheetAt;
    }

    public int getStartRow() {
        return startRow;
    }

    public Map<Integer, CellMapping> getObjectMappingRules() {
        return objectMappingRules;
    }
}
