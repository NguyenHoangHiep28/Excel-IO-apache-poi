import java.util.List;
import java.util.Map;

public class ExcelWriterStructural<T>{
    private String sheetName;
    private String[] headers;
    private Map<Integer, Integer> cellsWidth;
    private List<T> listItem;
    private Map<Integer, CellMapping> objectMappingRules;

    public ExcelWriterStructural() {
    }
    public ExcelWriterStructural(String sheetName, String[] headers, Map<Integer, Integer> cellsWidth, List<T> listItem, Map<Integer, CellMapping> objectMappingRules) {
        this.sheetName = sheetName;
        this.headers = headers;
        this.cellsWidth = cellsWidth;
        this.listItem = listItem;
        this.objectMappingRules = objectMappingRules;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public String[] getHeaders() {
        return headers;
    }

    public void setHeaders(String[] headers) {
        this.headers = headers;
    }

    public Map<Integer, Integer> getCellsWidth() {
        return cellsWidth;
    }

    public void setCellsWidth(Map<Integer, Integer> cellsWidth) {
        this.cellsWidth = cellsWidth;
    }

    public List<T> getListItem() {
        return listItem;
    }

    public void setListItem(List<T> listItem) {
        this.listItem = listItem;
    }

    public Map<Integer, CellMapping> getObjectMappingRules() {
        return objectMappingRules;
    }

    public void setObjectMappingRules(Map<Integer, CellMapping> objectMappingRules) {
        this.objectMappingRules = objectMappingRules;
    }
}
