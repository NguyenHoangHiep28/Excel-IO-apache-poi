public class CellMapping {
    private String fieldName;
    private int[] cellStyles;

    public CellMapping() {
    }

    public CellMapping(String fieldName) {
        this.fieldName = fieldName;
    }

    public CellMapping(String fieldName, int... cellStyles) {
        this.fieldName = fieldName;
        this.cellStyles = cellStyles;
    }

    public String getFieldName() {
        return fieldName;
    }

    public void setFieldName(String fieldName) {
        this.fieldName = fieldName;
    }

    public int[] getCellStyles() {
        return cellStyles;
    }

    public void setCellStyles(int... cellStyles) {
        this.cellStyles = cellStyles;
    }
}
