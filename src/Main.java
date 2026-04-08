public class Main {
    public static void main(String[] args) {
        try {
            ExcelReducer reducer = new ExcelReducer();
            reducer.processFile("data/input/sonoma.xlsx");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}