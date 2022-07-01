import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;

public class ObjectToExcel {
    public void makeExcel1(List<Person> persons, Map<String, String> mp) throws IllegalAccessException {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("test sheet");

        int headerIdx = 0;
        Row row = sheet.createRow(0);
        for (Field field : Person.class.getDeclaredFields()) {
            field.setAccessible(true);
            if (mp.get(field.getName()) == null || mp.get(field.getName()).equals("N")) continue;
            row.createCell(headerIdx++).setCellValue(mp.get(field.getName()));
        }

        for (int i = 1; i <= persons.size(); i++) {
            row = sheet.createRow(i);
            int idx = 0;
            for (Field field : Person.class.getDeclaredFields()) {
                field.setAccessible(true);
                if (mp.get(field.getName()) == null || mp.get(field.getName()).equals("N")) continue;

                if (field.get(persons.get(i-1)) instanceof String) {
                    row.createCell(idx++).setCellValue((String) field.get(persons.get(i-1)));
                }
                else if (field.get(persons.get(i-1)) instanceof Integer) {
                    Integer a = (Integer) field.get(persons.get(i-1));
                    row.createCell(idx++).setCellValue(a.toString());
                }
                else {
                    row.createCell(idx++).setCellValue("unknown type");
                }
            }
        }
        return;
    }
}
