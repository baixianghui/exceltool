package base;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

public class Excel2Module {

    @Test
    public void testXlsx() throws Exception {
        File file = new File("C:\\Users\\GROW\\Desktop\\111.xlsx");
        System.out.println(file.exists());
        //获取输入流
        InputStream stream = new FileInputStream(file);
        Workbook xssfWorkbook = new XSSFWorkbook(stream);
        Sheet sheet = xssfWorkbook.getSheetAt(1);

        //获取最后一行的num，即总行数。此处从0开始计数
        int maxRow = sheet.getLastRowNum();
        System.out.println("总行数为：" + maxRow);
        for (int row = 1; row <= maxRow; row++) {
            String title = sheet.getRow(row).getCell(0).toString();
            String url = sheet.getRow(row).getCell(1).toString();

            System.out.println(title);
            System.out.println(url);
        }
    }
}
