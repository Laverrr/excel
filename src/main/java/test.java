import cn.hutool.core.util.StrUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;


public class test {

    public static void main(String[] args) throws IOException {
        //文件地址
        String path = "C:\\Users\\L\\Desktop\\test3.xlsx";
        //txt文件路径（如果想读取网页,设置 txtPath = null）
        String txtPath = "C:\\Users\\L\\Desktop\\test.txt";
        //目标url
        String url = "https://zh.aiavitality.com.hk/vmp-hk/";
        String source;
        if (StrUtil.isNotBlank(txtPath)) {
            //读取txt文字
            source = TxtUtil.readTxt(txtPath);
        } else {
            //读取网站文字
            Document document = Jsoup.connect(url).get();
            Elements allElements = document.getAllElements();
            source = allElements.toString();
        }
        //读取第几个sheet
        int numberOfSheet = 0;
        //读取第几列（默认第二页，cell = 1 为第二列）
        int numberOfCell = 1;
        //1.读取Excel文档对象
        Workbook wb;
        File file = new File(path);
        if (file.getName().endsWith("xls")) {
            wb = new HSSFWorkbook(new FileInputStream(path));
        } else {
            wb = new XSSFWorkbook(new FileInputStream(path));
        }
        // 获得工作表
        Sheet sheet = wb.getSheetAt(numberOfSheet);
        //获取总行数
        int rows = sheet.getPhysicalNumberOfRows();
        for (int i = 0; i < rows; i++) {
            // 获取第i行数据
            Row row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            // 获取第0格数据(第二列就填1)
            String content = row.getCell(numberOfCell).toString();
            if (StrUtil.isBlank(content)) {
                continue;
            }
            //在下一列写结果
            Cell cell = row.createCell(numberOfCell + 1);
            if (source.contains(content)) {
                cell.setCellValue("Y");
            } else {
                cell.setCellValue("N");
            }
        }
        FileOutputStream fileOutputStream = null;
        try {
            fileOutputStream = new FileOutputStream(path);
            wb.write(fileOutputStream);
        } finally {
            if (fileOutputStream != null) {
                fileOutputStream.close();
            }
        }
        System.out.println("处理结束");
    }
}
