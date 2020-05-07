package cn.nano.excelhandler;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.util.ArrayList;
import java.util.Scanner;

public class Main {
    public static void main(String[] args) {


        String sourPath = "/Users/p365/Desktop/location/location.xlsx";
        String destPath = "/Users/p365/Desktop/location/%s.txt";
        String valueFromat = "<string name=\"%1s\">%2s</string>";

        Scanner scanner = new Scanner(System.in);
        System.out.println("请输入目标文件路径，按回车确认。如：/Users/p365/Desktop/location/location.xlsx");
        sourPath = scanner.nextLine().trim();
        if (sourPath.length() == 0) {
            System.out.println("目标文件路径错误");
            return;
        }

        System.out.println("请输入输出文件夹路径，按回车确认。如：/Users/p365/Desktop/location");
        String path = scanner.nextLine().trim();
        if (path.length() == 0) {
            System.out.println("输出文件夹路径错误。");
            return;
        }

        destPath = path + "/%s.txt";

        System.out.println("执行中，请稍等");

        ArrayList<FileWriter> fileWritersList = new ArrayList<>();

        File file = new File(sourPath);
        try {
            FileInputStream in = new FileInputStream(file);

            XSSFWorkbook wb = new XSSFWorkbook(in);

            Sheet sheet = wb.getSheetAt(0); //取得“测试.xlsx”中的第一个表单
            int firstRowNum = sheet.getFirstRowNum();
            int lastRowNum = sheet.getLastRowNum();
            Row row = null;
            Cell cell = null;

            //先获取第一行数据，第一行表示有多少国家语言
            row = sheet.getRow(0);
            int firstCellNum = row.getFirstCellNum();
            int lastCellNum = row.getLastCellNum();
            for (int i = 1; i <= lastCellNum; i++) {
                cell = row.getCell(i);
                if (cell == null) {
                    continue;
                }
                FileWriter fw = new FileWriter(String.format(destPath,
                        cell.getStringCellValue().trim()));
                fileWritersList.add(fw);
            }

            //遍历其他行列数据
            String key = null;
            for (int i = firstRowNum + 1; i <= lastRowNum; i++) {
                row = sheet.getRow(i);//取得第i行 （从第二行开始取，因为第一行是表头）
                firstCellNum = row.getFirstCellNum();
                lastCellNum = row.getLastCellNum();
                key = row.getCell(0).getStringCellValue().trim();

                for (int j = 1; j <= lastCellNum; j++) {
                    cell = row.getCell(j);
                    if (cell == null) {
                        continue;
                    }

                    FileWriter writer = fileWritersList.get(j - 1);
                    writer.write(String.format(valueFromat, key, cell.getStringCellValue().trim()) + "\r\n");
                }
            }

            if (fileWritersList.size() > 0) {
                for (FileWriter writer : fileWritersList) {
                    writer.flush();
                    writer.close();
                }
            }

            System.out.println("执行完成");

        } catch (Exception e) {
            System.out.println("执行失败: " + e.getMessage());
        }
    }
}
