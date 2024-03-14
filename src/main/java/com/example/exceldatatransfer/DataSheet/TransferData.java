package com.example.exceldatatransfer.DataSheet;

import com.example.exceldatatransfer.Data.Person;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.ApplicationArguments;
import org.springframework.boot.ApplicationRunner;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

@Component
public class TransferData implements ApplicationRunner {
    @Override
    public void run(ApplicationArguments args) {

        try {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet spreadsheet = workbook.createSheet("Demo Data");
            XSSFRow row;
            Map<String, Person[]> demoData = new TreeMap<String, Person[]>();
            demoData.put("1", new Person[]{new Person("Id", "NAME", "age")});
            demoData.put("2", new Person[]{new Person("1", "Angel", "18")});
            demoData.put("3", new Person[]{new Person("2", "Georgi", "19")});
            demoData.put("4", new Person[]{new Person("3", "Ivan", "20")});
            demoData.put("5", new Person[]{new Person("4", "Martin", "21")});
            demoData.put("6", new Person[]{new Person("5", "Petar", "22")});
            Set<String> keyId = demoData.keySet();
            int rowId = 0;
            for (String key : keyId) {
                row = spreadsheet.createRow(rowId++);
                Person[] people = demoData.get(key);
                int cellId = 0;
                for (Person person : people) {
                    Cell id = row.createCell(cellId++);
                    id.setCellValue(person.getId());
                    Cell name = row.createCell(cellId++);
                    name.setCellValue(person.getName());
                    Cell age = row.createCell(cellId++);
                    age.setCellValue(person.getAge());
                }
            }
            createBorders(spreadsheet);
            FileOutputStream out = new FileOutputStream(new File("C:/excelTransferredData/DemoSheet.xlsx"));
            workbook.write(out);
            out.close();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    private static void createBorders(XSSFSheet spreadsheet) {
        CellRangeAddress region2 = new CellRangeAddress(0, 5, 0, 2);
        RegionUtil.setBorderTop(BorderStyle.THIN, region2, spreadsheet);
        RegionUtil.setBorderBottom(BorderStyle.THIN, region2, spreadsheet);
        RegionUtil.setBorderLeft(BorderStyle.THIN, region2, spreadsheet);
        RegionUtil.setBorderRight(BorderStyle.THIN, region2, spreadsheet);
        CellRangeAddress region = new CellRangeAddress(0, 0, 0, 2);
        RegionUtil.setBorderTop(BorderStyle.THICK, region, spreadsheet);
        RegionUtil.setBorderLeft(BorderStyle.THICK, region, spreadsheet);
        RegionUtil.setBorderRight(BorderStyle.THICK, region, spreadsheet);
        RegionUtil.setBorderBottom(BorderStyle.THICK, region, spreadsheet);
    }
}
