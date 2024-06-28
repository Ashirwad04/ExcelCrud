package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

public class ExcelServiceImpl implements ExcelService {

    @Override
    public void createData(String filePath, Object[][] data) throws IOException {

        Workbook workbook;
        Sheet sheet;

        //check if the file exits
        try (FileInputStream fileInputStream = new FileInputStream(filePath)) {
            workbook = new XSSFWorkbook(fileInputStream);
            sheet = workbook.getSheetAt(0);
            }
        catch (IOException e){
            workbook = new XSSFWorkbook();
            sheet=workbook.createSheet("sheet1");
            //Add header if creating new file
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("ID");
            header.createCell(1).setCellValue("Name");
            header.createCell(2).setCellValue("Std");
            header.createCell(3).setCellValue("RollNo");
            header.createCell(4).setCellValue("Age");
            header.createCell(5).setCellValue("Address");

        }
        int lastId= getLastId(sheet);
        int rowNum=sheet.getLastRowNum();

        for (int i = 0; i < data.length; i++) {
            Object[] rowData = data[i];
            Row row = sheet.createRow(++rowNum);
            row.createCell(0).setCellValue(++lastId); // Increment ID
            int colNum = 1; // Start from the second column
            for (int j = 0; j < rowData.length; j++) {
                Object field = rowData[j];
                Cell cell = row.createCell(colNum++);
                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
                }
            }
        }

        try(FileOutputStream outputStream = new FileOutputStream(filePath)) {
            workbook.write(outputStream);
        }
        workbook.close();


    }



    @Override
    public void readData(String filePath) throws IOException {
        try(FileInputStream file = new FileInputStream(filePath)){
            Workbook workbook = new XSSFWorkbook(file);
            Sheet sheet = workbook.getSheetAt(0);

            int numberOfRows = sheet.getPhysicalNumberOfRows();
            for (int i = 0; i < numberOfRows; i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    int numberOfCells = row.getPhysicalNumberOfCells();
                    for (int j = 0; j < numberOfCells; j++) {
                        Cell cell = row.getCell(j);
                        if (cell != null) {
                            switch (cell.getCellType()) {
                                case STRING:
                                    System.out.print(cell.getStringCellValue() + "\t");
                                    break;
                                case NUMERIC:
                                    System.out.print(cell.getNumericCellValue() + "\t");
                                    break;
                                case BOOLEAN:
                                    System.out.print(cell.getBooleanCellValue() + "\t");
                                    break;
                                default:
                                    System.out.print("Unknown Type" + "\t");
                            }
                        }
                    }
                    System.out.println();
                }
            }

        workbook.close();

        }

    }

    @Override
    public void readDataById(String filePath, int id) throws IOException {
        try (FileInputStream file = new FileInputStream(filePath)){
            Workbook workbook= new XSSFWorkbook(file);
            Sheet sheet = workbook.getSheetAt(0);

            boolean found = false;

            int numberOfRows = sheet.getLastRowNum();
            for (int i = 0; i <= numberOfRows; i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell idCell = row.getCell(0);
                    if (idCell != null && idCell.getCellType() == CellType.NUMERIC && idCell.getNumericCellValue() == id) {
                        int numberOfCells = row.getLastCellNum();
                        for (int j = 0; j < numberOfCells; j++) {
                            Cell cell = row.getCell(j);
                            if (cell != null) {
                                switch (cell.getCellType()) {
                                    case STRING:
                                        System.out.print(cell.getStringCellValue() + "\t");
                                        break;
                                    case NUMERIC:
                                        System.out.print(cell.getNumericCellValue() + "\t");
                                        break;
                                    case BOOLEAN:
                                        System.out.print(cell.getBooleanCellValue() + "\t");
                                        break;
                                    default:
                                        System.out.print("Unknown Type" + "\t");
                                }
                            }
                        }
                        System.out.println();
                        found = true;
                        break;
                    }
                }
            }

            if (!found) {
                System.out.println("Record with ID " + id + " not found.");
            }
            workbook.close();
        }
    }

    @Override
    public void updateDataById(String filePath, int id, Object[] newData) throws IOException {
        try (FileInputStream file = new FileInputStream(filePath)){
            Workbook workbook = new XSSFWorkbook(file);
            Sheet sheet = workbook.getSheetAt(0);

            boolean found= false;

            for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue; // Skip empty rows
                Cell idCell = row.getCell(0);
                if (idCell != null && idCell.getCellType() == CellType.NUMERIC && idCell.getNumericCellValue() == id) {
                    int colNum = 1;
                    for (Object field : newData) {
                        Cell cell = row.getCell(colNum);
                        if (cell == null) {
                            cell = row.createCell(colNum);
                        }
                        if (field instanceof String) {
                            cell.setCellValue((String) field);
                        } else if (field instanceof Integer) {
                            cell.setCellValue((Integer) field);
                        }
                        colNum++;
                    }
                    found = true;
                    break;
                }
            }
            if (!found){
                System.out.println("Record with Id "+id+" not found.");

            }
            try (FileOutputStream outputStream=new FileOutputStream(filePath)){
                workbook.write(outputStream);
            }
            workbook.close();

        }

    }

    @Override
    public void deleteDataById(String filePath, int id) throws IOException {

        try(FileInputStream file=new FileInputStream(filePath)){
            Workbook workbook= new XSSFWorkbook(file);
            Sheet sheet=workbook.getSheetAt(0);
            boolean found = false;
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()){
                Row row=rowIterator.next();
                Cell idCell=row.getCell(0);
                if (idCell !=null && idCell.getCellType()==CellType.NUMERIC && idCell.getNumericCellValue()==id){
                    int rowIndex = row.getRowNum();
                    removeRow(sheet,rowIndex);
                    found = true;
                    break;

                }
            }
            if (!found){
                System.out.println("Record with ID "+id+" not found");


            }
            try (FileOutputStream outputStream= new FileOutputStream(filePath)){
                workbook.write(outputStream);

            }
            workbook.close();
        }
    }

    private void removeRow(Sheet sheet, int rowIndex) {
        int lastRowNum=sheet.getLastRowNum();
        if (rowIndex>=0 && rowIndex < lastRowNum){
            sheet.shiftRows(rowIndex+1,lastRowNum,-1);

        } else if (rowIndex == lastRowNum) {
            Row removingRow=sheet.getRow(rowIndex);
            if (removingRow != null){
                sheet.removeRow(removingRow);
            }

        }
    }


    private int getLastId(Sheet sheet) {

        int lastId = 0;
        int numberOfRows = sheet.getPhysicalNumberOfRows();
        for (int i = 0; i < numberOfRows; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                Cell idCell = row.getCell(0);
                if (idCell != null && idCell.getCellType() == CellType.NUMERIC) {
                    int id = (int) idCell.getNumericCellValue();
                    if (id > lastId) {
                        lastId = id;
                    }
                }
            }
        }
        return lastId;
    }
}
