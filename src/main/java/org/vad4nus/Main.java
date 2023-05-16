package org.vad4nus;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Scanner;
import java.util.concurrent.atomic.AtomicReference;

public class Main {
    public static void main(String[] args) {
        try {
            Scanner scanner = new Scanner(System.in);
            System.out.println("Введите путь к файлу .xlsx для нормализации");
            File file = new File(scanner.nextLine().replace("\"",""));
            System.out.println("Введите путь к файлу .xlsx для вывода");
            File outFile = new File(scanner.nextLine().replace("\"",""));
            XSSFWorkbook inpWorkbook = new XSSFWorkbook(file);
            XSSFWorkbook tempWorkbook = new XSSFWorkbook();
            XSSFWorkbook outWorkbook = new XSSFWorkbook();
            Sheet inpSheet = inpWorkbook.getSheetAt(0);
            Sheet calcSheet = tempWorkbook.createSheet("calculated");
            Sheet transpondedSheet = tempWorkbook.createSheet("transponded");
            Sheet outSheet = outWorkbook.createSheet("Лист1");

            for(int i = 0; i < inpSheet.getRow(0).getLastCellNum(); i++){
                transpondedSheet.createRow(i);
            }

            for (int r = 0; r < inpSheet.getLastRowNum(); r++) {
                for (int c = 0; c < inpSheet.getRow(r).getLastCellNum(); c++) {
                    Cell cell = transpondedSheet.getRow(c).createCell(r);
                    if (inpSheet.getRow(r).getCell(c).getCellType() == CellType.STRING) {
                        cell.setCellValue(inpSheet.getRow(r).getCell(c).getStringCellValue());
                    } else {
                        cell.setCellValue(inpSheet.getRow(r).getCell(c).getNumericCellValue());
                    }
                }
            }

            transpondedSheet.forEach(row -> {
                int rowNum = row.getRowNum();
                calcSheet.createRow(rowNum);
                AtomicReference<Double> max = new AtomicReference<>((double) 0),
                        min = new AtomicReference<>(Double.MAX_VALUE);
                row.forEach(cell -> {
                    if (cell.getCellType() == CellType.NUMERIC) {
                        if (cell.getNumericCellValue() < min.get()) {
                            min.set(cell.getNumericCellValue());
                        }
                        if (cell.getNumericCellValue() > max.get()) {
                            max.set(cell.getNumericCellValue());
                        }
                    }
                });
                row.forEach(cell -> {
                    if (cell.getColumnIndex() == 0) {
                        calcSheet.getRow(rowNum).createCell(cell.getColumnIndex())
                                .setCellValue(cell.getStringCellValue());
                    } else {
                        if (cell.getCellType() == CellType.NUMERIC) {
                            calcSheet.getRow(rowNum).createCell(cell.getColumnIndex()).
                                    setCellValue((cell.getNumericCellValue() - min.get()) / (max.get() - min.get()));
                        } else {
                            calcSheet.getRow(rowNum).createCell(cell.getColumnIndex()).setCellValue(cell.getStringCellValue());
                        }
                    }
                });
            });

            for(int i = 0; i < inpSheet.getLastRowNum(); i++){
                outSheet.createRow(i);
            }

            for (int r = 0; r < calcSheet.getLastRowNum()+1; r++) {
                for (int c = 0; c < calcSheet.getRow(r).getLastCellNum(); c++) {
                    outSheet.autoSizeColumn(c);
                    Cell cell = outSheet.getRow(c).createCell(r);
                    if (calcSheet.getRow(r).getCell(c).getCellType() == CellType.STRING) {
                        cell.setCellValue(calcSheet.getRow(r).getCell(c).getStringCellValue());
                    } else {
                        cell.setCellValue(calcSheet.getRow(r).getCell(c).getNumericCellValue());
                    }
                }
            }

            FileOutputStream out = new FileOutputStream(outFile);
            outWorkbook.write(out);
            out.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}