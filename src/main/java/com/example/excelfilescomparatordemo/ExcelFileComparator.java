package com.example.excelfilescomparatordemo;

import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.scene.text.TextAlignment;
import javafx.stage.Stage;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import java.util.ArrayList;
import java.util.List;

import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFileComparator extends Application {

    private static final String title = "File Comparator";
    private static final String firstFileLabelText = "Укажите путь к старому файлу:";
    private static final String secondFileLabelText = "Укажите путь к новому файлу:";
    private static final String outputFileLabelText = "Укажите путь к файлу сравнения";
    private static final String compareButtonText = "Сравнить";
    private static final String successfulOperationMessage = "Сравнение завершено. Результаты сохранены в файле: ";
    private static final String oldFile = "Старый файл: ";
    private static final String newFile = "Новый файл: ";
    private static final String comparisonFile = "Файл сравнения: ";
    private static final String outputWorkBookSheet1Name = "Удалено";
    private static final String outputWorkBookSheet2Name = "Обновлено";
    private static final String outputWorkBookSheet3Name = "Добавлено";
    private static final String simpleDateFormatPattern = "dd.MM.yyyy";

    public static void main(String[] args) {
        launch(args);
    }

    @Override
    public void start(Stage primaryStage) throws Exception {

        primaryStage.setTitle(title);

        // Создание элементов интерфейса
        Label firstFileLabel = new Label(firstFileLabelText);
        firstFileLabel.setTextAlignment(TextAlignment.RIGHT);
        TextField firstFileTextField = new TextField();

        Label secondFileLabel = new Label(secondFileLabelText);
        secondFileLabel.setTextAlignment(TextAlignment.RIGHT);
        TextField secondFileTextField = new TextField();

        Label outputFileLabel = new Label(outputFileLabelText);
        outputFileLabel.setTextAlignment(TextAlignment.RIGHT);
        TextField outputTextField = new TextField();

        Button compareButton = new Button(compareButtonText);

        // Создание компоновки HBox для каждой пары Label и TextField
        HBox firstFileBox = new HBox(10);
        firstFileBox.setAlignment(Pos.CENTER_LEFT);
        firstFileBox.getChildren().addAll(firstFileLabel, firstFileTextField);
        HBox.setHgrow(firstFileTextField, Priority.ALWAYS);

        HBox secondFileBox = new HBox(10);
        secondFileBox.setAlignment(Pos.CENTER_LEFT);
        secondFileBox.getChildren().addAll(secondFileLabel, secondFileTextField);
        HBox.setHgrow(secondFileTextField, Priority.ALWAYS);

        HBox outputFileBox = new HBox(10);
        outputFileBox.setAlignment(Pos.CENTER_LEFT);
        outputFileBox.getChildren().addAll(outputFileLabel, outputTextField);
        HBox.setHgrow(outputTextField, Priority.ALWAYS);

        VBox vbox = new VBox(10);
        vbox.setPadding(new Insets(10));
        vbox.getChildren().addAll(firstFileBox, secondFileBox, outputFileBox,
                compareButton);

        // Обработка нажатия кнопки
        compareButton.setOnAction(e -> {
            String firstFilePath = firstFileTextField.getText();
            String secondFilePath = secondFileTextField.getText();
            String outputFilePath = outputTextField.getText();

            // Здесь можно добавить логику сравнения файлов
            // используя указанные пути oldFilePath и newFilePath
            // и выполнить необходимые действия

            try {
                FileInputStream firstFile = new FileInputStream(firstFilePath);
                FileInputStream secondFile = new FileInputStream(secondFilePath);
                FileOutputStream outputFile = new FileOutputStream(outputFilePath);
                Workbook firstWorkBook = new XSSFWorkbook(firstFile);
                Workbook secondWorkbook = new XSSFWorkbook(secondFile);
                Workbook outputWorkbook = new XSSFWorkbook();

                compareAndGenerateReport(firstWorkBook, secondWorkbook, outputWorkbook);

                outputWorkbook.write(outputFile);
                System.out.println(successfulOperationMessage + outputFilePath);
            } catch (IOException exception) {
                throw new RuntimeException(exception);
            }

            // Пример вывода результата в консоль
            System.out.println(oldFile + firstFilePath);
            System.out.println(newFile + secondFilePath);
            System.out.println(comparisonFile + outputFilePath);
        });

        Scene scene = new Scene(vbox);
        primaryStage.setScene(scene);

        // Установка максимального размера окна, чтобы кнопки управления окном были видны
        primaryStage.setMaximized(true);

        primaryStage.show();
    }

    private static void compareAndGenerateReport(Workbook firstWorkBook, Workbook secondWorkbook, Workbook outputWorkbook) {
        List<CellData> deletedRecords = new ArrayList<>();
        List<CellData> updatedRecords = new ArrayList<>();
        List<CellData> addedRecords = new ArrayList<>();

        Sheet firstSheetOfTheFirstWorkbook = firstWorkBook.getSheetAt(0);
        Sheet firstSheetOfTheSecondWorkbook = secondWorkbook.getSheetAt(0);

        // Проверка удалённых записей и обновлённых записей
        int firstRows = firstSheetOfTheFirstWorkbook.getLastRowNum() + 1;
        int secondRows = firstSheetOfTheSecondWorkbook.getLastRowNum() + 1;
        int maxRows = Math.max(firstRows, secondRows);

        for (int rowIndex = 0; rowIndex < maxRows; rowIndex++) {
            Row firstRow = firstSheetOfTheFirstWorkbook.getRow(rowIndex);
            Row secondRow = firstSheetOfTheSecondWorkbook.getRow(rowIndex);

            if (secondRow == null) {
                // Ряд полностью удалён
                if (firstRow != null) {
                    for (Cell cell: firstRow) {
                        CellData cellData = getCellData(cell);
                        deletedRecords.add(cellData);
                    }
                }
            } else if (firstRow == null) {
                // Ряд полностью новый
                for (Cell cell: secondRow) {
                    CellData cellData = getCellData(cell);
                    addedRecords.add(cellData);
                }
            } else {
                // Проверка ячеек в ряду
                boolean rowUpdated = false;

                int firstCells = firstRow.getLastCellNum() + 1;
                int secondCells = secondRow.getLastCellNum() + 1;
                int maxCells = Math.max(firstCells, secondCells);

                for (int columnIndex = 0; columnIndex < maxCells; columnIndex++) {
                    Cell firstCell = firstRow.getCell(columnIndex);
                    Cell secondCell = secondRow.getCell(columnIndex);

                    if (secondCell == null) {
                        // Ячейка удалена
                        if (firstCell != null) {
                            CellData cellData = getCellData(firstCell);
                            deletedRecords.add(cellData);
                        }
                    } else if (firstCell == null) {
                        // Ячейка добавлена
                        CellData cellData = getCellData(secondCell);
                        addedRecords.add(cellData);
                        rowUpdated = true;
                    } else if (!isCellEqual(firstCell, secondCell)) {
                        // Ячейка изменена
                        CellData cellData = getCellData(secondCell);
                        updatedRecords.add(cellData);
                        rowUpdated = true;
                    }
                }

                if (!rowUpdated && secondCells > firstCells) {
                    // Ряд во второй книге имеет дополнительные ячейки, считаем его обновлённым
                    for (int columnIndex = firstCells; columnIndex < secondCells; columnIndex++) {
                        Cell cell = secondRow.getCell(columnIndex);
                        CellData cellData = getCellData(cell);
                        updatedRecords.add(cellData);
                    }
                }
            }
        }

        createReportSheet(outputWorkbook, outputWorkBookSheet1Name, deletedRecords);
        createReportSheet(outputWorkbook, outputWorkBookSheet2Name, updatedRecords);
        createReportSheet(outputWorkbook, outputWorkBookSheet3Name, addedRecords);
    }

    private static class CellData {
        private int rowIndex;
        private int columnIndex;
        private String value;

        public CellData(int rowIndex, int columnIndex, String value) {
            this.rowIndex = rowIndex;
            this.columnIndex = columnIndex;
            this.value = value;
        }

        public int getRowIndex() {
            return rowIndex;
        }

        public int getColumnIndex() {
            return columnIndex;
        }

        public String getValue() {
            return value;
        }
    }

    // В "Обновлено" будет показываться значение ячейки из второго (нового) файла (первоначальное).
    private static CellData getCellData(Cell cell) {
        int rowIndex = cell.getRowIndex();
        int columnIndex = cell.getColumnIndex();
        String value;

        // Обработка различных типов ячеек
        switch (cell.getCellType()) {
            case STRING:
                value = cell.getStringCellValue();
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    Date dateValue = cell.getDateCellValue();
                    SimpleDateFormat dateFormat = new SimpleDateFormat(simpleDateFormatPattern);
                    value = dateFormat.format(dateValue);
                } else {
                    double numericValue = cell.getNumericCellValue();
                    value = String.valueOf(numericValue);
                }
                break;
            case BOOLEAN:
                boolean booleanValue = cell.getBooleanCellValue();
                value = String.valueOf(booleanValue);
                break;
            default:
                value = "";
                break;
        }

        return new CellData(rowIndex, columnIndex, value);
    }

    private static boolean isCellEqual(Cell firstCell, Cell secondCell) {
        if (firstCell.getCellType() != secondCell.getCellType()) {
            return false;
        }

        // Сравнение значений разных типов ячеек
        switch (firstCell.getCellType()) {
            case STRING:
                return firstCell.getStringCellValue().equals(secondCell.getStringCellValue());
            case NUMERIC:
                return firstCell.getNumericCellValue() == secondCell.getNumericCellValue();
            case BOOLEAN:
                return firstCell.getBooleanCellValue() == secondCell.getBooleanCellValue();
            default:
                return true;
        }
    }

    private static void createReportSheet(Workbook workbook, String sheetName, List<CellData> cellDataList) {
        Sheet sheet = workbook.createSheet(sheetName);

        for (CellData cellData: cellDataList) {
            Row row = sheet.getRow(cellData.getRowIndex());
            if (row == null) {
                row = sheet.createRow(cellData.getRowIndex());
            }

            Cell cell = row.createCell(cellData.getColumnIndex());
            cell.setCellValue(cellData.getValue());
        }
    }
}
