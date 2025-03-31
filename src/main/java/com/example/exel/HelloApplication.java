package com.example.exel;

import javafx.application.Application;
import javafx.scene.Scene;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.scene.control.Button;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class HelloApplication extends Application {
    @Override
    public void start(Stage stage) {
        VBox root = new VBox();
        Button button = new Button("Выбрать Excel-файл");
        root.getChildren().add(button);

        button.setOnAction(e -> {
            FileChooser fileChooser = new FileChooser();
            fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Files", "*.xlsx", "*.xls"));
            File file = fileChooser.showOpenDialog(stage);

            if (file != null) { // Проверяем, выбрал ли пользователь файл
                try {
                    List<Product> products = readExcelFile(file); // Читаем Excel
                    System.out.println("Загружено продуктов: " + products.size()); // Выводим количество загруженных товаров
                } catch (IOException | InvalidFormatException ex) {
                    ex.printStackTrace();
                }
            }
        });

        Scene scene = new Scene(root, 320, 240);
        stage.setTitle("Excel Loader");
        stage.setScene(scene);
        stage.show();
    }

    public static List<Product> readExcelFile(File file) throws IOException, InvalidFormatException {
        List<Product> products = new ArrayList<>(); // Создаем список для хранения продуктов
        FileInputStream fis = new FileInputStream(file);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);

        for (int i = 1; i <= sheet.getLastRowNum(); i++) { // Начинаем с 1, если первая строка — заголовки
            int id = (int) sheet.getRow(i).getCell(0).getNumericCellValue();
            String name = sheet.getRow(i).getCell(1).getStringCellValue();
            double price = sheet.getRow(i).getCell(2).getNumericCellValue();
            int quantity = (int) sheet.getRow(i).getCell(3).getNumericCellValue();
            double finalPrice = sheet.getRow(i).getCell(4).getNumericCellValue();

            Product product = new Product(id, name, price, quantity, finalPrice, null);
            products.add(product);
        }

        System.out.println(products.get(4).getName());
        System.out.println(products.get(4).getPrice());
        System.out.println(products.size());

        workbook.close();
        return products; // Возвращаем список продуктов
    }

    public static void main(String[] args) {
        launch();
    }
}
