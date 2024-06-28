package org.example;

import java.io.IOException;
import java.util.InputMismatchException;
import java.util.Scanner;

public class Main {
    public static void main(String[] args) {
        ExcelService excelService = new ExcelServiceImpl();
        String filePath = "ExcelSheet.xlsx";

        try (Scanner scanner = new Scanner(System.in)) {
            while (true) {
                printMenu();

                int choice = getChoice(scanner);
                if (choice == 6) {
                    System.out.println("Exiting...");
                    break;
                }

                try {
                    executeChoice(choice, excelService, filePath, scanner);
                } catch (IOException e) {
                    System.err.println("An error occurred while processing the Excel file: " + e.getMessage());
                } catch (InputMismatchException e) {
                    System.err.println("Invalid input. Please enter the correct data type.");
                    scanner.nextLine(); // Clear the invalid input
                }
            }
        }
    }

    private static void printMenu() {
        System.out.println("\nChoose an option:");
        System.out.println("1. Create Data");
        System.out.println("2. Read All Data");
        System.out.println("3. Read Data by ID");
        System.out.println("4. Update Data by ID");
        System.out.println("5. Delete Data by ID");
        System.out.println("6. Exit");
        System.out.print("Enter your choice: ");
    }

    private static int getChoice(Scanner scanner) {
        while (!scanner.hasNextInt()) {
            System.out.print("Invalid input. Enter a number between 1 and 6: ");
            scanner.next();
        }
        return scanner.nextInt();
    }

    private static void executeChoice(int choice, ExcelService excelService, String filePath, Scanner scanner) throws IOException {
        scanner.nextLine(); // Clear the buffer

        switch (choice) {
            case 1:
                createData(excelService, filePath, scanner);
                break;
            case 2:
                excelService.readData(filePath);
                break;
            case 3:
                readDataById(excelService, filePath, scanner);
                break;
            case 4:
                updateDataById(excelService, filePath, scanner);
                break;
            case 5:
                deleteDataById(excelService, filePath, scanner);
                break;
            default:
                System.out.println("Invalid choice. Please try again.");
        }
    }

    private static void createData(ExcelService excelService, String filePath, Scanner scanner) throws IOException {
        System.out.print("Enter the number of records to add: ");
        int numRecords = getValidIntInput(scanner);

        Object[][] data = new Object[numRecords][];
        for (int i = 0; i < numRecords; i++) {
            System.out.println("Enter the details for record " + (i + 1) + ":");
            String name = getStringInput(scanner, "Name: ");
            int std = getIntInput(scanner, "Std: ");
            int rollNo = getIntInput(scanner, "RollNo: ");
            int age = getIntInput(scanner, "Age: ");
            String address = getStringInput(scanner, "Address: ");
            data[i] = new Object[]{name, std, rollNo, age, address};
        }

        excelService.createData(filePath, data);
        System.out.println("Data added successfully.");
    }

    private static void readDataById(ExcelService excelService, String filePath, Scanner scanner) throws IOException {
        int id = getIntInput(scanner, "Enter the ID to read: ");
        excelService.readDataById(filePath, id);
    }

    private static void updateDataById(ExcelService excelService, String filePath, Scanner scanner) throws IOException {
        int id = getIntInput(scanner, "Enter the ID to update: ");
        System.out.println("Enter the new details:");
        String name = getStringInput(scanner, "Name: ");
        int std = getIntInput(scanner, "Std: ");
        int rollNo = getIntInput(scanner, "RollNo: ");
        int age = getIntInput(scanner, "Age: ");
        String address = getStringInput(scanner, "Address: ");

        Object[] newData = {name, std, rollNo, age, address};
        excelService.updateDataById(filePath, id, newData);
        System.out.println("Data updated successfully.");
    }

    private static void deleteDataById(ExcelService excelService, String filePath, Scanner scanner) throws IOException {
        int id = getIntInput(scanner, "Enter the ID to delete: ");
        excelService.deleteDataById(filePath, id);
        System.out.println("Data deleted successfully.");
    }

    private static String getStringInput(Scanner scanner, String prompt) {
        System.out.print(prompt);
        return scanner.nextLine();
    }

    private static int getIntInput(Scanner scanner, String prompt) {
        System.out.print(prompt);
        while (!scanner.hasNextInt()) {
            System.out.print("Invalid input. Please enter a number: ");
            scanner.next();
        }
        int input = scanner.nextInt();
        scanner.nextLine();
        return input;
    }

    private static int getValidIntInput(Scanner scanner) {
        while (!scanner.hasNextInt()) {
            System.out.print("Invalid input. Please enter a valid number: ");
            scanner.next();
        }
        int input = scanner.nextInt();
        scanner.nextLine();
        return input;
    }
}
