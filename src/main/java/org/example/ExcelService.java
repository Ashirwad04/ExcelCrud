package org.example;

import java.io.IOException;

public interface ExcelService {
    void createData(String filePath, Object[][] data) throws IOException;
    void readData(String filePath) throws IOException;
    void readDataById(String filePath, int id) throws IOException;
    void updateDataById(String filePath, int id, Object[] newData) throws IOException;
    void deleteDataById(String filePath, int id) throws IOException;
}
