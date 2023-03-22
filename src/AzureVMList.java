import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Map;

public class AzureVMList {

    public static void main(String[] args) throws IOException {
        // Prompt the user to enter the resource group name
        System.out.print("Enter the resource group name: ");
        Scanner scanner = new Scanner(System.in);
        String resourceGroup = scanner.nextLine();

        // Execute the Azure CLI command and capture the output as a string
        ProcessBuilder processBuilder = new ProcessBuilder("az", "vm", "list", "-g", resourceGroup, "-d", "--query", "[].{Name:name, Region:location, ComputerName:osProfile.computerName, UserName:osProfile.adminUsername, Password:osProfile.adminPassword}", "-o", "json");
        Process process = processBuilder.start();
        String output = new String(process.getInputStream().readAllBytes());
        try {
            process.waitFor();
        } catch (InterruptedException e) {
            e.printStackTrace();
        }

        // Convert the JSON output to a Java object
        Gson gson = new Gson();
        List<Map<String, Object>> data = gson.fromJson(output, new TypeToken<List<Map<String, Object>>>(){}.getType());

        // Create a new Excel workbook and worksheet
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("VM List");

        // Add headers to the worksheet
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Name");
        headerRow.createCell(1).setCellValue("Region");
        headerRow.createCell(2).setCellValue("ComputerName");
        headerRow.createCell(3).setCellValue("UserName");
        headerRow.createCell(4).setCellValue("Password");

        // Loop through the Java object and populate the worksheet
        int row = 1;
        for (Map<String, Object> item : data) {
            Row dataRow = sheet.createRow(row++);
            dataRow.createCell(0).setCellValue((String) item.get("Name"));
            dataRow.createCell(1).setCellValue((String) item.get("Region"));
            dataRow.createCell(2).setCellValue((String) item.get("ComputerName"));
            dataRow.createCell(3).setCellValue((String) item.get("UserName"));
            dataRow.createCell(4).setCellValue((String) item.get("Password"));
        }

        // Save the workbook
        File file = new File("C:/Users/jselvaraj/Desktop/file.xlsx");
        file.getParentFile().mkdirs();
        workbook.write(file);
        workbook.close();
    }
}
