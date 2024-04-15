import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ArunchalpradeshHospitalScraperToExcel {
    public static void main(String[] args) {
        String url = "http://example.com/ Arunchalpradesh-hospitals"; // Replace with actual URL
        String excelFilePath = " Arunchalpradesh_hospitals.xlsx"; // Path to Excel file
        try {
            Document doc = Jsoup.connect(url).get();
            Elements hospitals = doc.select("div.hospital"); // Adjust CSS selector as per the HTML structure

            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet(" Arunchalpradesh Hospitals");
            int rowNum = 0;

            for (Element hospital : hospitals) {
                String name = hospital.select("h2.name").text();
                String address = hospital.select("p.address").text();
                String phone = hospital.select("p.phone").text();
                String pincode = extractPincode(address); // Extract pin code from address

                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(name);
                row.createCell(1).setCellValue(address);
                row.createCell(2).setCellValue(pincode);
                row.createCell(3).setCellValue(phone);
            }

            FileOutputStream outputStream = new FileOutputStream(excelFilePath);
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();

            System.out.println("Data exported successfully to Excel file.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String extractPincode(String address) {
        String pincode = "";
        Pattern pattern = Pattern.compile("\\b\\d{6}\\b"); // Assuming pin code is 6 digits
        Matcher matcher = pattern.matcher(address);
        if (matcher.find()) {
            pincode = matcher.group();
        }
        return pincode;
    }
}