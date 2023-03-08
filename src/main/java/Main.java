import io.restassured.RestAssured;
import io.restassured.http.ContentType;
import io.restassured.response.Response;
import io.restassured.specification.RequestSpecification;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;
import org.json.JSONArray;

import java.io.*;
import java.net.URISyntaxException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Main {
    static String  jarLocation;
    static String[] orderedHeaders = {"cui", "denumire", "data", "act", "stare_inregistrare", "scpTVA", "data_inceput_ScpTVA","data_sfarsit_ScpTVA","data_anul_imp_ScpTVA","mesaj_ScpTVA","dataInceputTvaInc","dataSfarsitTvaInc","dataActualizareTvaInc","dataPublicareTvaInc","tipActTvaInc","statusTvaIncasare","dataInactivare","dataReactivare","dataPublicare","dataRadiere","statusInactivi","dataInceputSplitTVA","dataAnulareSplitTVA","statusSplitTVA"};

    static {
        try {
            jarLocation = Main.class.getProtectionDomain().getCodeSource().getLocation().toURI().getPath() + "\\..\\";
        } catch (URISyntaxException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) throws IOException, InterruptedException {
        RestAssured.baseURI = "https://webservicesp.anaf.ro/PlatitorTvaRest/api/v7/ws/tva";

        RequestSpecification request = RestAssured.given();
        request.contentType(ContentType.JSON);

        JSONArray jarray = new JSONArray();

        LocalDate currentDate = LocalDate.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");
        String formattedDate = currentDate.format(formatter);

        ArrayList<Long> cuiuri = ExcelRead();


        for(Long cui : cuiuri){
            JSONObject json = new JSONObject();
            json.put("data", formattedDate);
            json.put("cui", cui);
            jarray.put(json);
        }

        // split jarray in multiple arrays of maximum 500 elements
        var splitArrays = splitJSONArray(jarray, 400);


        JSONArray result = new JSONArray();
        String jsonResponse = null;

        for(JSONArray array : splitArrays){
            request.body(array.toString());
            Response response = request.post();

            jsonResponse = response.getBody().asString();

            // get all elements from "found" array
            JSONArray found = new JSONObject(jsonResponse).getJSONArray("found");
            for (int i = 0; i < found.length(); i++) {
                JSONObject obj = found.getJSONObject(i);
                JSONObject subObj = new JSONObject();
                // System.out.println(obj.toString(4));
                try{
                    obj.remove("adresa_sediu_social");
                }catch (Exception ignored){}

                try{
                    obj.remove("adresa_domiciliu_fiscal");
                }catch (Exception ignored){}

                JSONObject dateGenerale = obj.getJSONObject("date_generale");
                // System.out.println(dateGenerale.get("cui"));
                subObj.put("cui", dateGenerale.get("cui"));
                subObj.put("denumire", dateGenerale.get("denumire"));
                subObj.put("data", dateGenerale.get("data"));
                subObj.put("act", dateGenerale.get("act"));
                subObj.put("stare_inregistrare", dateGenerale.get("stare_inregistrare"));

                try{
                    obj.remove("date_generale");
                }catch (Exception ignored){}

                for(String key : obj.keySet()){
                    JSONObject thisObj = obj.getJSONObject(key);
                    for(String key2 : thisObj.keySet()){
                        subObj.put(key2, thisObj.get(key2));
                    }
                }

                result.put(subObj);
            }
            // only sleep if there are more arrays to process
            if(splitArrays.indexOf(array) != splitArrays.size() - 1)
                Thread.sleep(2000);
        }

        ExcelWrite(result);
    }


    public static ArrayList<Long> ExcelRead() throws IOException {
        String filePath = jarLocation + "input.xlsx";
        ArrayList<Long> numbers = new ArrayList<>();

        FileInputStream inputStream = new FileInputStream(filePath);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);

        int lastRowNum = sheet.getLastRowNum();
        for (int i = 0; i <= lastRowNum; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                Cell cell = row.getCell(0);
                if (cell != null) {
                    double number = cell.getNumericCellValue();
                    numbers.add((long) number);
                }
            }
        }

        return numbers;
    }


    public static void ExcelWrite(JSONArray jsonArray) {    //TODO handle order of columns in excel file (now it's random)
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Data");

            writeHeaderRow(sheet, jsonArray.getJSONObject(0));
            writeDataRows(sheet, jsonArray);

            FileOutputStream outputStream = new FileOutputStream(jarLocation + "output.xlsx");
            workbook.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void writeHeaderRow(Sheet sheet, JSONObject json) {
        Row headerRow = sheet.createRow(0);

        int cellNum = 0;

        for (String key : orderedHeaders){
            Cell cell = headerRow.createCell(cellNum++);
            cell.setCellValue(key);
        }
    }

    private static void writeDataRows(Sheet sheet, JSONArray jsonArray) {
        int rowNum = 1;
        for (int i = 0; i < jsonArray.length(); i++) {
            JSONObject jsonObject = jsonArray.getJSONObject(i);
            Row row = sheet.createRow(rowNum++);

            int cellNum = 0;

            for (String key : orderedHeaders) {
                Cell cell = row.createCell(cellNum++);
                cell.setCellValue(jsonObject.get(key).toString());
            }
        }
    }

    public static List<JSONArray> splitJSONArray(JSONArray jsonArray, int maxSubarraySize) {
        List<JSONArray> subarrays = new ArrayList<>();
        int length = jsonArray.length();
        int startIndex = 0;

        while (startIndex < length) {
            int endIndex = Math.min(startIndex + maxSubarraySize, length);
            JSONArray subarray = new JSONArray();

            for (int i = startIndex; i < endIndex; i++) {
                subarray.put(jsonArray.get(i));
            }

            subarrays.add(subarray);
            startIndex = endIndex;
        }

        return subarrays;
    }
}
