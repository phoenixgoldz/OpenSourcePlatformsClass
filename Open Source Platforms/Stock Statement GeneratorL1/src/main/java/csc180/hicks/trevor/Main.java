package csc180.hicks.trevor;

import com.aspose.cells.*;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;

import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.OutputStream;

public class Main {

    public static void main(String[] args) {
        System.out.println("Hello and welcome!");

        processStatements();
    }

    public static void processStatements() {
        try {
            String filename = "Data\\JsonFile\\all_stocks.json";
            Object json = new JSONParser().parse(new FileReader(filename));
            JSONArray jsonArray = (JSONArray) json;

            for (Object object : jsonArray) {
                JSONObject customerRecord = (JSONObject) object;

                String ssn = (String) customerRecord.get("ssn");
                String fullName = customerRecord.get("first_name") + " " + customerRecord.get("last_name");
                String email = (String) customerRecord.get("email");
                String phone = (String) customerRecord.get("phone");
                String accountNumber = String.valueOf(customerRecord.get("account_number"));
                double beginningBalance = Double.parseDouble(((String) customerRecord.get("beginning_balance")).replace("$", "").replace(",", ""));

                // Create an Excel workbook for each customer
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.getWorksheets().get(0);

                // Set column widths for a more professional layout
                worksheet.getCells().setColumnWidth(0, 20);
                worksheet.getCells().setColumnWidth(1, 30);
                worksheet.getCells().setColumnWidth(2, 15);
                worksheet.getCells().setColumnWidth(3, 30);
                worksheet.getCells().setColumnWidth(4, 15);

                // Insert customer details in the Excel sheet
                Cell cellA1 = worksheet.getCells().get("A1");
                cellA1.putValue("Statement Date:");
                cellA1.setStyle(createBoldStyle(workbook));

                Cell cellB1 = worksheet.getCells().get("B1");
                cellB1.putValue("Account Holder's Full Name:");
                cellB1.setStyle(createBoldStyle(workbook));

                Cell cellC1 = worksheet.getCells().get("C1");
                cellC1.putValue("SSN:");
                cellC1.setStyle(createBoldStyle(workbook));

                Cell cellD1 = worksheet.getCells().get("D1");
                cellD1.putValue("Email Address:");
                cellD1.setStyle(createBoldStyle(workbook));

                Cell cellE1 = worksheet.getCells().get("E1");
                cellE1.putValue("Phone:");
                cellE1.setStyle(createBoldStyle(workbook));

                Cell cellF1 = worksheet.getCells().get("F1");
                cellF1.putValue("Account Number:");
                cellF1.setStyle(createBoldStyle(workbook));

                worksheet.getCells().get("A2").putValue(java.time.LocalDate.now().toString());
                worksheet.getCells().get("B2").putValue(fullName);
                worksheet.getCells().get("C2").putValue(ssn);
                worksheet.getCells().get("D2").putValue(email);
                worksheet.getCells().get("E2").putValue(phone);
                worksheet.getCells().get("F2").putValue(accountNumber);

                // Add a header row for stock trades
                int headerRow = 4;
                worksheet.getCells().get("A" + headerRow).putValue("Type");
                worksheet.getCells().get("B" + headerRow).putValue("Stock Symbol");
                worksheet.getCells().get("C" + headerRow).putValue("Price per Share");
                worksheet.getCells().get("D" + headerRow).putValue("Number of Shares");
                worksheet.getCells().get("E" + headerRow).putValue("Total Amount");

                // Apply bold formatting to the header row
                Style style = worksheet.getCells().get("A" + headerRow).getStyle();
                Font font = style.getFont();
                font.setBold(true);
                worksheet.getCells().get("A" + headerRow).setStyle(style);

                style = worksheet.getCells().get("B" + headerRow).getStyle();
                font = style.getFont();
                font.setBold(true);
                worksheet.getCells().get("B" + headerRow).setStyle(style);

                style = worksheet.getCells().get("C" + headerRow).getStyle();
                font = style.getFont();
                font.setBold(true);
                worksheet.getCells().get("C" + headerRow).setStyle(style);

                style = worksheet.getCells().get("D" + headerRow).getStyle();
                font = style.getFont();
                font.setBold(true);
                worksheet.getCells().get("D" + headerRow).setStyle(style);

                style = worksheet.getCells().get("E" + headerRow).getStyle();
                font = style.getFont();
                font.setBold(true);
                worksheet.getCells().get("E" + headerRow).setStyle(style);

                // Process stock trades
                JSONArray stockTrades = (JSONArray) customerRecord.get("stock_trades");
                int row = 5;  // Start at row 5 for stock trades

                for (Object tradeObject : stockTrades) {
                    JSONObject trade = (JSONObject) tradeObject;

                    String tradeType = (String) trade.get("type");
                    String stockSymbol = (String) trade.get("stock_symbol");
                    int countShares = Integer.parseInt(String.valueOf(trade.get("count_shares")));
                    double pricePerShare = Double.parseDouble(((String) trade.get("price_per_share")).replace("$", "").replace(",", ""));
                    double totalAmount = pricePerShare * countShares;

                    // Insert stock trade details in the Excel sheet
                    worksheet.getCells().get("A" + row).putValue(tradeType);
                    worksheet.getCells().get("B" + row).putValue(stockSymbol);
                    worksheet.getCells().get("C" + row).putValue(pricePerShare);
                    worksheet.getCells().get("D" + row).putValue(countShares);
                    worksheet.getCells().get("E" + row).putValue(totalAmount);

                    row++;
                }

                // Calculate total cash and stock holdings
                double totalCash = beginningBalance;
                double totalStockValue = 0;

                for (Object tradeObject : stockTrades) {
                    JSONObject trade = (JSONObject) tradeObject;
                    String tradeType = (String) trade.get("type");
                    double pricePerShare = Double.parseDouble(((String) trade.get("price_per_share")).replace("$", "").replace(",", ""));
                    int countShares = Integer.parseInt(String.valueOf(trade.get("count_shares")));

                    if (tradeType.equals("Buy")) {
                        totalCash -= pricePerShare * countShares;
                        totalStockValue += pricePerShare * countShares;
                    } else if (tradeType.equals("Sell")) {
                        totalCash += pricePerShare * countShares;
                        totalStockValue -= pricePerShare * countShares;
                    }
                }

                // Add totals at the end of the table
                int totalCashRow = row + 2;
                int totalStockHoldingsRow = row + 3;

                // Set cell style to bold
                style = worksheet.getCells().get("D" + totalCashRow).getStyle();
                style.getFont().setBold(true);
                worksheet.getCells().get("D" + totalCashRow).setStyle(style);

                style = worksheet.getCells().get("D" + totalStockHoldingsRow).getStyle();
                style.getFont().setBold(true);
                worksheet.getCells().get("D" + totalStockHoldingsRow).setStyle(style);

                // Set the values (totals) without changing the font style
                worksheet.getCells().get("D" + totalCashRow).putValue("Total Cash:");
                worksheet.getCells().get("E" + totalCashRow).putValue(totalCash);
                worksheet.getCells().get("D" + totalStockHoldingsRow).putValue("Total Stock Holdings:");
                worksheet.getCells().get("E" + totalStockHoldingsRow).putValue(totalStockValue);

                // Save the Excel workbook as a PDF
                PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
                pdfSaveOptions.setOnePagePerSheet(true);

                OutputStream os = new FileOutputStream("output_" + accountNumber + ".pdf");
                workbook.save(os, pdfSaveOptions);
                os.close();
            }

        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    private static Style createBoldStyle(Workbook workbook) {
        Style style = workbook.createStyle();
        Font font = style.getFont();
        font.setBold(true);
        return style;
    }

}
