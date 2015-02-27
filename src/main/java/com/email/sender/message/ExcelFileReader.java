package com.email.sender.message;


import com.sendgrid.SendGrid;
import com.sendgrid.SendGridException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;


/**
 * @author ajay
 */

public class ExcelFileReader {

    static int counter = 0;

    public static void main(String args[]) {
        try {

            FileInputStream file = new FileInputStream(new File("D:\\Testfile-Java-sendgrid.xlsx"));

            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            //Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);

            //Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                //For each row, iterate through all the columns
                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    //Check the cell type and format accordingly
                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "t");
                            break;
                        case Cell.CELL_TYPE_STRING:
                            counter++;
                            if (counter <= 400) {
                                emailSender(cell.getStringCellValue());
                            } else {
                                System.out.println("==========you have sent 400 emails ===");
                            }
                            break;
                    }

                }
                System.out.println("");
            }
            file.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    static void emailSender(String toEmailId) {

        /* The below line used to pass username and password for the send mail from send grid
        *  Firs argument is the username and second one is password.
        * */
        SendGrid sendgrid = new SendGrid("username", "password");
        SendGrid.Email email = new SendGrid.Email();


        /* From this we can set the email. From to subject and body. SO its very easy to pass email body of any body
        *  in from attribute and receiver will think that sender is abc...
        * */
        email.addTo(toEmailId.trim());
        email.setFrom("test@gmail.com");
        email.setSubject("Hello World");
        email.setText("My first email with SendGrid Java!");
        try {
            SendGrid.Response response = sendgrid.send(email);
            System.out.println(response.getMessage());
        } catch (SendGridException e) {
            System.err.println(e);
        }
    }
}
