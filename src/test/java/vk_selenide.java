import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.awt.*;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.SocketTimeoutException;

import static com.codeborne.selenide.Selenide.open;

/**
 * Created by SretenskyVD on 18.02.2020.
 */
public class vk_selenide {
    public static void main(String[] args) throws IOException, AWTException {
        System.setProperty("javax.net.ssl.trustStore", "S:/ProjectJava/vk/vk.crt.cer.jks");
        String Tovar = "РязаньСейчас";
        String Manual_category =Tovar;
        Robot robot = new Robot();
         Robot robot2 = new Robot();
//        open("https://vk.com/search?c%5Bgroup%5D=23213239&c%5Bsection%5D=people");
        String Path = "https://vk.com/search?c%5Bgroup%5D=23213239&c%5Bsection%5D=people";


        //cd C:\Program Files\Java\jdk1.7.0_79\bin
//keytool -import -v -file S:/ProjectJava/vk/vk.crt.cer -keystore S:/ProjectJava/vk/vk.crt.cer.jks -storepass drowssap

        String CatalogName = Tovar;
        Workbook wb = new HSSFWorkbook();
        CreationHelper createHelper = wb.getCreationHelper();
        Sheet sheet1 = wb.createSheet(CatalogName);
        FileOutputStream fileOut = new FileOutputStream("book_" + CatalogName + ".xls");


        try {
            wb.write(fileOut);
            fileOut.close();
        } catch (IOException e) {
            e.printStackTrace();


        }
        Sheet sheet = wb.getSheetAt(0);

        int countLinks =33;
         while (countLinks>=0) {
             open("https://vk.com/search?c%5Bgroup%5D=23213239&c%5Bsection%5D=people");
//             robot.keyPress(KeyEvent.VK_PAGE_DOWN);
//                robot.keyRelease(KeyEvent.VK_PAGE_DOWN);

            robot2.delay(1000);
             robot2.mouseMove(963, 250);
             robot2.mouseWheel(150);


             Document doc1 = Jsoup.connect(Path).get();

             Elements links3 = doc1.getElementsByClass("labeled name");
             int yyy = 0;
             for (Element link3 : links3) {
                countLinks=countLinks+33-links3.size();
                 System.out.println();
                 String addressUrl3 = (links3.get(yyy).select("a[href]").attr("abs:href"));
                 System.out.println(addressUrl3);

                 try {
                     Document doc4 = Jsoup.connect(addressUrl3).get();

                     String NamePrduct = doc4.getElementsByTag("h1").text();
                     System.out.println(NamePrduct);


                     int rowCount = sheet.getLastRowNum();
                     Row row = sheet.createRow(++rowCount);


                     Cell cell227 = row.createCell(1);
                     cell227.setCellValue(NamePrduct);


                     Cell cell1 = row.createCell(2);
                     cell1.setCellValue(Tovar);



                     try {
                         Elements Label = doc4.getElementsByClass("labeled");

                         for (Element Labels : Label) {
                             int LabelLength = Label.size();
                             System.out.println(LabelLength);

                             for (int i = 0; i <= LabelLength; i++) {
                                 Element Label_insta1 = doc4.getElementsByClass("labeled").get(i);
                                 String Label_insta = Label_insta1.toString();

                                 String sub = "http://instagram.com/";

                                 if (Label_insta.indexOf(sub) != -1) {

                                     String outInsta = doc4.getElementsByClass("labeled").get(i).text();
                                     System.out.println(outInsta);
                                     Cell cell5555opc1 = row.createCell(6);
                                     cell5555opc1.setCellValue(outInsta);

                                 }
                             }
                         }
                     } catch (java.lang.NullPointerException e) {
                         e.printStackTrace();
                     }

///////////////////////////////////////////////////////////


                 } catch (IllegalArgumentException e) {
                     e.printStackTrace();
                 } catch (SocketTimeoutException e) {
                     e.printStackTrace();
                 } catch (IndexOutOfBoundsException e) {
                     e.printStackTrace();
                 } catch (NullPointerException e) {
                     e.printStackTrace();
                 }
//            open("https://vk.com/search?c%5Bgroup%5D=23213239&c%5Bsection%5D=people");
//            Robot robot2 = new Robot();
//            robot2.delay(1000);
//            robot2.mouseMove(963, 250);
//            robot2.mouseWheel(150);


                 System.out.println();
                 yyy++;

                 robot2.delay(1000);
                 robot2.mouseMove(963, 250);
                 robot2.mouseWheel(150);
             } //конец do
                 try {
                     FileOutputStream fileOut1 = new FileOutputStream("book_" + CatalogName + ".xls");
                     wb.write(fileOut1);
                     fileOut1.close();

                 } catch (FileNotFoundException e) {
                     e.printStackTrace();

                 } catch (IOException e) {
                     e.printStackTrace();
                 }
//////////////


             }


    }

}
