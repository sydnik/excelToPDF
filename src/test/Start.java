package test;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;


public class Start {
    public static void main(String[] args) {
        long i1 = System.currentTimeMillis();
        try(FileInputStream stream = new FileInputStream("Прайс.xlsx");) {
            XSSFWorkbook workbook = new XSSFWorkbook(stream);
            XSSFSheet sheet = workbook.getSheetAt(workbook.getActiveSheetIndex());
            ReadExcel readExcel = new ReadExcel(sheet);
            Document document = new Document(new Rectangle(0,0,readExcel.getWidthPage(), readExcel.getHeightPage()),25f,25f,25f,25f);//Косяк не правильная установка ширины и  высоты я считываю всех строки а надо только несколько
            System.out.println("W"+readExcel.getWidthPage());
            System.out.println("H"+readExcel.getHeightPage());
            PdfWriter.getInstance(document,new FileOutputStream("Рузультат.pdf"));
            document.open();
            PDFCreater pdfCreater = new PDFCreater(document,readExcel);
            pdfCreater.createdPDF();
            document.close();
            System.out.println(System.currentTimeMillis()-i1);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (DocumentException e) {
            e.printStackTrace();
        }


    }
}
