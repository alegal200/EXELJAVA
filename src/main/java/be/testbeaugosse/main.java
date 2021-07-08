package be.testbeaugosse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class main{
    public static void main(String[]args) throws IOException {

        System.out.println("hello word");

        String fileLocation = "D:\\alex\\perso\\java\\intelijiexcelTest\\src\\main\\resources\\a.xlsx";
        FileInputStream file = new FileInputStream(new File(fileLocation));
        Workbook workbook = new XSSFWorkbook(file);

        Sheet sheet = workbook.getSheetAt(0);

        int numrow  ; // recupe le nbr de lignes

        numrow = sheet.getLastRowNum() ;

        for (int i = 0; i < numrow; i++) {
            for (int j = 0; j < sheet.getRow(i).getLastCellNum(); j++) { // recup la dernier case de la ligne
                if(sheet.getRow(i).getCell(j) != null )
                System.out.print( sheet.getRow(i).getCell(j)+"\t" );
            }
            System.out.println("");         // affiche l ensemble du contenu

        }

        // écriture

            Row row1 = sheet.createRow(0);

        for (int i = 0; i < 15; i++) {
            Cell c1 =   row1.createCell(i) ;
            c1.setCellValue("val"+i);

        }

        FileOutputStream fos = new FileOutputStream("D:\\alex\\perso\\java\\intelijiexcelTest\\src\\main\\resources\\a.xlsx") ;
        workbook.write(fos);
        fos.flush();

    }
}

