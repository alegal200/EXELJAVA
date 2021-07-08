package be.testbeaugosse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class main{
    public static void main(String[]args) throws IOException {

        System.out.println("hello word");

        String fileLocation = "D:\\alex\\perso\\java\\intelijiexcelTest\\src\\main\\resources\\a.xlsx";
        FileInputStream file = new FileInputStream(new File(fileLocation));
        Workbook workbook = new XSSFWorkbook(file);

        Sheet sheet = workbook.getSheetAt(0);

        int numrow  ;

        numrow = sheet.getLastRowNum() ;
        System.out.println("nbr de colones : "+ sheet.getRow(0).getLastCellNum() );
        for (int i = 0; i < numrow; i++) {
            for (int j = 0; j < sheet.getRow(i).getLastCellNum(); j++) {
                System.out.print( sheet.getRow(i).getCell(j)+"\t" );
            }
            System.out.println("");

        }
    }
}

