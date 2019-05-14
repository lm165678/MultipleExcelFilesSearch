import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.regex.Pattern;

public class Test {

    public static void main(String[] args) {




  String d = " SADig , fdfsf".toLowerCase().trim();

  String check = "sadig".toLowerCase();


  String modified = d.replaceAll("((\\s*,\\s*)|\\s+)"," ");

        System.out.println(modified
        );


  }





    }
