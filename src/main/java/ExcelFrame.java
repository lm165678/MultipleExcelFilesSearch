import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.*;
import java.util.ArrayList;

public class ExcelFrame extends JFrame {

    JFileChooser chooser;
    File[] files;

    public ExcelFrame(){
        //=========================================//

FileNameExtensionFilter filter = new FileNameExtensionFilter("fles","xlsx");
        chooser = new JFileChooser();
        chooser.setFileFilter(filter);
        chooser.setCurrentDirectory(new File("."));
        chooser.setMultiSelectionEnabled(true);
        JButton open = new  JButton("open");
        add(open, BorderLayout.CENTER);
        open.addActionListener(e -> openFile());
    }


    public void openFile()  {     // This method selects file , and conversts it to XSSFWorkbook
        int i = chooser.showOpenDialog(this);
        if (i!= JFileChooser.APPROVE_OPTION){
            return;
        }
        files = chooser.getSelectedFiles();
      String wordName = "Sadig";


        for (File file:files){
           try(FileInputStream in = new FileInputStream(file);
               XSSFWorkbook workbook = new XSSFWorkbook(in)) {
               XSSFSheet sheet = workbook.getSheetAt(0);

               ArrayList<Integer> listofRows = select(sheet,"Sadig");

               System.out.println(listofRows.get(0));

               for (Integer integer : listofRows){
                   System.out.println(integer);
               }


               XSSFWorkbook newWorkbook = new XSSFWorkbook();
               XSSFSheet newsheet = newWorkbook.createSheet();
               XSSFRow newRow ;
               XSSFCell newCell;

               String result;

               for (int c = 0 ; c < listofRows.size(); c++){
                   Row source = sheet.getRow(listofRows.get(c));
                   newRow = newsheet.createRow(c);

                   for (int b = 0 ; b < source.getLastCellNum();b++){

                       newCell = newRow.createCell(b);

                       if (source.getCell(b) == null){
                           result = " " ;
                       }else {
                           result = source.getCell(b).getRichStringCellValue().getString();
                       }
                       newCell.setCellValue(result);
                   }
               }

FileOutputStream fileOutputStream = new FileOutputStream("result.xlsx");
               newWorkbook.write(fileOutputStream);
               newWorkbook.close();

           }catch (IOException e){


           }

        }



    }




    public static ArrayList<Integer> select(XSSFSheet sheet,String name){
       ArrayList<Integer> list = new ArrayList<>();
        for (Row row:sheet){

            for (Cell cell : row){


                if (cell.getRichStringCellValue().toString().trim().equals(name)){
                    list.add(cell.getRowIndex());
                }

            }

        }
        return list;
    }
}
