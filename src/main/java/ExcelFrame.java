import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.Font;
import java.io.*;
import java.util.ArrayList;

public class ExcelFrame extends JFrame {

    XSSFSheet newsheet;
    XSSFWorkbook newWorkbook;
    XSSFRow newRow ;
    XSSFCell newCell;
static int counter = 0;
    JFileChooser chooser;
    String name;
    File[] files;

    public ExcelFrame(){
        //=========================================//

FileNameExtensionFilter filter = new FileNameExtensionFilter("fles","xlsx");
        chooser = new JFileChooser();
        chooser.setFileFilter(filter);
        chooser.setCurrentDirectory(new File("."));
        chooser.setMultiSelectionEnabled(true);
        JButton open = new  JButton("open");
        add(open, BorderLayout.SOUTH);
        open.addActionListener(e -> openFile());


        JLabel label = new JLabel();
        label.setText("Search APP");
        label.setFont(new Font("arial",72,72));
        add(label,BorderLayout.NORTH);


        JTextField text = new JTextField("insert here");
        add(text,BorderLayout.CENTER);

        JButton search = new JButton("click for Search");
        add(search,BorderLayout.EAST);
        search.addActionListener(e -> {
            name = text.getText();
        });

        JButton export = new JButton("Export ");
        add(export,BorderLayout.WEST);
        export.addActionListener(e -> export(newWorkbook));

    }

    public void openFile()  {     // This method selects file , and conversts it to XSSFWorkbook
        int i = chooser.showOpenDialog(this);
        if (i!= JFileChooser.APPROVE_OPTION){
            return;
        }
        files = chooser.getSelectedFiles();
        newWorkbook = new XSSFWorkbook();
        newsheet = newWorkbook.createSheet();

        for (File file:files){

           try(FileInputStream in = new FileInputStream(file);
               XSSFWorkbook workbook = new XSSFWorkbook(in)) {
               XSSFSheet sheet = workbook.getSheetAt(0);
               ArrayList<Integer> listofRows = select(sheet,name);

               System.out.println(name);
               System.out.println(listofRows.get(0));

               for (Integer integer : listofRows){
                   System.out.println(integer);
               }


               for (int c = 0 ; c < listofRows.size(); c++){
                   Row source = sheet.getRow(listofRows.get(c));
                   newRow = newsheet.createRow(counter);
                   counter++;

                   for (int b = 0 ; b < source.getLastCellNum();b++){

                       System.out.println(source.getLastCellNum());

                       newCell = newRow.createCell(b);

                       if (source.getCell(b)!=null){
                           switch (source.getCell(b).getCellType()){
                               case STRING: newCell.setCellValue(source.getCell(b).getRichStringCellValue());
                                   break;
                               case NUMERIC: newCell.setCellValue(source.getCell(b).getNumericCellValue());
                                   System.out.println(" i am here");
                                   break;
                           }

                       }

                   }
               }

           }catch (IOException e){


           }

        }

    }


static void export(XSSFWorkbook workbook){

try{
    FileOutputStream fileOutputStream = new FileOutputStream("result.xlsx");
    workbook.write(fileOutputStream);
    workbook.close();
}catch(IOException e){

}


}

    static ArrayList<Integer> select(XSSFSheet sheet,String name){
       ArrayList<Integer> list = new ArrayList<>();
        for (Row row:sheet){

            for (Cell cell : row){


                switch (cell.getCellType()){
                    case STRING:  if (cell.getRichStringCellValue().toString().trim().equals(name)){
                        list.add(cell.getRowIndex());
                    }
                    break;
                    case NUMERIC:
                        System.out.println("numeric");
                    break;
                    case BLANK:
                        System.out.println("empty");
                        break;
                }
            }

        }
        return list;
    }
}
