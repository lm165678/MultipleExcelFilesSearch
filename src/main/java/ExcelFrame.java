
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.Font;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Pattern;

public class ExcelFrame extends JFrame {

    String name;
    String result;


    XSSFSheet sheet;
    ArrayList<Integer> listofRows;
    XSSFSheet newsheet;
    XSSFWorkbook newWorkbook;
    XSSFRow newRow;
    XSSFCell newCell;

    JFileChooser chooser;
    File[] files;

    JButton start ;
    ProgressMonitor progressMonitor ;
    Timer timer ;
    SearchWork searchWork;
    JButton open;
    JButton export;
    JLabel label;
    JTextField text;
    JTextField resultText;

    JButton search;
    JLabel enterKey;
    JLabel openfile;
    JLabel exportF;



    GridLayout gl;
    Font fL;
    Font fT;

    public ExcelFrame() {


fL = new Font("Courier",Font.BOLD,35);
fT = new Font("Courier",Font.PLAIN,25);


resultText = new JTextField("insert here");
resultText.setFont(fT);

        text = new JTextField("insert here",20);
        text.setFont(fT);

        enterKey = new JLabel("Enter keyword");
        enterKey.setFont(fL);

        openfile = new JLabel("Chose files");
        openfile.setFont(fL);

        exportF = new JLabel("Create new File");
        exportF.setFont(fL);

        search = new JButton("Search");
        search.setFont(fT);

        open = new JButton("open");
        open.setFont(fT);

        start = new JButton("start");
        start.setFont(fT);

        export = new JButton("export");
        export.setFont(fT);



        export.addActionListener(e ->{
            export(newWorkbook);
            result = resultText.getText();
        });

        FileNameExtensionFilter filter = new FileNameExtensionFilter("fles", "xlsx");
        chooser = new JFileChooser();
        chooser.setFileFilter(filter);
        chooser.setCurrentDirectory(new File("."));
        chooser.setMultiSelectionEnabled(true);



        open.addActionListener(e -> openFile());

        search.addActionListener(e -> {
            name = text.getText().trim();
        });



        start.addActionListener(e -> {
            start.setEnabled(false);
            int max = files.length;
            searchWork = new SearchWork();
            searchWork.execute();
            progressMonitor = new ProgressMonitor(ExcelFrame.this,"status",null,0,max);
            timer.start();

        });


        text.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
                text.setText("");
            }
        });

        resultText.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
                resultText.setText("");
            }
        });



        timer = new Timer(500,e -> {
            if (progressMonitor.isCanceled()){
                searchWork.cancel(true);
                start.setEnabled(true);
            }else if (searchWork.isDone()){
                progressMonitor.close();
                start.setEnabled(true);
            }else {
                progressMonitor.setProgress(searchWork.getProgress());
            }
        });






        gl = new GridLayout(3,3);



        JPanel jp = new JPanel();
        jp.setLayout(gl);

        jp.add(enterKey);
        jp.add(text);
        jp.add(search);

        jp.add(openfile);
        jp.add(open);
        jp.add(start);

        jp.add(exportF);
        jp.add(resultText);
        jp.add(export);



        add(jp);





    }





    public void openFile() {     // This method selects file , and conversts it to XSSFWorkbook
        int i = chooser.showOpenDialog(this);
        if (i != JFileChooser.APPROVE_OPTION) {
            return;
        }
        files = chooser.getSelectedFiles();


    }





    class SearchWork extends SwingWorker<Boolean,Integer>{


        @Override
        protected Boolean doInBackground() throws Exception {


            int counter = 0;
            newWorkbook = new XSSFWorkbook();
            newsheet = newWorkbook.createSheet();

            for (File file : files) {
                try (FileInputStream in = new FileInputStream(file);
                     XSSFWorkbook workbook = new XSSFWorkbook(in)) {

                    for (int z = 0; z < workbook.getNumberOfSheets(); z++) {


                        sheet = workbook.getSheetAt(z);


                        if (Pattern.matches("^\\d*\\.\\d*||\\d*", name)) {
                            listofRows = selectNumber(sheet, Double.parseDouble(name));
                        } else {
                            listofRows = selectString(sheet, name);
                        }

                        System.out.println(listofRows.get(0));

                        for (Integer integer : listofRows) {
                            System.out.println(integer);
                        }



                        if (!listofRows.isEmpty()) {
                            for (int c = 0; c < listofRows.size(); c++) {
                                Row source = sheet.getRow(listofRows.get(c));
                                newRow = newsheet.createRow(counter);
                                counter++;
                                publish(counter);

                                for (int b = 0; b < source.getLastCellNum(); b++) {

                                    System.out.println(source.getLastCellNum());

                                    newCell = newRow.createCell(b);

                                    if (source.getCell(b) != null) {
                                        switch (source.getCell(b).getCellType()) {
                                            case STRING:
                                                newCell.setCellValue(source.getCell(b).getRichStringCellValue());
                                                break;
                                            case NUMERIC:
                                                newCell.setCellValue(source.getCell(b).getNumericCellValue());
                                                System.out.println(" i am here");
                                                break;
                                        }

                                    }

                                }
                            }

                        }

                    }


                } catch (IOException e) {


                }

            }

            return true;

        }
    }


 void export(XSSFWorkbook workbook){

try{
    FileOutputStream fileOutputStream = new FileOutputStream(result);
    workbook.write(fileOutputStream);
    workbook.close();
}catch(IOException e){

}
}

    static ArrayList<Integer> selectString(XSSFSheet sheet,String key){
       ArrayList<Integer> list = new ArrayList<>();
        for (Row row:sheet){

            for (Cell cell : row){


                switch (cell.getCellType()){
                    case STRING:  if (cell.getRichStringCellValue().toString().trim().equalsIgnoreCase(key)){
                        list.add(cell.getRowIndex());
                    }
                    break;
                    case NUMERIC:
                    break;
                    case BLANK:
                        System.out.println("empty");
                        break;
                }
            }

        }
        return list;
    }

    static ArrayList<Integer> selectNumber(XSSFSheet sheet,double key){
        ArrayList<Integer> list = new ArrayList<>();
        for (Row row:sheet){

            for (Cell cell : row){


                switch (cell.getCellType()){
                    case STRING:
                        break;
                    case NUMERIC:
                        if (cell.getNumericCellValue() == key ){
                            list.add(cell.getRowIndex());
                        }
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
