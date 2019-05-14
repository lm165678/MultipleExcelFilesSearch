
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import javax.imageio.ImageIO;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.Color;
import java.awt.Font;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.MouseListener;
import java.io.*;
import java.lang.reflect.Array;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.regex.Pattern;



public class ExcelFrame extends JFrame {

    String name;  //keyword
    String result; // resulting file


    XSSFSheet sheet;
    ArrayList<Integer> listofRows;
    XSSFSheet newsheet;
    XSSFWorkbook newWorkbook;
    XSSFRow newRow;
    XSSFCell newCell;

    JFileChooser chooser;
    File[] files;

    JButton start ;
    JButton open;
    JButton export;
    JButton search;


    ProgressMonitor progressMonitor ;
    Timer timer ;
    SearchWork searchWork;


    JLabel label;
    JTextField text;
    JTextField resultText;
    JLabel enterKey;
    JLabel openfile;
    JLabel exportF;
    JLabel dinamicstatus;


    JLabel status;

    Font fL;
    Font fT;



    String checkString;
    public ExcelFrame() {




        fL = new Font("Arial",Font.BOLD,35);
        fT = new Font("Arial",Font.PLAIN,25);

resultText = new JTextField("insert here");
text = new JTextField("insert here",20);


        resultText = new JTextField("insert here");
        resultText.setFont(fT);

        text = new JTextField("insert here",20);
        text.setFont(fT);

        enterKey = new JLabel("Enter keyword");
        enterKey.setFont(fL);
        enterKey.setForeground(Color.WHITE);

        status = new JLabel("Status");
        status.setFont(fL);
        status.setForeground(Color.WHITE);

        dinamicstatus = new JLabel("");
        dinamicstatus.setFont(fL);
        dinamicstatus.setForeground(Color.WHITE);

        openfile = new JLabel("Chose files");
        openfile.setFont(fL);
        openfile.setForeground(Color.WHITE);

        exportF = new JLabel("Create File");
        exportF.setFont(fL);
        exportF.setForeground(Color.WHITE);

        search = new JButton("search");
        search.setPreferredSize(new Dimension(30,20));
        search.setFont(fT);

        open = new JButton("open");
        open.setFont(fT);

        start = new JButton("start");
        start.setFont(fT);

        export = new JButton("export");
        export.setFont(fT);



        export.addActionListener(e ->{
            result = resultText.getText() + ".xlsx";
            export(newWorkbook);

        });

        FileNameExtensionFilter filter = new FileNameExtensionFilter("fles", "xlsx");
        chooser = new JFileChooser();
        chooser.setFileFilter(filter);
        chooser.setCurrentDirectory(new File("."));
        chooser.setMultiSelectionEnabled(true);


        search.addActionListener(e -> {
            name = text.getText().trim();

        });

        open.addActionListener(e -> openFile());

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






//        gl = new GridLayout(3,3);



        CustomPanel jp = new CustomPanel();
      jp.setLayout(null);


        enterKey.setBounds(10,20,300,50);

        jp.add(enterKey);
        text.setBounds(350,20,300,50);
        jp.add(text);
        search.setBounds(700,20,200,50);
        jp.add(search);



        openfile.setBounds(10,90,300,50);

        jp.add(openfile);

        open.setBounds(350,90,200,50);
        jp.add(open);

        start.setBounds(700,90,200,50);
        jp.add(start);


        exportF.setBounds(10,160,300,50);

        jp.add(exportF);

        resultText.setBounds(350,160,300,50);
        jp.add(resultText);

        export.setBounds(700,160,200,50);
        jp.add(export);


        status.setBounds(10,400,150,50);
        jp.add(status);

        dinamicstatus.setBounds(200,400,200,50);
        jp.add(dinamicstatus);




        add(jp);

    }


    //===================================================================================================/
    //
    //
    //
    //
    //
    // /




    public void openFile() {     // This method selects file , and conversts it to XSSFWorkbook
        int i = chooser.showOpenDialog(this);
        if (i != JFileChooser.APPROVE_OPTION) {
            return;
        }
        files = chooser.getSelectedFiles();
    }


    class SearchWork extends SwingWorker<Boolean,Integer>{
        @Override
        protected Boolean doInBackground(){
            int counter = 0;
            newWorkbook = new XSSFWorkbook();
            newsheet = newWorkbook.createSheet();
            for (File file : files) {
                try (FileInputStream in = new FileInputStream(file);
                     XSSFWorkbook workbook = new XSSFWorkbook(in)
                ) {
                    for (int z = 0; z < workbook.getNumberOfSheets(); z++) {

                        sheet = workbook.getSheetAt(z);
                        System.out.println(workbook.getNumberOfSheets());



                     name = name.replaceAll("\\s+"," ").toLowerCase().trim();

                        listofRows = selectString(name);



                        if (!listofRows.isEmpty()) {

                            for (int c = 0; c < listofRows.size(); c++) {
                                Row source = sheet.getRow(listofRows.get(c));
                                newRow = newsheet.createRow(counter);
                                counter++;
                                publish(counter);

                                for (int b = 0; b < source.getLastCellNum(); b++) {


                                    newCell = newRow.createCell(b);


                                    if (source.getCell(b) != null) {
                                        switch (source.getCell(b).getCellType()) {
                                            case STRING:
                                                newCell.setCellValue(source.getCell(b).getRichStringCellValue());
                                                break;
                                            case NUMERIC:
                                                if (HSSFDateUtil.isCellDateFormatted(source.getCell(b))){

                                                    newCell.setCellValue(source.getCell(b).getDateCellValue());

                                                }else {
                                                    newCell.setCellValue(source.getCell(b).getNumericCellValue());

                                                }

                                                break;
                                        }

                                    }

                                }
                            }

                        } else {
                            continue;
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

    ArrayList<Integer> selectString(String key){
        ArrayList<Integer> list = new ArrayList<>();
        DataFormatter dataFormatter = new DataFormatter();
        for (Row row:sheet){
            for (Cell cell : row){
                switch (cell.getCellType()){
                    case STRING:
                        String cellValue = dataFormatter.formatCellValue(cell);
                        String  newcellValue = cellValue.replaceAll("((\\s*,\\s*)|\\s+)"," ").trim().toLowerCase();

                        dinamicstatus.setText(newcellValue);

                        if (Pattern.matches(key+"\\s\\w+",newcellValue)){
                            list.add(cell.getRowIndex());
                        }else if (Pattern.matches("\\w+\\s"+key,newcellValue)){
                            list.add(cell.getRowIndex());
                        }
                        else if (Pattern.matches(key+"\\s\\w+\\s\\w+",newcellValue)){
                            list.add(cell.getRowIndex());
                        }else if (Pattern.matches("\\w+\\s"+key+"\\s\\w+",newcellValue)){
                            list.add(cell.getRowIndex());
                        }else if (Pattern.matches("\\w+\\s\\w+\\s"+key,newcellValue)){
                            list.add(cell.getRowIndex());
                        }
                        else if(Pattern.matches(key,newcellValue)){
                            list.add(cell.getRowIndex());
                        }
                        break;
                }




            }
        }
        return list;
    }
}











 class CustomPanel extends JPanel{
     Image image;
     CustomPanel(){

        ImageIcon icon  = new ImageIcon(getClass().getClassLoader().getResource("back.jpg"));
        image = icon.getImage();
    }
    @Override
     protected  void paintComponent(Graphics g){

        super.paintComponent(g);
        g.drawImage(image,0,0,this);
    }
}
