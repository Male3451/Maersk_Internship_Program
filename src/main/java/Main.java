
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;


public class Main {
    public static void main(String[] args) {

        XSSFWorkbook workbook=new XSSFWorkbook();

        XSSFSheet sheet= workbook.createSheet();

        Font headerFont=workbook.createFont();
        headerFont.setBold(true);

        CellStyle headerStyle=workbook.createCellStyle();
        headerStyle.setFont(headerFont);

        Row header_row=sheet.createRow(0);
        Cell header_cell= header_row.createCell(0);
        header_cell.setCellValue("Fibo Series    ");
        header_cell.setCellStyle(headerStyle);


        CellStyle oddCellStyle=workbook.createCellStyle();
        oddCellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
        oddCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

         ArrayList<Integer> fibo_series = generateFiboSeries(20);

        for(int i=fibo_series.size()-1;i>=0;i--)
        {
            Row row=sheet.createRow(fibo_series.size()-i);
           Cell cell= row.createCell(0);
           cell.setCellValue(fibo_series.get(i));

           if(fibo_series.get(i)%2==1)
               cell.setCellStyle(oddCellStyle);


        }

        sheet.autoSizeColumn(0);

        try{
            FileOutputStream out =new FileOutputStream(new File("output.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("SuccessFully Write");
        }
        catch(Exception e){
            e.printStackTrace();
        }
    }

    public static  ArrayList<Integer> generateFiboSeries(int n)
    {
     ArrayList<Integer> seris= new ArrayList<Integer>();

     int a=0;
     int b=1;

     seris.add(a);
     seris.add(b);

     while(a+b<n)
     {
        int c=a+b;
        a=b;
        b=c;
        seris.add(c);

     }

     return seris;
    }

}

