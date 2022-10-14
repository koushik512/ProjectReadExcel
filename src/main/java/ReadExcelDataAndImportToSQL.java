import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelDataAndImportToSQL
{
    public static void main(String[] args)
    {
        try
        {
            // Give the Absolute Path of your file here
            // Please verify the sql files are generated in resources module.
            File file = new File("src/main/resources/Details.xlsx");
            String fileName = file.getName().split("\\.")[0];
            FileInputStream inputStream = new FileInputStream(file);
            XSSFWorkbook wb=new XSSFWorkbook(inputStream);
            Sheet sheet=wb.getSheetAt(0);
            FormulaEvaluator formulaEvaluator=wb.getCreationHelper().createFormulaEvaluator();
            // Headerlist
            List<String> rowHeaderFromExcel = new ArrayList<>();
            // retrive headers into one list :
            Row firstRow = sheet.getRow(0);
            for(Cell cell: firstRow)    //iteration over cell using for each loop
            {
                switch(formulaEvaluator.evaluateInCell(cell).getCellType())
                {
                    case NUMERIC:
                        rowHeaderFromExcel.add(String.valueOf(cell.getNumericCellValue()));
                        break;
                    case STRING:
                        rowHeaderFromExcel.add(String.valueOf(cell.getStringCellValue()));
                        break;
                    case BOOLEAN:
                        rowHeaderFromExcel.add(String.valueOf(cell.getBooleanCellValue()));
                        break;
                }
            }
//            System.out.println(fileName);
//            System.out.println(rowHeaderFromExcel);
            // this method is for create query :
            createQuery(rowHeaderFromExcel, fileName);


            // List(List(string) for body:
            List<List<String>> rowBodyFromExcel = new ArrayList<>();
            for(Row row: sheet)     //iteration over row using for each loop
            {
                List<String> rowCells = new ArrayList<>();
                for(Cell cell: row)    //iteration over cell using for each loop
                {
                    switch(formulaEvaluator.evaluateInCell(cell).getCellType())
                    {
                        case NUMERIC:
                            rowCells.add(String.valueOf(cell.getNumericCellValue()));
                            break;
                        case STRING:
                            rowCells.add(String.valueOf(cell.getStringCellValue()));
                            break;
                        case BOOLEAN:
                            rowCells.add(String.valueOf(cell.getBooleanCellValue()));
                            break;
                    }
                }
                rowBodyFromExcel.add(rowCells);
            }
//            System.out.println(rowBodyFromExcel);
            createInsertQueries(rowBodyFromExcel, fileName);

            wb.close();
            inputStream.close();
        }
        catch(Exception e){
            e.printStackTrace();
        }
    }
    private static void createQuery(List<String> rowHeaderFromExcel, String fileName)
    {
        StringBuilder sb = new StringBuilder();
        sb.append("Create Table "+fileName + " ( ");
        sb.append(fileName+"_id serial PRIMARY KEY, ");
        for(int i = 0 ; i <rowHeaderFromExcel.size() ;i++){
            sb.append(rowHeaderFromExcel.get(i)+" VARCHAR(200) NOT NULL,");
        }
        String createQuery = sb.substring(0, sb.length()-1)+");";
        System.out.println(createQuery);

        Path path= Paths.get("src/main/resources/Create.sql");
        try{
            Files.writeString(path, createQuery, StandardCharsets.UTF_8);
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }

    }


    private static void createInsertQueries(List<List<String>> rowBodyFromExcel, String fileName)
    {
        StringBuilder insertQueryBuilder = new StringBuilder();
        insertQueryBuilder.append("INSERT INTO "+fileName+ " ( ");
        List<String> rowHeader = rowBodyFromExcel.get(0);
        for(int i = 0; i < rowHeader.size(); i++){
            insertQueryBuilder.append(" " + rowHeader.get(i)+" ,");
        }
        String insertQuery = insertQueryBuilder.substring(0, insertQueryBuilder.length()-1 )+
                (")  VALUES (");
//        System.out.println("-->"+insertQuery);


        List<String> insertQueries = new ArrayList<>();

        for(int i = 1; i < rowBodyFromExcel.size(); i++)
        {
            StringBuilder valuesBuilder = new StringBuilder();
            List<String> rowCell = rowBodyFromExcel.get(i);
            for(int j=0; j < rowCell.size(); j++ )
            {
                valuesBuilder.append(" \""+rowBodyFromExcel.get(i).get(j)+"\" ,");
            }
            //System.out.println(insertQuery + valuesBuilder.substring(0,valuesBuilder.length()-1)+");") ;
            insertQueries.add(insertQuery + valuesBuilder.substring(0,valuesBuilder.length()-1)+");");
        }
        System.out.println("List of insert Queries---\n "+ insertQueries);

        String finalInsertQueries = "";
        for(int i  = 0; i < insertQueries.size() ; i++){
            finalInsertQueries = finalInsertQueries + insertQueries.get(i)+"\n";
        }

        System.out.println("List of insert Queries---"+ finalInsertQueries.substring(0, finalInsertQueries.length()-1));

        // Save the queries in the insert.sql file
        Path path= Paths.get("src/main/resources/Insert.sql");
        try{
            Files.writeString(path, finalInsertQueries, StandardCharsets.UTF_8);
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }

}
