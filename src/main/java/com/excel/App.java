package com.excel;
import org.apache.poi.ss.util.WorkbookUtil;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App 
{
	 List<String> listOfDifferences = new ArrayList<>();

	  private static class Locator {
	        Workbook workbook;
	        Sheet sheet;
	        Row row;
	        Cell cell;
	    }
	  private static final String CELL_DATA_DOES_NOT_MATCH = "Cell Data does not Match ::";
	    private static final String CELL_FONT_ATTRIBUTES_DOES_NOT_MATCH = "Cell Font Attributes does not Match ::";
    public static void main( String[] args ) throws EncryptedDocumentException, IOException
    {
        App excelComparator=new App();
        try (Workbook wb1 = WorkbookFactory.create(new File("C:/Users/Hp/Documents/city.xlsx"), null, true)) {
            try (Workbook wb2 = WorkbookFactory.create(new File("C:/Users/Hp/Documents/city1.xlsx"), null, true)) {
                for (String d : App.compare(wb1, wb2)) {
                    System.out.println(d);
                }
            }
        }
    }

        
    private static List<String> compare(Workbook wb1, Workbook wb2) {
		// TODO Auto-generated method stub
    	 App excelComparator=new App();
    	 Locator loc1 = new Locator();
         Locator loc2 = new Locator();
         loc1.workbook = wb1;
         loc2.workbook = wb2;
    	excelComparator.compareSheetData(loc1, loc2);
        return excelComparator.listOfDifferences;
		//return null;
	}




	private void compareSheetData(Locator loc1, Locator loc2) {
		loc1.sheet = loc1.workbook.getSheetAt(0);
        loc2.sheet = loc2.workbook.getSheetAt(0);
        compareDataInSheet(loc1, loc2);
		
	}


	private void compareDataInSheet(Locator loc1, Locator loc2) {
		 for (int j = 0; j <= loc1.sheet.getLastRowNum(); j++) {
	            if (loc2.sheet.getLastRowNum() <j) {
	                return;
	            }
	            loc1.row = loc1.sheet.getRow(j);
	            loc2.row = loc2.sheet.getRow(j);
	            if ((loc1.row == null) || (loc2.row == null)) {
	                continue;
	            }
	            compareDataInRow(loc1, loc2);
	        }
		
	}


	private void compareDataInRow(Locator loc1, Locator loc2) {
		for (int k = 0; k <= loc1.row.getLastCellNum(); k++) {
            if (loc2.row.getLastCellNum() <k) {
                return;
            }
            loc1.cell = loc1.row.getCell(k);
            loc2.cell = loc2.row.getCell(k);
            if ((loc1.cell == null) || (loc2.cell == null)) {
                continue;
            }
            compareDataInCell(loc1, loc2);
        }
		
	}

	 private void addMessage(Locator loc1, Locator loc2, String messageStart, String value1, String value2) {
	        String str =
	            String.format(Locale.ROOT, "%s\nworkbook1 -> %s -> %s [%s] != workbook2 -> %s -> %s [%s]",
	                messageStart,
	                loc1.sheet.getSheetName(), new CellReference(loc1.cell).formatAsString(), value1,
	                loc2.sheet.getSheetName(), new CellReference(loc2.cell).formatAsString(), value2
	            );
	        listOfDifferences.add(str);
	    }
	private void compareDataInCell(Locator loc1, Locator loc2) {
		 if (isCellTypeMatches(loc1, loc2)) {
	            final CellType loc1cellType = loc1.cell.getCellType();
	            switch(loc1cellType) {
	                case BLANK:
	                case STRING:
	                case ERROR:
	                    isCellContentMatches(loc1,loc2);
	                    break;
	                case BOOLEAN:
	                    isCellContentMatchesForBoolean(loc1,loc2);
	                    break;
	                case FORMULA:
	                    isCellContentMatchesForFormula(loc1,loc2);
	                    break;
	                case NUMERIC:
	                    if (DateUtil.isCellDateFormatted(loc1.cell)) {
	                        isCellContentMatchesForDate(loc1,loc2);
	                    } else {
	                        isCellContentMatchesForNumeric(loc1,loc2);
	                    }
	                    break;
	                default:
	                    throw new IllegalStateException("Unexpected cell type: " + loc1cellType);
	            
	            }
		 }
	}
		 private void isCellContentMatches(Locator loc1, Locator loc2) {
		        String str1 = loc1.cell.toString();
		        String str2 = loc2.cell.toString();
		        if (!str1.equals(str2)) {
		            addMessage(loc1,loc2,CELL_DATA_DOES_NOT_MATCH,str1,str2);
		        }
		    }
		    /**
		     * Checks if cell content matches for boolean.
		     */
		    private void isCellContentMatchesForBoolean(Locator loc1, Locator loc2) {
		        boolean b1 = loc1.cell.getBooleanCellValue();
		        boolean b2 = loc2.cell.getBooleanCellValue();
		        if (b1 != b2) {
		            addMessage(loc1,loc2,CELL_DATA_DOES_NOT_MATCH,Boolean.toString(b1),Boolean.toString(b2));
		        }
		    }
		    /**
		     * Checks if cell content matches for date.
		     */
		    private void isCellContentMatchesForDate(Locator loc1, Locator loc2) {
		        Date date1 = loc1.cell.getDateCellValue();
		        Date date2 = loc2.cell.getDateCellValue();
		        if (!date1.equals(date2)) {
		            addMessage(loc1, loc2, CELL_DATA_DOES_NOT_MATCH, date1.toString(), date2.toString());
		        }
		    }
		    /**
		     * Checks if cell content matches for formula.
		     */
		    private void isCellContentMatchesForFormula(Locator loc1, Locator loc2) {
		        // TODO: actually evaluate the formula / NPE checks
		        String form1 = loc1.cell.getCellFormula();
		        String form2 = loc2.cell.getCellFormula();
		        if (!form1.equals(form2)) {
		            addMessage(loc1, loc2, CELL_DATA_DOES_NOT_MATCH, form1, form2);
		        }
		    }
		    /**
		     * Checks if cell content matches for numeric.
		     */
		    private void isCellContentMatchesForNumeric(Locator loc1, Locator loc2) {
		        // TODO: Check for NaN
		        double num1 = loc1.cell.getNumericCellValue();
		        double num2 = loc2.cell.getNumericCellValue();
		        if (num1 != num2) {
		            addMessage(loc1, loc2, CELL_DATA_DOES_NOT_MATCH, Double.toString(num1), Double.toString(num2));
		        }
		    }

		    private boolean isCellTypeMatches(Locator loc1, Locator loc2) {
		        CellType type1 = loc1.cell.getCellType();
		        CellType type2 = loc2.cell.getCellType();
		        if (type1 == type2) {
		            return true;
		        }
		        addMessage(loc1, loc2,
		            "Cell Data-Type does not Match in :: ",
		            type1.name(), type2.name()
		        );
		        return false;
		    }

}
