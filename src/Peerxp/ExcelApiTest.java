package Peerxp;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelApiTest {
	public FileInputStream fis = null;
	public FileOutputStream fos = null;
	public XSSFWorkbook workbook = null;
	public XSSFSheet sheet = null;
	public XSSFRow row = null;
	public XSSFCell cell = null;
	String xlFilePath;

	public ExcelApiTest(String xlFilePath) throws Exception
	{
		this.xlFilePath= xlFilePath;
		fis=new FileInputStream(xlFilePath);
		workbook= new XSSFWorkbook(fis);
		fis.close();
	}
	public String getcellData(String SheetName,String colName,int rowNum)
	{
		try
		{
			int colNum=-1;
			sheet = workbook.getSheet(SheetName);
			row = sheet.getRow(1);
			for(int i=0;i<row.getLastCellNum();i++)
			{
				if(row.getCell(i).getStringCellValue().trim().equals(colName.trim()))
					colNum= i;
			}
		    row = sheet.getRow(rowNum-1);
		    cell = row.getCell(colNum);
		    
		    if(cell.getCellTypeEnum()== CellType.STRING)
		    	return cell.getStringCellValue();
		    else if (cell.getCellTypeEnum()==CellType.NUMERIC || cell.getCellTypeEnum()==CellType.FORMULA)
		    {
		    	String cellvalue = new BigDecimal(cell.getNumericCellValue()).toPlainString();

		    	if(HSSFDateUtil.isCellDateFormatted(cell))
		    	{
		    		DateFormat df = new SimpleDateFormat("dd/mm/yy");
		    		Date date = cell.getDateCellValue();
		    		cellvalue = df.format(date);
		    	}
		    	return cellvalue;
		    }else if(cell.getCellTypeEnum()==CellType.BLANK)
		    	return " ";
		    else
		    	return String.valueOf(cell.getBooleanCellValue());
		}
		catch (Exception e)
		{
			e.printStackTrace();
			return "row"+rowNum+" orcolumn"+colName+"does not exit in Excel";
		}
		
	}
	
	public String getcellData(String SheetName, int colNum,int rowNum)
	{
		try
		{
			sheet = workbook.getSheet(SheetName);
			row = sheet.getRow(rowNum);
			cell = row.getCell(colNum);
			if(cell.getCellTypeEnum()==CellType.STRING)
			return cell.getStringCellValue();
			else if (cell.getCellTypeEnum()==CellType.NUMERIC ||cell.getCellTypeEnum()==CellType.FORMULA)
			{
				String cellValue = new BigDecimal(cell.getNumericCellValue()).toPlainString();
				if(HSSFDateUtil.isCellDateFormatted(cell))
				{
					DateFormat df = new SimpleDateFormat("dd/mm/yy");
					Date date = cell.getDateCellValue();
					cellValue = df.format(date);
				}
			return cellValue;
			}else if(cell.getCellTypeEnum()==CellType.BLANK)
				return " ";
			else
				return String.valueOf(cell.getBooleanCellValue());
			}
		catch(Exception e)
		{
			e.printStackTrace();
			return "row"+rowNum+" orcolumn"+colNum+"does not exit in Excel";
		}
	}
	public boolean setcellData(String SheetName, int colNumber,int rowNum,String value)
	{
		try
		{
			sheet = workbook.getSheet(SheetName);
			row = sheet.getRow(rowNum);
			if(row == null)
				row = sheet.createRow(rowNum);
				cell = row.getCell(colNumber);
			if(cell == null)
				cell = row.createCell(colNumber);
				cell.setCellValue(value);
				fos = new FileOutputStream(xlFilePath);
				workbook.write(fos);
				fos.close();
		}
		catch (Exception ex)
		{
			ex.printStackTrace();
			return false;
		}
	return true;
	}
	
	public boolean setcellData(String SheetName, String colName,int rowNum,String value)
	{
		try
		{
			int colNum=-1;
			sheet = workbook.getSheet(SheetName);
			row = sheet.getRow(1);
			for(int i=0;i<row.getLastCellNum();i++)
			{
				if(row.getCell(i).getStringCellValue().trim().equals(colName))
				{
					colNum=i;
				}
			}
			sheet.autoSizeColumn(colNum);
			row = sheet.getRow(rowNum-1);
			if(row==null)
				row = sheet.createRow(rowNum-1);
				cell = row.getCell(colNum);
			if(cell==null)
				cell = row.createCell(colNum);
			    cell.setCellValue(value);
				fos = new FileOutputStream(xlFilePath);
				workbook.write(fos);
				fos.close();
		}
		catch(Exception ex)
		{
			ex.printStackTrace();
			return false;
		}
		return true;
	}
	
	public int getRowCount(String SheetName)
	{
		sheet = workbook.getSheet(SheetName);
		int rowCount = sheet.getLastRowNum()+1;
		return rowCount;
	}
	public int getcolumnCount(String SheetName)
	{
		sheet = workbook.getSheet(SheetName);
		row = sheet.getRow(0);
		int colCount = row.getLastCellNum();
		return colCount;
	}
}


