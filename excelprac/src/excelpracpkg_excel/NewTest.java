package excelpracpkg_excel;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;
import jxl.*;
import jxl.read.*;
import jxl.write.*;
import java.io.*;


import org.testng.annotations.Test;

public class NewTest {
  @Test
  public void f() {
  }
  public void testImportexport1() throws Exception {
	  FileInputStream fi = new FileInputStream("C:\\Users\\91996\\Desktop\\Book1.xls.xls");
	  Workbook w = Workbook.getWorkbook(fi);
	  Sheet s = w.getSheet(0);
	  String a[][] = new String[s.getRows()][s.getColumns()];
	  FileOutputStream fo = new FileOutputStream("C:\\Users\\91996\\Desktop\\Book2.xls.xls");
	  WritableWorkbook wwb = Workbook.createWorkbook(fo);
	  WritableSheet ws = wwb.createSheet("result1", 0);
	  for (int i = 0; i < s.getRows(); i++)
	  for (int j = 0; j < s.getColumns(); j++)
	  {
	  a[i][j] = s.getCell(j, i).getContents();
	  Label l2 = new Label(j, i, a[i][j]);
	  ws.addCell(l2);
	  Label l1 = new Label(6, 0, "Result");
	  ws.addCell(l1);
	  }
	  for (int i = 1; i < s.getRows(); i++) {
	  for (int j = 2; j < s.getColumns(); j++)

	  {
	  a[i][j] = s.getCell(j, i).getContents();
	  int x=Integer.parseInt(a[i][j]);
	  if(x > 35)
	  {
	  Label l1 = new Label(6, i, "pass");
	  ws.addCell(l1);
	  }
	  else
	  {
	  Label l1 = new Label(6, i, "fail");
	  ws.addCell(l1);
	  break; } }
	  }
	  wwb.write();
	  wwb.close(); } 

}
