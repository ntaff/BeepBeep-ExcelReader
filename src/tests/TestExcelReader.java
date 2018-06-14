package tests;

import static org.junit.Assert.assertEquals;

import java.util.ArrayList;

import org.junit.Test;

import ca.uqac.lif.cep.Pullable;
import excel.ExcelReader;
import excel.ExcelReaderExample;
import excel.ExcelReaderExceptions;

public class TestExcelReader extends BaseTestExcelReader
{

  public TestExcelReader()
  {
    super();
    // TODO Auto-generated constructor stub
  }
  
  @Test
  public void testProcNum() throws Exception{
    
  String path = "Feuille_de_test.xls";
  System.out.println("\n-----------------------------------\n");
  ExcelReaderExample.createXLSnum();
  ExcelReaderExample.readSheetNum(path);
  ExcelReader test = new ExcelReader(path);

  Pullable p = test.getPullableOutput();

  for (int i = 1; i < 11; i++)
  {
    double x = (Double) p.pull();
   
    assertEquals(i, x, 0.001);
  } 
 }
  
  @Test
  public void testProcString() throws Exception{
    
  String path = "Feuille_de_test.xls";
  System.out.println("\n-----------------------------------\n");
  ExcelReaderExample.createXLSstring();
  ExcelReaderExample.readSheetString(path);
  ExcelReader test = new ExcelReader(path);

  Pullable p = test.getPullableOutput();

  for (int i = 0; i < 11; i++)
  {
    String x = (String) p.pull();
   
    assertEquals(String.valueOf(Character.toChars('a' + i)), x);
  } 
 }
  
  
  @Test (expected = ExcelReaderExceptions.class)
  public void testExtension() throws ExcelReaderExceptions{
    
  String path = "Feuille_de_test.xlss";
 
  //On verifie que le format du fichier est correct
  if (!path.endsWith("xls"))
  {
    throw new ExcelReaderExceptions("Format de fichier incorrect !");
  }
 
}
  
  
  @Test (expected = ExcelReaderExceptions.class)
  public void testColonnes() throws ExcelReaderExceptions{
    
  int tableauColonnes[] = {4,8,12,3,8,-1,14,2,5,5};
  ArrayList<Integer> m_columntab = new ArrayList<Integer>();
 
  //On verifie que les numéros de colonnes sont valides
  for (int i = 0; i < tableauColonnes.length; i++) 
  {
    m_columntab.add(tableauColonnes[i]);
    if(tableauColonnes[i] < 0) 
    {
      throw new ExcelReaderExceptions("Numéro de colonne invalide !");
    }
  }
  
 }
  

}
