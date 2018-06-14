package tests;

import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;

import java.io.File;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public abstract class BaseTestExcelReader {
    
    public BaseTestExcelReader()
    {
      // TODO Auto-generated constructor stub
    }

    @BeforeClass
    /**
     * Fonction executée avant les tests : On crée une feuille de test contenant des données sous forme de nombre
     * @param : null
     * @return : void
     */
    public static void setUpBeforeClass() throws Exception {
      
      final int nbRow = 5;
      final int nbColumns = 5;
      String path = "Feuille_de_test.xls";

     // Workbook wb = _excelTestProvider.createWorkbook();
      Workbook wb = new HSSFWorkbook();
      Sheet sheet = wb.createSheet(path);
      int k = 0;

      for (int i = 0; i < nbRow; i++)
      {
        Row row1 = sheet.createRow(i);
        for (int j = 0; j < nbColumns; j++)
        {
          k++;
          row1.createCell(j).setCellValue(k);

        } 
      } 
    }
    
    @AfterClass
    /**
     * Fonction executée après les tests : On supprime la feuille de test
     * @param : null
     * @return : void
     */
    public static void setUpAfterClass() throws Exception {
      
      File feuilleTest = new File("Feuille_de_test.xls");
      feuilleTest.delete();
    }
    
    @Before
    public void setUpBefore() throws Exception {
      
      
    }

    @After
    public void setUpAfter() throws Exception {
      
      
    }

   
}
