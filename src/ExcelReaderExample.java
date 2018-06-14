package ca.uqac.lif.cep.excelReader;

import java.io.FileOutputStream;
import java.io.OutputStream;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;

import ca.uqac.lif.cep.Pullable;

/**
 * Permet de créer un fichier Excel rempli de valeurs numériques pour tester le processeur ExcelReader
 **/

public class ExcelReaderExample
{
  @SuppressWarnings("resource")
  
  public static Workbook createXLSnum() {
    
    // Constantes permettant de modifier le nombre de lignes et de colonnes du
    // fichier
    final int nbRow = 5;
    final int nbColumns = 5;

    // On créer un nouveau classeur
    Workbook wb = new HSSFWorkbook();

    // On créer des nouvelles feuilles
    Sheet num = wb.createSheet("Numeric");
 

    // Compteur de cellules
    int k = 0;

    // On crée X lignes
    for (int i = 0; i < nbRow; i++)
    {

      // On créer des nouvelles lignes
      Row rowNum = num.createRow(i);

      // On créer X colonnes
      for (int j = 0; j < nbColumns; j++)
      {

        // On incrémente le compteur de cellule avant chaque création de cellule
        k++;

        // On créer des cellules et on leur ajoute du contenu
        rowNum.createCell(j).setCellValue(k);
    

      } // On ferme le second for
    } // On ferme le premier for
    
    return wb ;
  }
  
  
  public static Workbook createXLSstring() {
    
    // Constantes permettant de modifier le nombre de lignes et de colonnes du
    // fichier
    final int nbRow = 5;
    final int nbColumns = 5;

    // On créer un nouveau classeur
    Workbook wb = new HSSFWorkbook();

    // On créer des nouvelles feuilles
    Sheet string = wb.createSheet("String");
 

    // Compteur de cellules
    int k = 0;

    // On crée X lignes
    for (int i = 0; i < nbRow; i++)
    {

      // On créer des nouvelles lignes
      Row rowString = string.createRow(i);

      // On créer X colonnes
      for (int j = 0; j < nbColumns; j++)
      {

        // On incrémente le compteur de cellule avant chaque création de cellule
        k++;

        // On créer des cellules et on leur ajoute du contenu

        rowString.createCell(j).setCellValue(String.valueOf(Character.toChars('a' + (k-1))));
    

      } // On ferme le second for
    } // On ferme le premier for
    
    return wb ;
  }
  public static void readSheetNum(String path)throws Exception
  {
    
  
  Workbook wb = createXLSnum();
  ExcelReaderExceptions ExcelExceptions = null;
  // On écrit le contenu dans le fichier de sortie
  try {OutputStream fileOut = new FileOutputStream(path);
 
        {
          wb.write(fileOut);
        } 
  
      } catch (ExcelReaderExceptions e) {
        ExcelExceptions = e;
        throw e;
      }
  finally {
  // Objets permettant de formatter le contenu d'une cellule
  DataFormatter formatter = new DataFormatter();

  // On récupère la feuille courante
  Sheet sheetNum = wb.getSheetAt(0);

  // On parcoure les colonnes
  for (Row row1 : sheetNum)
  {

    // On parcoure les lignes
    for (Cell cell : row1)
    {

      // On récupère la localisation de la cellule
      CellReference cellRef = new CellReference(row1.getRowNum(), cell.getColumnIndex());

      // On affiche la localisation de la cellule
      System.out.print(cellRef.formatAsString());

      // Pour séparer la localisation de la cellule de son contenu
      System.out.print(" - ");

      // On recupère le contenu brut de la cellule
      String text = formatter.formatCellValue(cell);

      // On affiche ce contenu
      System.out.println(text);

    } // On ferme le second for
  } // On ferme le premier for
  
  System.out.println("\n-----------------------------------\n");

 }
}
  
  public static void readSheetString(String path)throws Exception
  {
    
  Workbook wb = createXLSstring();
  ExcelReaderExceptions ExcelExceptions = null;
  // On écrit le contenu dans le fichier de sortie
  try {OutputStream fileOut = new FileOutputStream(path);
 
        {
          wb.write(fileOut);
        } 
  
      } catch (ExcelReaderExceptions e) {
        ExcelExceptions = e;
        throw e;
      }
  finally {
  // Objets permettant de formatter le contenu d'une cellule
  DataFormatter formatter = new DataFormatter();

  // On récupère la feuille courante
  Sheet sheetString = wb.getSheetAt(0);

  // On parcoure les colonnes
  for (Row row1 : sheetString)
  {

    // On parcoure les lignes
    for (Cell cell : row1)
    {

      // On récupère la localisation de la cellule
      CellReference cellRef = new CellReference(row1.getRowNum(), cell.getColumnIndex());

      // On affiche la localisation de la cellule
      System.out.print(cellRef.formatAsString());

      // Pour séparer la localisation de la cellule de son contenu
      System.out.print(" - ");

      // On recupère le contenu brut de la cellule
      String text = formatter.formatCellValue(cell);

      // On affiche ce contenu
      System.out.println(text);

    } // On ferme le second for
  } // On ferme le premier for
  
  System.out.println("\n-----------------------------------\n");

 }
}
  
  public static void main(String[] args) throws Exception
  {
 
    

    System.out.println("\n-----------------------------------\n");
    readSheetNum("C:\\Users\\Taffoureau\\Music\\Excel Tests\\workbook1.xls");
    System.out.println("\n-----------------------------------\n");
    readSheetString("C:\\Users\\Taffoureau\\Music\\Excel Tests\\workbook2.xls");

    
    ExcelReader testNum = new ExcelReader("C:\\Users\\Taffoureau\\Music\\Excel Tests\\workbook1.xls");
    Pullable p = testNum.getPullableOutput();

    for (int i = 0; i < 10; i++)
    {
      double x = (Double) p.pull();

      // On affiche à l'écran
      System.out.println("Le fichier contient: " + x);

    } 
    
    System.out.println("\n-----------------------------------\n");
    
    ExcelReader testString = new ExcelReader("C:\\Users\\Taffoureau\\Music\\Excel Tests\\workbook2.xls");
    Pullable p2 = testString.getPullableOutput();

    for (int i = 0; i < 10; i++)
    {
      String x = (String) p2.pull();

      // On affiche à l'écran
      System.out.println("Le fichier contient: " + x);

    }
   } // On ferme le main
  } // On ferme la classe

