package ca.uqac.lif.cep.excelReader;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Queue;

import ca.uqac.lif.cep.Processor;
import ca.uqac.lif.cep.tmf.Source;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Permet de récupérer le contenu d'un fichier Excel pour faire des tests sur
 * les valeurs contenues dans les cellules. Ce processeur prend en entrée le nom
 * d'un fichier Excel (.xls) Et renvoi en sortie le contenu (linéaire) de
 * celui-ci, ligne par ligne
 * 
 * @author Nicolas Taffoureau
 */

public class ExcelReader extends Source
{
  String m_file;
  int m_column = -1;
  ArrayList<Integer> m_columntab = new ArrayList<Integer>();

  
  //Constructeur de base
  public ExcelReader(String path) throws ExcelReaderExceptions 
  {
    super(1);
    m_file = path;
    
  //On verifie que le format du fichier est correct
    if (!path.endsWith("xls"))
    {
      throw new ExcelReaderExceptions("Format de fichier incorrect !");
    }
  }

  //Possibilité de retourner juste la ou les colonnes données en paramètre
  public ExcelReader(String path, int... colonnes) throws ExcelReaderExceptions 
  {
    super(1) ;
    m_file = path;
    m_column = 0 ;
    
    //On verifie que le format du fichier est correct
    if (!path.endsWith("xls"))
    {
      throw new ExcelReaderExceptions("Format de fichier incorrect !");
    }
    
    //On verifie que les numéros de colonnes sont valides
    for (int i = 0; i < colonnes.length; i++) 
    {
      m_columntab.add(colonnes[i]);
      if(colonnes[i] < 0) 
      {
        throw new ExcelReaderExceptions("Numéro de colonne invalide !");
      }
    }
  }
 

  
  /**
   * Ajoute les valeurs voulant être recupérées dans une arraylist
   *
   * @param : cell : cellule voulant être recupérée
   * @param : contenuFeuille : Arraylist où va être stockée la valeur
   * @return : void
   */
  public void ajoutValeur (Cell cell, ArrayList<Object> contenuFeuille ) {
    
    // Tests sur le type de texte que contient une cellule
    switch (cell.getCellTypeEnum())
    {

      // Si c'est une chaîne de carectères
      case STRING:

        // On affiche son contenu
        contenuFeuille.add(cell.getRichStringCellValue().getString());
        break;

        // Si c'est un nombre
      case NUMERIC:

        // plus précisément si c'est une date
        if (DateUtil.isCellDateFormatted(cell))
        {
          // On affiche son contenu
          contenuFeuille.add(cell.getDateCellValue());
        }

        else
        {
          // On affiche son contenu
          contenuFeuille.add(cell.getNumericCellValue());
        }
        break;

        // Si c'est un booleen
      case BOOLEAN:
        // On affiche son contenu
        contenuFeuille.add(cell.getBooleanCellValue());
        break;

        // Si c'est une formule
      case FORMULA:
        // On affiche son contenu
        contenuFeuille.add(cell.getCellFormula());
        break;

        // Si la case est vide
      case BLANK:
        // On affiche rien
        contenuFeuille.add("");
        break;

        // Par défaut on suppose que la case est vide
      default:
        // Donc on affiche rien
        contenuFeuille.add("");
     }// On ferme le switch
    
  }
  

  @Override

  public boolean compute(Object[] inputs, Queue<Object[]> outputs)
  {

    // On récupère le nom du fichier à lire dans un InputStream
    InputStream nomFichier;

    try
    {

      // On crée un FileInputStream correspondant au nom de la feuille rentré en
      // parametre
      nomFichier = new FileInputStream(m_file);
      
      //On crée un objet feuille
      Sheet sheet1 ;
       
      if (m_file.endsWith("xls"))
      {
        // On définie un nouveau Workbook de type HSSF
        HSSFWorkbook wb;

        // Le workbook est associé au fichier passé en parametre
        wb = new HSSFWorkbook(nomFichier);
        
        // On récupère la feuille courante
        sheet1 = wb.getSheetAt(0);
      }
      
      else
      {
        // On définie un nouveau Workbook de type XSSF
        XSSFWorkbook wb;

        // Le workbook est associé à la feuille passée en parametre
        wb = new XSSFWorkbook(nomFichier);
        
        // On récupère la feuille courante
        sheet1 = wb.getSheetAt(0);
      }

      // Pour parcourir l'ArrayList
      int i = 0;

      // Pour stocker le contenu de la feuille
      ArrayList<Object> contenuFeuille = new ArrayList<Object>();
      
      //Si aucun numéro de colonne n'est passé en paramètre
      if(m_column == -1)
      {
        
        // On parcoure les colonnes
        for (Row row1 : sheet1)
        {
  
          // On parcoure les lignes
          for (Cell cell : row1)
          {
  
            //On ajoute les valeurs dans une ArrayList
            ajoutValeur(cell, contenuFeuille);
  
            // On ajoute le contenu de l'ArrayList courante à l'output
            outputs.add(new Object[] { contenuFeuille.get(i) });
  
            // On parcoure l'ArrayList
            i++;
  
          } // On ferme le second for
        } // On ferme le premier for
      }
      else if (m_columntab != null)
      {
        for (Integer thiscolumn : m_columntab) {
          
        
        for(Row r : sheet1) 
        {
          //On récupère les valeurs de la colonne passée en paramètre
          Cell cell = r.getCell(thiscolumn);
            if(cell != null) 
            {
              //On ajoute les valeurs dans une ArrayList
              ajoutValeur(cell, contenuFeuille);

            // On ajoute le contenu de l'ArrayList courante à l'output
            outputs.add(new Object[] { contenuFeuille.get(i) });

            // On parcoure l'ArrayList
            i++;
            
            }
        }
        }
      }
      
    }
    // Si le fichier n'existe pas
    catch (FileNotFoundException e)
    {
      e.printStackTrace();
    }
    // Si une erreur de lecture survient
    catch (IOException e)
    {
      e.printStackTrace();
    }
    // Le processeur a terminé sa tâche
    return true;
  }

  @Override
  public Processor duplicate(boolean with_state)
  {
   
      return new ExcelReader(m_file);
    
  }

}// On ferme la classe
