package ca.uqac.lif.cep.excelReader;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Queue;

import basic.PipingUnary.Doubler;
import ca.uqac.lif.cep.*;
import ca.uqac.lif.cep.tmf.Source;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Permet de récupérer le contenu d'un fichier Excel pour faire des tests sur
 * les valeurs Contenues dans les cellules Ce processeur prend en entrée le nom
 * d'un fichier Excel (.xls) Et renvoi en sortie le contenu (linéaire) de
 * celui-ci, ligne par ligne
 * 
 * @author Nicolas Taffoureau
 */

public class ExcelReader extends Source
{
  String m_events;

  public ExcelReader()
  {
    super(1);
  }

  public ExcelReader addEvent(String e)
  {

    if (!e.endsWith("xls"))
    {
      System.out.println("Le fichier n'est pas valide");
    }
    else
    {
      m_events = e;
    }

    return this;

  }

  @Override
  @SuppressWarnings("resource")

  public boolean compute(Object[] inputs, Queue<Object[]> outputs)
  {

    // On récupère le nom du fichier à lire dans un InputStream
    InputStream nomFichier;

    try
    {

      // On crée un FileInputStream correspondant au nom de la feuille rentré en
      // parametre
      nomFichier = new FileInputStream(m_events);

      // On définie un nouveau Workbook
      HSSFWorkbook wb;

      // Le workbook est associé à la feuille passée en parametre
      wb = new HSSFWorkbook(nomFichier);

      // On récupère la feuille courante
      Sheet sheet1 = wb.getSheetAt(0);

      // Pour parcourir l'ArrayList
      int i = 0;

      // Pour stocker le contenu de la feuille
      ArrayList<Object> contenuFeuille = new ArrayList<Object>();

      // On parcoure les colonness
      for (Row row1 : sheet1)
      {

        // On parcoure les lignes
        for (Cell cell : row1)
        {

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

          // On ajoute le contenu de l'ArrayList courante à l'output
          outputs.add(new Object[] { contenuFeuille.get(i) });

          // On parcoure l'ArrayList
          i++;

        } // On ferme le second for
      } // On ferme le premier for

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
    return new ExcelReader();
  }

  public static void main(String[] args) throws Exception
  {

    ExcelReader test = new ExcelReader();
    test.addEvent("C:\\Users\\Taffoureau\\Music\\Excel Tests\\workbook.xls");

    Doubler doubler = new Doubler();
    Connector.connect(test, doubler);
    Pullable p = doubler.getPullableOutput();

    for (int i = 0; i < 51; i++)
    {
      int x = (Integer) p.pull();

      // On affiche à l'écran
      System.out.println("Le fichier contient: " + x);

    } // On ferme le for

  }// On ferme le main

}// On ferme la classe
