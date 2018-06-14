# Excel palette for BeepBeep 3

This repository contain a palette (extension) of [BeepBeep3](https://liflab.github.io/beepbeep-3/) for the manipulation of Excel data as events streams. All projects require `beepbeep-3.jar` in their classpath (or alternately, must point to the Core source files from BeepBeep's repository) in order to compile and run.

The palette now contain the following files :

- `ExcelReader.java`: It's the main file of the palette, containing constructors and methods for the extraction of data.

- `ExcelReaderExample.java`: Create an Excel file with numbers in, and, in the main method, call an ExcelReader processor to extract and show the data of this file.

- `BaseTestExcelReader.java`: Functions to execute before and after tests.

- `TestExcelReader.java`: Some JUnit tests of ExcelReader.java. (Not finished yet)

- `ExcelReaderExceptions.java`: Exceptions.

## How to build the palette ?

Follow the instructions given at [this page](https://github.com/liflab/beepbeep-3-palettes)
You also need the Apache POI librairy .jar, you can download it at this page

## How it's work ?

Basically, if you want to extract the data of an Excel file, you need (after import the correct packages), create an new processor ExcelReader, who have two differents constructors :

The first constructor take in input only the name of the Excel file, this will extract all of the files, row by row :
```
 ExcelReader test = new ExcelReader("absolutPathOfYourFile");
```
The second constructor take in input the name of the file, and one or more columns you want to extract :
```
ExcelReader test = new ExcelReader("absolutPathOfYourFile", 4);
```
or
```
ExcelReader test = new ExcelReader("absolutPathOfYourFile", 4, 5, 7, 54);
```

You can now connect the output of this processor with others, like `doubler` for example.
Here's an example of the extraction of the third column, containing 10 rows of numbers :

```
ExcelReader Exceltest = new ExcelReader("absolutPathOfYourFile", 2);
Doubler doubler = new Doubler();
Connector.connect(Exceltest, doubler);
Pullable p = doubler.getPullableOutput();

    for (int i = 0; i < 10; i++)
    {
      int x = (Integer) p.pull();
      System.out.println("Le fichier contient: " + x);
    }
```



About the authors                                                  {#about}
-----------------

This palette was written by Nicolas Taffoureau.
