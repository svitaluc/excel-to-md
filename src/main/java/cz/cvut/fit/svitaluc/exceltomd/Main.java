package cz.cvut.fit.svitaluc.exceltomd;

import cz.cvut.fit.svitaluc.exceltomd.translator.ExcelToMdTranslator;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;

public class Main {
    private static final Logger LOGGER = Logger.getLogger(Main.class.getName());
    public static void main(String[] args) {
        if (args.length < 2) {
            LOGGER.log(Level.SEVERE, "Provide an argument with an excel file followed by an argument with " +
                    "an output directory.");
            return;
        }

        String inputFileLocation = args[0];
        String outputDirectory = args[1];

        Workbook workbook;
        File inputFile;
        try {
            inputFile = new File(inputFileLocation);
            FileInputStream file = new FileInputStream(inputFile);
            workbook = new XSSFWorkbook(file);
        } catch (IOException e) {
            LOGGER.log(Level.SEVERE, "The provided path to the excel file {0} is faulty.", inputFileLocation);
            return;
        }

        ExcelToMdTranslator translator = new ExcelToMdTranslator();
        byte[] bytesToWrite = translator.translate(workbook);
        translator.saveToFile(bytesToWrite, outputDirectory, inputFile.getName());
    }
}