package cz.cvut.fit.svitaluc.exceltomd.translator;

import org.apache.commons.io.output.ByteArrayOutputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Comment;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.PrintWriter;
import java.nio.charset.StandardCharsets;
import java.util.logging.Level;
import java.util.logging.Logger;

public class ExcelToMdTranslator {

    private static final Logger LOGGER = Logger.getLogger(ExcelToMdTranslator.class.getName());

    public byte[] translate(Workbook workbook) {
        ByteArrayOutputStream byteArray = new ByteArrayOutputStream();
        PrintWriter writer = new PrintWriter(new OutputStreamWriter(byteArray, StandardCharsets.UTF_8));

        Sheet sheet = workbook.getSheetAt(0);
        for (Row row : sheet) {
            StringBuilder mdLine = new StringBuilder();

            mdLine.append("| ");
            for (Cell cell : row) {
                String value = switch (cell.getCellType()) {
                    case STRING -> cell.getStringCellValue();
                    case NUMERIC -> String.valueOf(cell.getNumericCellValue());
                    case BLANK -> "";
                    default -> "UNEXPECTED VALUE";
                };
                Comment comment = cell.getCellComment();

                if (comment != null) {
                    mdLine.append("[")
                            .append(value)
                            .append("](")
                            .append(getTextFromComment(comment.getString().getString()))
                            .append(")");
                } else {
                    mdLine.append(value);
                }

                mdLine.append(" | ");
            }
            mdLine.append("\n");
            writer.write(mdLine.toString());

            if (row.getRowNum() == 0) {
                writeHeaderBorder(row, writer);
            }
        }

        writer.flush();
        return byteArray.toByteArray();
    }

    private String getTextFromComment(String fullComment) {
        return fullComment.replaceFirst("\\n\\t-[A-zÀ-ú]*\\s[A-zÀ-ú]*", "");
    }

    private void writeHeaderBorder(Row row, PrintWriter writer) {
        StringBuilder lineBuilder = new StringBuilder();
        lineBuilder.append("| ");
        lineBuilder.append(" --- |".repeat(Math.max(0, row.getLastCellNum())));
        lineBuilder.append("\n");
        writer.write(lineBuilder.toString());
    }

    public void saveToFile(byte[] byteArray, String outputDirectory, String originalFileName) {
        String fileName = originalFileName.replace(".xlsx", ".md");
        String outputName = outputDirectory + "\\" + fileName;
        try (FileOutputStream fos = new FileOutputStream(outputName)) {
            fos.write(byteArray);
        } catch (IOException e) {
            LOGGER.log(Level.SEVERE, "Cannot create output file {0}.", outputName);
        }
        LOGGER.log(Level.INFO, "Created a translated file: {0}", outputName);
    }
}
