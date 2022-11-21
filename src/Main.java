import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;

import javax.swing.*;
import java.awt.*;
import java.io.*;
import java.util.ArrayList;
import java.util.Random;

public class Main extends Component {
    private final ArrayList<String> classes = new ArrayList<>();
    private final ArrayList<String> firstNames = new ArrayList<>();
    private final ArrayList<String> lastNames = new ArrayList<>();
    private final ArrayList<Record> records = new ArrayList<>();
    private String sourcePath;
    private String savingPath;
    private HSSFSheet sheet;

    public static void main(String[] args) {
        Main program = new Main();
        program.Core();
    }

    private void Core() {
        System.out.println("Укажите папку с ресурсами для генерации");
        sourcePath = openFolderDialog();
        if (sourcePath == null) {
            System.out.println("Папка с ресурсами не была выбрана");
            System.exit(0);
        }
        System.out.println("Укажите папку для сохранения сгенерированного файла");
        savingPath = openFolderDialog();
        if (savingPath == null) {
            System.out.println("Папка сохранения не была выбрана");
            System.exit(0);
        }
        readSources();
        startGenerate();
        writeToExcel();
    }

    private void readSources() {
        File firstNamesFile = new File(sourcePath + "\\firstNames.txt");
        File lastNamesFile = new File(sourcePath + "\\lastNames.txt");
        File classesFile = new File(sourcePath + "\\classes.txt");
        try (FileReader firstNamesFileReader = new FileReader(firstNamesFile);
             FileReader lastNamesFileReader = new FileReader(lastNamesFile);
             FileReader classesFileReader = new FileReader(classesFile);
             BufferedReader firstNamesBufferedReader = new BufferedReader(firstNamesFileReader);
             BufferedReader lastNamesBufferedReader = new BufferedReader(lastNamesFileReader);
             BufferedReader classesBufferedReader = new BufferedReader(classesFileReader)) {
            String currentLine = firstNamesBufferedReader.readLine();
            while (currentLine != null) {
                firstNames.add(currentLine);
                currentLine = firstNamesBufferedReader.readLine();
            }
            currentLine = lastNamesBufferedReader.readLine();
            while (currentLine != null) {
                lastNames.add(currentLine);
                currentLine = lastNamesBufferedReader.readLine();
            }
            currentLine = classesBufferedReader.readLine();
            while (currentLine != null) {
                classes.add(currentLine);
                currentLine = classesBufferedReader.readLine();
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private void startGenerate() {
        int rowsCount = 50;
        Random random = new Random();
        Record record;
        for (int i = 0; i < rowsCount; i++) {
            record = new Record();
            record.addValue(random.nextInt(10000000 - 6000000) + 6000000);
            record.addValue(classes.get(random.nextInt(classes.size())));
            record.addValue(firstNames.get(random.nextInt(firstNames.size())) + " " + lastNames.get(
                    random.nextInt(lastNames.size())));
            record.addValue(30);
            record.addValue(random.nextInt(10) + 20);
            record.addValue(0);
            record.addValue(1100);
            record.addValue(random.nextInt(200) + 100);
            record.addValue(random.nextInt(200) + 800);
            records.add(record);
        }
        for (Record r : records) {
            System.out.println(r);
        }
    }

    private void writeToExcel() {
        HSSFWorkbook workbook = new HSSFWorkbook();
        sheet = workbook.createSheet("Отчет");
        sheet.createRow(0).createCell(7).setCellValue("УТВЕРЖДАЮ:");
        sheet.createRow(1).createCell(7).setCellValue("Директор:");
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 8, 9));
        sheet.addMergedRegion(new CellRangeAddress(3, 3, 8, 9));
        sheet.addMergedRegion(new CellRangeAddress(4, 4, 8, 9));
        sheet.addMergedRegion(new CellRangeAddress(9, 9, 0, 9));
        sheet.addMergedRegion(new CellRangeAddress(10, 10, 0, 9));
        sheet.addMergedRegion(new CellRangeAddress(12, 12, 0, 9));
        sheet.addMergedRegion(new CellRangeAddress(14, 15, 0, 0));
        sheet.addMergedRegion(new CellRangeAddress(14, 15, 1, 1));
        sheet.addMergedRegion(new CellRangeAddress(14, 15, 2, 2));
        sheet.addMergedRegion(new CellRangeAddress(14, 15, 3, 3));
        sheet.addMergedRegion(new CellRangeAddress(14, 14, 4, 5));
        sheet.addMergedRegion(new CellRangeAddress(14, 15, 6, 6));
        sheet.addMergedRegion(new CellRangeAddress(14, 15, 7, 7));
        sheet.addMergedRegion(new CellRangeAddress(14, 15, 8, 8));
        sheet.addMergedRegion(new CellRangeAddress(14, 15, 9, 9));
        setDataToCell(2, 8, "(сокращенное наименование образовательного учреждения)");
        setDataToCell(3, 7, "_____________");
        setDataToCell(3, 8, "___________________________");
        setDataToCell(4, 7, "(подпись)");
        setDataToCell(4, 8, "(расшифровка подписи)");
        setDataToCell(6, 7, "14.05.2022");
        setDataToCell(7, 7, "М.П.");
        setDataToCell(9, 0, "Отчёт о фактическом предоставленном бесплатном питании");
        setDataToCell(10, 0, "за период с 01.05.2022 по 31.05.2022");
        setDataToCell(12, 0, "(сокращенное наименование образовательного учреждения)");
        setDataToCell(14, 0, "№ п/п");
        setDataToCell(14, 1, "№ счета");
        setDataToCell(14, 2, "Ф.И. ребенка");
        setDataToCell(14, 3, "Класс");
        setDataToCell(14, 4, "Дни посещения");
        setDataToCell(15, 4, "плановые");
        setDataToCell(15, 5, "фактические");
        setDataToCell(14, 6, "Остаток на начало месяца, руб.");
        setDataToCell(14, 7, "Поступило в текущем месяце на питание, руб.");
        setDataToCell(14, 8, "Остаток на конец месяца, руб.");
        setDataToCell(14, 9, "Израсходовано в текущем месяце на питание, руб.");
        int offset = 16;
        int iterator = 0;
        for (Record record : records) {
            setDataToCell(offset, 0, String.valueOf(iterator));
            for (int i = 1; i < 10; i++) {
                setDataToCell(offset, i, record.getStringValue(i - 1));
            }
            offset++;
            iterator++;
        }
        setDataToCell(offset + 3, 1, "Отчет составлен в двух экземплярах.");
        setDataToCell(offset + 5, 1, "Подписи сторон:");
        setDataToCell(offset + 7, 1, "Лицо, ответственное за организацию питания");
        sheet.addMergedRegion(new CellRangeAddress(offset + 7, offset + 7, 4, 5));
        setDataToCell(offset + 7, 4, "_____________");
        sheet.addMergedRegion(new CellRangeAddress(offset + 8, offset + 8, 4, 5));
        setDataToCell(offset + 8, 4, "(подпись)");
        setDataToCell(offset + 9, 1, "Заведующий производством предприятия общественного питания");
        sheet.addMergedRegion(new CellRangeAddress(offset + 9, offset + 9, 4, 5));
        setDataToCell(offset + 9, 4, "_____________");
        sheet.addMergedRegion(new CellRangeAddress(offset + 10, offset + 10, 4, 5));
        setDataToCell(offset + 10, 4, "(подпись)");
        for (int i = 0; i < 20; i++) {
            sheet.autoSizeColumn(i);
        }
        File file = new File(savingPath + "/test.xls");
        try {
            workbook.write(new FileOutputStream(file));
            workbook.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private void setDataToCell(int row, int column, String value) {
        Row currentRow = sheet.getRow(row);
        if (currentRow == null) {
            currentRow = sheet.createRow(row);
        }
        currentRow.createCell(column).setCellValue(value);
    }

    private String openFolderDialog() {
        final JFileChooser fc = new JFileChooser();
        fc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        if (fc.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            return String.valueOf(fc.getSelectedFile());
        }
        return null;
    }
}