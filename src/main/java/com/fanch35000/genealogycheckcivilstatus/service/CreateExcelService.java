
package com.fanch35000.genealogycheckcivilstatus.service;

import com.fanch35000.genealogycheckcivilstatus.entity.civilstatus.CivilStatus;
import com.fanch35000.genealogycheckcivilstatus.entity.ged.Fam;
import com.fanch35000.genealogycheckcivilstatus.entity.ged.Indi;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

public class CreateExcelService {

    public Set<Integer> listSosa = new HashSet<>();

    private final HSSFWorkbook workbook = new HSSFWorkbook();
    private final HSSFCellStyle styleGeneration = createStyleForGeneration(workbook);
    private final HSSFCellStyle styleTitle = createStyleForTitle(workbook);
    private final HSSFCellStyle styleGrey = createStyleGrey(workbook);
    private final HSSFCellStyle styleOK = createStyleGreen(workbook);
    private final HSSFCellStyle styleWARN = createStyleOrange(workbook);
    private final HSSFCellStyle styleKO = createStyleRed(workbook);

    public final void launch(String xlsLocation, HashMap<Integer, CivilStatus> ancetres, HashMap<String, Indi> indis) {
        //HSSFSheet actesSheet = workbook.createSheet("Actes sheet");
        HSSFSheet gedSheet = workbook.createSheet("GED sheet");

        try {
            //updateActesSheet(actesSheet, ancetres);
            updateGedSheet(workbook, gedSheet, indis, ancetres);

            File file = new File(xlsLocation);
            file.getParentFile().mkdirs();
            FileOutputStream outFile = new FileOutputStream(file);
            workbook.write(outFile);
            System.out.println("Created file: " + file.getAbsolutePath());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static HSSFCellStyle createStyleForGeneration(HSSFWorkbook workbook) {
        HSSFFont font = workbook.createFont();
        font.setItalic(true);
        font.setBold(true);
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.TOP);
        return style;
    }

    private static HSSFCellStyle createStyleForTitle(HSSFWorkbook workbook) {
        HSSFFont font = workbook.createFont();
        font.setBold(true);
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setFillForegroundColor(workbook.getCustomPalette().findSimilarColor(153, 204, 255).getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
    }

    private static HSSFCellStyle createStyleGrey(HSSFWorkbook workbook) {
        HSSFFont font = workbook.createFont();
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    private static HSSFCellStyle createStyleGreen(HSSFWorkbook workbook) {
        HSSFFont font = workbook.createFont();
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setFillForegroundColor(workbook.getCustomPalette().findSimilarColor(153, 204, 0).getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    private static HSSFCellStyle createStyleOrange(HSSFWorkbook workbook) {
        HSSFFont font = workbook.createFont();
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setFillForegroundColor(workbook.getCustomPalette().findSimilarColor(255, 204, 0).getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    private static HSSFCellStyle createStyleRed(HSSFWorkbook workbook) {
        HSSFFont font = workbook.createFont();
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setFillForegroundColor(workbook.getCustomPalette().findSimilarColor(255, 102, 0).getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    public void updateActesSheet(HSSFSheet sheet, HashMap<Integer, CivilStatus> ancetres) throws IOException {

        int rownum = 0;
        Cell cell;
        Row row;

        row = sheet.createRow(rownum);

        // Sosa
        cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("Sosa");
        cell.setCellStyle(styleTitle);
        // Nom
        cell = row.createCell(1, CellType.STRING);
        cell.setCellValue("Nom");
        cell.setCellStyle(styleTitle);
        // Prenom
        cell = row.createCell(2, CellType.STRING);
        cell.setCellValue("Prenom");
        cell.setCellStyle(styleTitle);
        // Naissance
        cell = row.createCell(3, CellType.STRING);
        cell.setCellValue("Naissance");
        cell.setCellStyle(styleTitle);
        // Mariage
        cell = row.createCell(4, CellType.STRING);
        cell.setCellValue("Mariage");
        cell.setCellStyle(styleTitle);
        // Décès
        cell = row.createCell(5, CellType.STRING);
        cell.setCellValue("Décès");
        cell.setCellStyle(styleTitle);

        // Data
        ArrayList<Integer> listeId = new ArrayList<>();
        listeId.addAll(ancetres.keySet());
        Collections.sort(listeId);
        DataFormatter dataFormatter = new DataFormatter();
        dataFormatter.addFormat("yyyy-MM-dd", new java.text.SimpleDateFormat("yyyy-MM-dd"));
        Integer sosaMax = listeId.get(listeId.size() - 1);
        for (int sosa = 1; sosa <= sosaMax; sosa++) {
            CivilStatus civilStatus = ancetres.get(sosa);
            if (civilStatus == null) {
                civilStatus = new CivilStatus();
                civilStatus.setSosa(sosa);
            }

            rownum++;
            row = sheet.createRow(rownum);

            // Sosa
            cell = row.createCell(0, CellType.NUMERIC);
            cell.setCellValue(civilStatus.getSosa());
            // Nom
            cell = row.createCell(1, CellType.STRING);
            cell.setCellValue(civilStatus.getNom());
            // Prénom
            cell = row.createCell(2, CellType.STRING);
            cell.setCellValue(civilStatus.getPrenom());
            // Naissance
            cell = row.createCell(3, CellType.STRING);
            cell.setCellValue(toExcelDate(civilStatus.getDateNaissance()));
            dataFormatter.formatCellValue(cell);
            // Mariage
            cell = row.createCell(4, CellType.STRING);
            cell.setCellValue(toExcelDate(civilStatus.getDateMariage()));
            dataFormatter.formatCellValue(cell);
            // Décès
            cell = row.createCell(5, CellType.STRING);
            cell.setCellValue(toExcelDate(civilStatus.getDateDeces()));
            dataFormatter.formatCellValue(cell);
        }

        int dim = 2048;
        sheet.setColumnWidth(0, dim);
        sheet.setColumnWidth(1, dim * 2);
        sheet.setColumnWidth(2, dim * 4);
        sheet.setColumnWidth(3, dim * 2);
        sheet.setColumnWidth(4, dim * 2);
        sheet.setColumnWidth(5, dim * 2);

    }

    private final static Pattern PATTERN_YEAR = Pattern.compile(".*(\\d{4}).*");

    private final static int getYear(String date) {
        // xxx1923xxx => 1923
        if (date != null) {
            Matcher macher = PATTERN_YEAR.matcher(date);
            if (macher.matches()) {
                return Integer.parseInt(macher.group(1));
            }
            return 0;
        }
        return 0;
    }

    private final static Pattern PATTERN_DATE = Pattern.compile("(\\d+)-(\\d+)-(\\d+)");

    private final static String toExcelDate(String date) {
        // yyyy-MM-dd => dd/MM/yyyy
        if (date != null) {
            Matcher macher = PATTERN_DATE.matcher(date);
            if (macher.matches()) {
                return macher.group(3) + "/" + macher.group(2) + "/" + macher.group(1);
            }
            return date;
        }
        return null;
    }

    private static final int LIMIT_YEAR = 1920;

    public final static Integer getMinSosa(List<Integer> sosas) {
        if (sosas == null || sosas.isEmpty()) {
            return -1;
        } else {
            return sosas.stream().sorted().findFirst().get();
        }
    }

    public void updateGedSheet(HSSFWorkbook workbook, HSSFSheet sheet, HashMap<String, Indi> indis, HashMap<Integer, CivilStatus> ancetres) throws IOException {

        //final int ID_GED = 0;
        int colnum = 0;
        final int GENERATION = colnum++;
        final int SOSA = colnum++;
        final int NOM = colnum++;
        final int PRENOM = colnum++;
        final int DATE_DE_NAISSANCE = colnum++;
        final int LIEU_DE_NAISSANCE = colnum++;
        final int ACTE_DE_NAISSANCE = colnum++;
        final int DATE_DE_MARIAGE = colnum++;
        final int LIEU_DE_MARIAGE = colnum++;
        final int ACTE_DE_MARIAGE = colnum++;
        final int DATE_DE_DECES = colnum++;
        final int LIEU_DE_DECES = colnum++;
        final int ACTE_DE_DECES = colnum++;
        final int NOMBRE_D_ACTES = colnum++;

        int rownum = 0;
        Cell cell;
        Cell firstcell;
        Row row;

        // Entete 1
        row = sheet.createRow(rownum);

        // Naissance
        cell = row.createCell(DATE_DE_NAISSANCE, CellType.STRING);
        cell.setCellValue("Naissance");
        cell.setCellStyle(styleTitle);

        CellRangeAddress regionNaissance = new CellRangeAddress(row.getRowNum(), row.getRowNum(), DATE_DE_NAISSANCE, ACTE_DE_NAISSANCE);
        sheet.addMergedRegion(regionNaissance);
        RegionUtil.setBorderBottom(BorderStyle.THIN, regionNaissance, sheet);
        RegionUtil.setBorderTop(BorderStyle.THIN, regionNaissance, sheet);
        RegionUtil.setBorderLeft(BorderStyle.THIN, regionNaissance, sheet);
        RegionUtil.setBorderRight(BorderStyle.THIN, regionNaissance, sheet);


        // Mariage
        cell = row.createCell(ACTE_DE_MARIAGE, CellType.STRING);
        cell.setCellStyle(styleTitle);
        cell = row.createCell(DATE_DE_MARIAGE, CellType.STRING);
        cell.setCellValue("Mariage");
        cell.setCellStyle(styleTitle);

        CellRangeAddress regionMariage = new CellRangeAddress(row.getRowNum(), row.getRowNum(), DATE_DE_MARIAGE, ACTE_DE_MARIAGE);
        sheet.addMergedRegion(regionMariage);
        RegionUtil.setBorderBottom(BorderStyle.THIN, regionMariage, sheet);
        RegionUtil.setBorderTop(BorderStyle.THIN, regionMariage, sheet);
        RegionUtil.setBorderLeft(BorderStyle.THIN, regionMariage, sheet);
        RegionUtil.setBorderRight(BorderStyle.THIN, regionMariage, sheet);

        // Décès
        cell = row.createCell(DATE_DE_DECES, CellType.STRING);
        cell.setCellValue("Décès");
        cell.setCellStyle(styleTitle);

        CellRangeAddress regionDeces = new CellRangeAddress(row.getRowNum(), row.getRowNum(), DATE_DE_DECES, ACTE_DE_DECES);
        sheet.addMergedRegion(regionDeces);
        RegionUtil.setBorderBottom(BorderStyle.THIN, regionDeces, sheet);
        RegionUtil.setBorderTop(BorderStyle.THIN, regionDeces, sheet);
        RegionUtil.setBorderLeft(BorderStyle.THIN, regionDeces, sheet);
        RegionUtil.setBorderRight(BorderStyle.THIN, regionDeces, sheet);

        // Entete 2
        rownum++;
        row = sheet.createRow(rownum);

        // Generation
        cell = row.createCell(GENERATION, CellType.STRING);
        cell.setCellValue("#Génération");
        cell.setCellStyle(styleTitle);
        firstcell = cell;

        // Sosa
        cell = row.createCell(SOSA, CellType.STRING);
        cell.setCellValue("#Sosa");
        cell.setCellStyle(styleTitle);

        // Nom
        cell = row.createCell(NOM, CellType.STRING);
        cell.setCellValue("Nom");
        cell.setCellStyle(styleTitle);
        // Prenom
        cell = row.createCell(PRENOM, CellType.STRING);
        cell.setCellValue("Prenom");
        cell.setCellStyle(styleTitle);

        // Date Naissance
        cell = row.createCell(DATE_DE_NAISSANCE, CellType.STRING);
        cell.setCellValue("Date");
        cell.setCellStyle(styleTitle);
        // Acte Naissance
        cell = row.createCell(ACTE_DE_NAISSANCE, CellType.STRING);
        cell.setCellValue("Acte");
        cell.setCellStyle(styleTitle);
        // Lieu Naissance
        cell = row.createCell(LIEU_DE_NAISSANCE, CellType.STRING);
        cell.setCellValue("Lieu");
        cell.setCellStyle(styleTitle);

        // Date Mariage
        cell = row.createCell(DATE_DE_MARIAGE, CellType.STRING);
        cell.setCellValue("Date");
        cell.setCellStyle(styleTitle);
        // Lieu de mariage
        cell = row.createCell(LIEU_DE_MARIAGE, CellType.STRING);
        cell.setCellValue("Lieu");
        cell.setCellStyle(styleTitle);
        // Acte de mariage
        cell = row.createCell(ACTE_DE_MARIAGE, CellType.STRING);
        cell.setCellValue("Acte");
        cell.setCellStyle(styleTitle);

        // Date de décès
        cell = row.createCell(DATE_DE_DECES, CellType.STRING);
        cell.setCellValue("Date");
        cell.setCellStyle(styleTitle);
        // Lieu de décès
        cell = row.createCell(LIEU_DE_DECES, CellType.STRING);
        cell.setCellValue("Lieu");
        cell.setCellStyle(styleTitle);
        // Acte de décès
        cell = row.createCell(ACTE_DE_DECES, CellType.STRING);
        cell.setCellValue("Acte");
        cell.setCellStyle(styleTitle);

        // Nombre d'actes
        cell = row.createCell(NOMBRE_D_ACTES, CellType.STRING);
        cell.setCellValue("Nombre d'actes");
        cell.setCellStyle(styleTitle);

        // Data
        ArrayList<String> listeId = new ArrayList<>();
        listeId.addAll(indis.keySet());
        Collections.sort(listeId);
        DataFormatter dataFormatter = new DataFormatter();
        dataFormatter.addFormat("yyyy-MM-dd", new java.text.SimpleDateFormat("yyyy-MM-dd"));

        List<Indi> indisSorted = indis.values().stream().sorted((i1, i2) ->
                getMinSosa(i1.getSosas()).compareTo(getMinSosa(i2.getSosas()))
        ).collect(Collectors.toList());

        Row firstRowRange = null;
        Indi firstIndiRange = null;
        Indi lastIndi = null;
        if (!indisSorted.isEmpty()) {
            lastIndi = indisSorted.get(indisSorted.size() - 1);
        }

        final Map<Integer, Integer> MIN_SOSA = new HashMap();
        final Map<Integer, Indi> INDIS = new HashMap();
        for (Indi indi : indisSorted) {
            Integer min = getMinSosa(indi.getSosas());
            for (Integer sosa : indi.getSosas()) {
                MIN_SOSA.put(sosa, min);
                INDIS.put(sosa, indi);
            }
        }

        List<Integer> sosasSorted = MIN_SOSA.keySet().stream().sorted().collect(Collectors.toList());

        for (Integer sosa : sosasSorted) {

            Indi indi = INDIS.get(sosa);
            Integer minSosa = MIN_SOSA.get(sosa);

            listSosa.add(minSosa);

            rownum++;
            row = sheet.createRow(rownum);

            if (sosa != minSosa) {

                // Id
                cell = row.createCell(SOSA, CellType.STRING);

                cell.setCellValue(sosa);
                cell.setCellComment(getComment(workbook, sheet, row, cell, "Cf. #" + minSosa));
                cell.setCellStyle(styleGrey);

                cell = row.createCell(NOM, CellType.STRING);
                cell.setCellValue("");
                cell.setCellStyle(styleGrey);
                cell = row.createCell(PRENOM, CellType.STRING);
                cell.setCellValue("");
                cell.setCellStyle(styleGrey);
                cell = row.createCell(DATE_DE_NAISSANCE, CellType.STRING);
                cell.setCellValue("");
                cell.setCellStyle(styleGrey);
                cell = row.createCell(LIEU_DE_NAISSANCE, CellType.STRING);
                cell.setCellValue("");
                cell.setCellStyle(styleGrey);
                cell = row.createCell(ACTE_DE_NAISSANCE, CellType.STRING);
                cell.setCellValue("");
                cell.setCellStyle(styleGrey);
                cell = row.createCell(DATE_DE_MARIAGE, CellType.STRING);
                cell.setCellValue("");
                cell.setCellStyle(styleGrey);
                cell = row.createCell(LIEU_DE_MARIAGE, CellType.STRING);
                cell.setCellValue("");
                cell.setCellStyle(styleGrey);
                cell = row.createCell(ACTE_DE_MARIAGE, CellType.STRING);
                cell.setCellValue("");
                cell.setCellStyle(styleGrey);
                cell = row.createCell(DATE_DE_DECES, CellType.STRING);
                cell.setCellValue("");
                cell.setCellStyle(styleGrey);
                cell = row.createCell(LIEU_DE_DECES, CellType.STRING);
                cell.setCellValue("");
                cell.setCellStyle(styleGrey);
                cell = row.createCell(ACTE_DE_DECES, CellType.STRING);
                cell.setCellValue("");
                cell.setCellStyle(styleGrey);
                cell = row.createCell(NOMBRE_D_ACTES, CellType.STRING);
                cell.setCellValue("");
                cell.setCellStyle(styleGrey);

            } else {

                CivilStatus actes = ancetres.get(minSosa);

                // Génération
                cell = row.createCell(GENERATION, CellType.NUMERIC);
                if (indi.getSosas() != null && !indi.getSosas().isEmpty()) {

                    Integer nbFind = 1;

                    if (firstRowRange == null) {
                        firstRowRange = row;
                    }
                    if (firstIndiRange == null) {
                        firstIndiRange = indi;
                    }
                    if (log2(getMinSosa(firstIndiRange.getSosas())) != log2(getMinSosa(indi.getSosas()))) {
                        if (firstRowRange.getRowNum() != row.getRowNum() - 1) {
                            sheet.addMergedRegion(new CellRangeAddress(firstRowRange.getRowNum(), row.getRowNum() - 1, GENERATION, GENERATION));

                            CellRangeAddress regionDataDeces = new CellRangeAddress(firstRowRange.getRowNum(), row.getRowNum() - 1, GENERATION, NOMBRE_D_ACTES);
                            RegionUtil.setBorderTop(BorderStyle.THIN, regionDataDeces, sheet);
                            RegionUtil.setBorderBottom(BorderStyle.THIN, regionDataDeces, sheet);
                            nbFind = row.getRowNum() - firstRowRange.getRowNum();

                            Integer numGeneration = log2(getMinSosa(firstIndiRange.getSosas()));
                            Integer nbMax = (int) Math.pow(2, numGeneration);
                            firstRowRange.getCell(GENERATION).setCellValue(numGeneration + " (" + nbFind + "/" + nbMax + ")");
                            firstRowRange.getCell(GENERATION).setCellStyle(styleGeneration);
                        }
                        firstRowRange = row;
                        firstIndiRange = indi;
                    }
                    if (lastIndi == indi) {
                        sheet.addMergedRegion(new CellRangeAddress(firstRowRange.getRowNum(), row.getRowNum(), GENERATION, GENERATION));

                        CellRangeAddress regionDataDeces = new CellRangeAddress(firstRowRange.getRowNum(), row.getRowNum(), GENERATION, NOMBRE_D_ACTES);
                        RegionUtil.setBorderTop(BorderStyle.THIN, regionDataDeces, sheet);
                        RegionUtil.setBorderBottom(BorderStyle.THIN, regionDataDeces, sheet);
                        nbFind = row.getRowNum() - firstRowRange.getRowNum() + 1;

                        Integer numGeneration = log2(getMinSosa(firstIndiRange.getSosas()));
                        Integer nbMax = (int) Math.pow(2, numGeneration);
                        firstRowRange.getCell(GENERATION).setCellValue(numGeneration + " (" + nbFind + "/" + nbMax + ")");
                        firstRowRange.getCell(GENERATION).setCellStyle(styleGeneration);
                    }
                }

                // Id
                cell = row.createCell(SOSA, CellType.STRING);
                if (minSosa != null) {
                    cell.setCellValue(minSosa);
                }
                // Nom
                cell = row.createCell(NOM, CellType.STRING);
                cell.setCellValue(indi.getNom());
                // Prénom
                cell = row.createCell(PRENOM, CellType.STRING);
                cell.setCellValue(indi.getPrenom());

                // Date de naissance
                cell = row.createCell(DATE_DE_NAISSANCE, CellType.STRING);
                cell.setCellValue(toExcelDate(toExcelDate(indi.getDateNaissance())));
                dataFormatter.formatCellValue(cell);
                // Lieu de naissance
                cell = row.createCell(LIEU_DE_NAISSANCE, CellType.STRING);
                cell.setCellValue(compress(indi.getLieuNaissance()));
                cell.setCellComment(getComment(workbook, sheet, row, cell, indi.getLieuNaissance()));
                dataFormatter.formatCellValue(cell);
                // Acte de naissance
                Cell cellActeNaissance = row.createCell(ACTE_DE_NAISSANCE, CellType.STRING);
                if (actes != null && actes.getDateNaissance() != null) {
                    cellActeNaissance.setCellValue(toExcelDate(toExcelDate(actes.getDateNaissance())));
                    cellActeNaissance.setCellComment(getComment(workbook, sheet, row, cell, actes.getVilleRegistreNaissance()));
                    cellActeNaissance.setCellStyle(styleOK);
                } else {
                    cellActeNaissance.setCellValue("");
                    if (getYear(indi.getDateNaissance()) > LIMIT_YEAR) {
                        cellActeNaissance.setCellStyle(styleGrey);
                    }
                }
                dataFormatter.formatCellValue(cellActeNaissance);

                Fam mariage = null;
                if (!indi.getMariages().isEmpty()) {
                    mariage = indi.getMariages().get(0);
                }
                if (mariage == null) {
                    mariage = new Fam();
                }
                Cell cellActeMariage = null;
                if (sosa % 2 == 0) {
                    // Date Mariage
                    cell = row.createCell(DATE_DE_MARIAGE, CellType.STRING);
                    cell.setCellValue(toExcelDate(toExcelDate(mariage.getDateMariage())));
                    dataFormatter.formatCellValue(cell);
                    // Lieu Mariage
                    cell = row.createCell(LIEU_DE_MARIAGE, CellType.STRING);
                    cell.setCellValue(compress(mariage.getLieuMariage()));
                    cell.setCellComment(getComment(workbook, sheet, row, cell, mariage.getLieuMariage()));
                    dataFormatter.formatCellValue(cell);
                    // Acte Mariage
                    cellActeMariage = row.createCell(ACTE_DE_MARIAGE, CellType.STRING);
                    if (actes != null && actes.getDateMariage() != null) {
                        cellActeMariage.setCellValue(toExcelDate(toExcelDate(actes.getDateMariage())));
                        cellActeMariage.setCellComment(getComment(workbook, sheet, row, cell, actes.getVilleRegistreMariage()));
                        cellActeMariage.setCellStyle(styleOK);
                    } else {
                        cellActeMariage.setCellValue("");
                        if (getYear(indi.getDateNaissance()) > LIMIT_YEAR || getYear(mariage.getDateMariage()) > LIMIT_YEAR) {
                            cellActeMariage.setCellStyle(styleGrey);
                        }
                    }
                } else {
                    // Date Mariage
                    cell = row.createCell(DATE_DE_MARIAGE, CellType.STRING);
                    cell.setCellValue("");
                    cell.setCellStyle(styleGrey);
                    // Lieu Mariage
                    cell = row.createCell(LIEU_DE_MARIAGE, CellType.STRING);
                    cell.setCellValue("");
                    cell.setCellStyle(styleGrey);
                    // Acte Mariage
                    cellActeMariage = row.createCell(ACTE_DE_MARIAGE, CellType.STRING);
                    cellActeMariage.setCellValue("");
                    cellActeMariage.setCellStyle(styleGrey);
                }
                dataFormatter.formatCellValue(cellActeMariage);

                // Date Décès
                cell = row.createCell(DATE_DE_DECES, CellType.STRING);
                cell.setCellValue(toExcelDate(toExcelDate(indi.getDateDeces())));
                dataFormatter.formatCellValue(cell);
                // Lieu Décès
                cell = row.createCell(LIEU_DE_DECES, CellType.STRING);
                cell.setCellValue(compress(indi.getLieuDeces()));
                cell.setCellComment(getComment(workbook, sheet, row, cell, indi.getLieuDeces()));
                dataFormatter.formatCellValue(cell);
                // Acte Décès
                Cell cellActeDeces = row.createCell(ACTE_DE_DECES, CellType.STRING);
                if (actes != null && actes.getDateDeces() != null) {
                    cellActeDeces.setCellValue(toExcelDate(toExcelDate(actes.getDateDeces())));
                    cellActeDeces.setCellComment(getComment(workbook, sheet, row, cell, actes.getVilleRegistreDeces()));
                    cellActeDeces.setCellStyle(styleOK);
                } else {
                    cellActeDeces.setCellValue("");
                    if (getYear(indi.getDateNaissance()) > LIMIT_YEAR || getYear(mariage.getDateMariage()) > LIMIT_YEAR || getYear(indi.getDateDeces()) > LIMIT_YEAR) {
                        cellActeDeces.setCellStyle(styleGrey);
                    }
                }
                dataFormatter.formatCellValue(cellActeDeces);

                // Nombre d'actes
                cell = row.createCell(NOMBRE_D_ACTES, CellType.FORMULA);
                cell.setCellFormula("IF(ISBLANK(" + cellActeNaissance.getAddress() + "),\"0\",\"1\")+IF(ISBLANK(" + cellActeMariage.getAddress() + "),\"0\",\"1\")+IF(ISBLANK(" + cellActeDeces.getAddress() + "),\"0\",\"1\")");
                dataFormatter.formatCellValue(cell);
            }
        }

        int dim = 2048;
        //sheet.setColumnWidth(ID_GED, dim);
        sheet.setColumnWidth(GENERATION, dim * 2);
        sheet.setColumnWidth(SOSA, dim * 2);
        sheet.setColumnWidth(NOM, dim * 2);
        sheet.setColumnWidth(PRENOM, dim * 4);
        sheet.setColumnWidth(DATE_DE_NAISSANCE, dim * 2);
        sheet.setColumnWidth(LIEU_DE_NAISSANCE, dim * 3);
        sheet.setColumnWidth(ACTE_DE_NAISSANCE, dim * 2);
        sheet.setColumnWidth(DATE_DE_MARIAGE, dim * 2);
        sheet.setColumnWidth(LIEU_DE_MARIAGE, dim * 3);
        sheet.setColumnWidth(ACTE_DE_MARIAGE, dim * 2);
        sheet.setColumnWidth(DATE_DE_DECES, dim * 2);
        sheet.setColumnWidth(LIEU_DE_DECES, dim * 3);
        sheet.setColumnWidth(ACTE_DE_DECES, dim * 2);
        sheet.setColumnWidth(NOMBRE_D_ACTES, dim * 2);

        sheet.setAutoFilter(new CellRangeAddress(firstcell.getRowIndex(), cell.getRowIndex(), firstcell.getColumnIndex(), cell.getColumnIndex()));
        sheet.createFreezePane(0, 2);


        CellRangeAddress regionDataInfo = new CellRangeAddress(firstcell.getRowIndex(), cell.getRowIndex(), GENERATION, PRENOM);
        RegionUtil.setBorderBottom(BorderStyle.THIN, regionDataInfo, sheet);
        RegionUtil.setBorderTop(BorderStyle.THIN, regionDataInfo, sheet);
        RegionUtil.setBorderLeft(BorderStyle.THIN, regionDataInfo, sheet);
        RegionUtil.setBorderRight(BorderStyle.DOUBLE, regionDataInfo, sheet);

        CellRangeAddress regionDataNaissance = new CellRangeAddress(firstcell.getRowIndex(), cell.getRowIndex(), DATE_DE_NAISSANCE, ACTE_DE_NAISSANCE);
        RegionUtil.setBorderBottom(BorderStyle.THIN, regionDataNaissance, sheet);
        RegionUtil.setBorderTop(BorderStyle.THIN, regionDataNaissance, sheet);
        RegionUtil.setBorderLeft(BorderStyle.DOUBLE, regionDataNaissance, sheet);
        RegionUtil.setBorderRight(BorderStyle.DOUBLE, regionDataNaissance, sheet);

        CellRangeAddress regionDataMariage = new CellRangeAddress(firstcell.getRowIndex(), cell.getRowIndex(), DATE_DE_MARIAGE, ACTE_DE_MARIAGE);
        RegionUtil.setBorderBottom(BorderStyle.THIN, regionDataMariage, sheet);
        RegionUtil.setBorderTop(BorderStyle.THIN, regionDataMariage, sheet);
        RegionUtil.setBorderLeft(BorderStyle.DOUBLE, regionDataMariage, sheet);
        RegionUtil.setBorderRight(BorderStyle.DOUBLE, regionDataMariage, sheet);

        CellRangeAddress regionDataDeces = new CellRangeAddress(firstcell.getRowIndex(), cell.getRowIndex(), DATE_DE_DECES, ACTE_DE_DECES);
        RegionUtil.setBorderBottom(BorderStyle.THIN, regionDataDeces, sheet);
        RegionUtil.setBorderTop(BorderStyle.THIN, regionDataDeces, sheet);
        RegionUtil.setBorderLeft(BorderStyle.DOUBLE, regionDataDeces, sheet);
        RegionUtil.setBorderRight(BorderStyle.DOUBLE, regionDataDeces, sheet);

        CellRangeAddress regionDataTotal = new CellRangeAddress(firstcell.getRowIndex(), cell.getRowIndex(), NOMBRE_D_ACTES, NOMBRE_D_ACTES);
        RegionUtil.setBorderBottom(BorderStyle.THIN, regionDataTotal, sheet);
        RegionUtil.setBorderTop(BorderStyle.THIN, regionDataTotal, sheet);
        RegionUtil.setBorderLeft(BorderStyle.DOUBLE, regionDataTotal, sheet);
        RegionUtil.setBorderRight(BorderStyle.THIN, regionDataTotal, sheet);

    }

    public static int log2(int x) {
        return (int) (Math.log(x) / Math.log(2));
    }

    public static String compress(String place) {
        if (place != null) {
            if (place.contains("] - ")) {
                place = place.substring(place.indexOf("] - ") + 4);
            } else if (place.contains(", ")) {
                place = place.substring(0, place.indexOf(", "));
            } else if (place.endsWith(",")) {
                place = place.substring(0, place.indexOf(","));
            } else if (place.contains(",")) {
                place = place.substring(0,place.indexOf(",")) + " [" + place.substring(place.lastIndexOf(",")+1,place.length()) + "]";
            }
        }
        return place;
    }

    public static Comment getComment(HSSFWorkbook workbook, HSSFSheet sheet, Row row, Cell cell, String commentaire) {
        if (commentaire != null) {
            CreationHelper factory = workbook.getCreationHelper();
            Drawing drawing = sheet.createDrawingPatriarch();
            // When the comment box is visible, have it show in a 1x3 space
            ClientAnchor anchor = factory.createClientAnchor();
            anchor.setCol1(cell.getColumnIndex());
            anchor.setCol2(cell.getColumnIndex() + 1);
            anchor.setRow1(row.getRowNum());
            anchor.setRow2(row.getRowNum() + 6);
            Comment comment = drawing.createCellComment(anchor);
            String commentaire2 = commentaire;
            if(commentaire2.contains(", ")) {
                // Format 1
                commentaire2 = commentaire2.replaceAll("] - ", "]\n");
                commentaire2 = commentaire2.replaceAll(", ", "\n");
            } else {
                // Format 2
                commentaire2 = commentaire2.replaceAll(",", "\n");
            }
            RichTextString str = factory.createRichTextString(commentaire2);
            comment.setString(str);
            comment.setAuthor("Apache POI");
            return comment;
        }
        return null;
    }

}
