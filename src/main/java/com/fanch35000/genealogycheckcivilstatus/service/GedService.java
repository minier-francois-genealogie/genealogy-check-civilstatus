package com.fanch35000.genealogycheckcivilstatus.service;

import com.fanch35000.genealogycheckcivilstatus.entity.ged.Fam;
import com.fanch35000.genealogycheckcivilstatus.entity.ged.Indi;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Optional;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class GedService {

    public final HashMap<String, Indi> launch(String gedLocation) {
        File file = getGed(gedLocation);
        analyseFile(file);
        evaluateSosa();

        return INDIS;
    }

    private final static File getGed(String gedLocation) {
        return new File(gedLocation);
    }

    private final void evaluateSosa() {
        // find root
        Optional<Indi> first = INDIS.values().stream().findAny();
        if (first.isPresent()) {
            Indi root = findRoot(first.get());
            evaluateSosa(root, 1);
        }
    }

    private void evaluateSosa(Indi indi, int sosa) {
        indi.getSosas().add(sosa);
        if (indi.getPere() != null) {
            evaluateSosa(indi.getPere(), sosa * 2);
        }
        if (indi.getMere() != null) {
            evaluateSosa(indi.getMere(), sosa * 2 + 1);
        }
    }

    private Indi findRoot(Indi indi) {
        if (!indi.getMariages().isEmpty()) {
            Fam fam = indi.getMariages().get(0);
            if (!fam.getEnfants().isEmpty()) {
                Indi indiChild = fam.getEnfants().get(0);
                return findRoot(indiChild);
            }
        }
        return indi;
    }

    protected final String BLOCK_PREFIX = "0 ";
    protected final static Pattern PATTERN_INDI = Pattern.compile("0 @(.+)@\\sINDI");

    private final HashMap<String, Indi> INDIS = new HashMap();

    private final Indi getIndi(String idGed) {
        Indi indi = INDIS.get(idGed);
        if(indi == null){
            indi = new Indi();
            indi.setIdGed(idGed);
            INDIS.put(idGed, indi);
        }
        return indi;
    }

    private final void analyseFile(File file) {
        try {
            Scanner scanner = new Scanner(file);

            //renvoie true s'il y a une autre ligne à lire
            Object currentRecord = null;
            String event = null;
            while (scanner.hasNextLine()) {
                String line = scanner.nextLine();
                if (line.startsWith(BLOCK_PREFIX)) {
                    if (line.contains("INDI")) {
                        Indi currentIndi = null;
                        Matcher macher = PATTERN_INDI.matcher(line);
                        if (macher.matches()) {
                            currentIndi = getIndi(macher.group(1));
                        }
                        currentRecord = currentIndi;
                    } else if (line.contains("FAM")) {
                        currentRecord = new Fam();
                    } else {
                        currentRecord = null;
                    }
                } else {
                    if (line.length() >= 6) {
                        String prefix = line.substring(0, 6);
                        if (currentRecord instanceof Indi) {
                            Indi currentIndi = (Indi) currentRecord;
                            switch (prefix) {
                                // Indi
                                case "1 NAME":
                                    analyseNAME(currentIndi, line);
                                    break;
                                case "1 SEX ":
                                    analyseSEX(currentIndi, line);
                                    break;
                                case "1 OCCU":
                                    analyseOCCU(currentIndi, line);
                                    break;
                                case "1 BIRT":
                                    event = "BIRT";
                                    break;
                                case "1 DEAT":
                                    event = "DEAT";
                                    break;
                                case "2 DATE":
                                    analyseDATE(event, currentIndi, line);
                                    break;
                                case "2 PLAC":
                                    analysePLAC(event, currentIndi, line);
                                    break;
                                case "1 FAMC":
                                    break;
                                case "1 FAMS":
                                    break;
                            }
                        } else if (currentRecord instanceof Fam) {
                            Fam currentFam = (Fam) currentRecord;
                            switch (prefix) {
                                // Family
                                case "1 HUSB":
                                    analyseHUSB(currentFam, line);
                                    break;
                                case "1 WIFE":
                                    analyseWIFE(currentFam, line);
                                    break;
                                case "1 CHIL":
                                    analyseCHIL(currentFam, line);
                                    break;
                                // Common
                                case "1 SOUR":
                                    break;
                                case "1 MARR":
                                    event = "MARR";
                                    break;
                                case "2 DATE":
                                    analyseDATE(event, currentFam, line);
                                    break;
                                case "2 PLAC":
                                    analysePLAC(event, currentFam, line);
                                    break;
                            }
                        }
                    }
                }
            }
            scanner.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public final static Pattern PATTERN_HUSB = Pattern.compile("1 HUSB @(.+)@");

    private final void analyseHUSB(Fam currentRecord, String lineFam) {
        Matcher macher = PATTERN_HUSB.matcher(lineFam);
        if (macher.matches()) {
            String husb = macher.group(1).trim();
            Indi indi = INDIS.get(husb);
            currentRecord.setHomme(indi);
            indi.getMariages().add(currentRecord);
        }
    }

    private final static Pattern PATTERN_WIFE = Pattern.compile("1 WIFE @(.+)@");

    private final void analyseWIFE(Fam currentRecord, String lineFam) {
        Matcher macher = PATTERN_WIFE.matcher(lineFam);
        if (macher.matches()) {
            String wife = macher.group(1).trim();
            Indi indi = INDIS.get(wife);
            currentRecord.setFemme(indi);
            indi.getMariages().add(currentRecord);
        }
    }

    private final static Pattern PATTERN_CHIL = Pattern.compile("1 CHIL @(.+)@");

    private final void analyseCHIL(Fam currentRecord, String lineFam) {
        Matcher macher = PATTERN_CHIL.matcher(lineFam);
        if (macher.matches()) {
            String chil = macher.group(1).trim();
            Indi indi = INDIS.get(chil);
            currentRecord.getEnfants().add(indi);
            indi.setPere(currentRecord.getHomme());
            indi.setMere(currentRecord.getFemme());
        }
    }

    private final static Pattern PATTERN_NAME = Pattern.compile("(.+)\\s/(.+)/");

    private final void analyseNAME(Indi currentRecord, String lineIndi) {
        String suffix = lineIndi.substring(6);
        Matcher macher = PATTERN_NAME.matcher(suffix);
        if (macher.matches()) {
            String prenom = macher.group(1).trim();
            String nom = macher.group(2).trim();
            currentRecord.setNom(nom);
            currentRecord.setPrenom(prenom);
        }
    }

    private final void analyseSEX(Indi currentRecord, String lineIndi) {
        String suffix = lineIndi.substring(6).trim();
        currentRecord.setSexe(suffix);
    }

    private final void analyseOCCU(Indi currentRecord, String lineIndi) {
        String suffix = lineIndi.substring(6).trim();
        currentRecord.setProfession(suffix);
    }

    private final void analyseDATE(String event, Indi currentRecord, String lineIndi) {
        String suffix = lineIndi.substring(6).trim();
        switch (event) {
            case "BIRT":
                currentRecord.setDateNaissance(convertDate(suffix));
                break;
            case "DEAT":
                currentRecord.setDateDeces(convertDate(suffix));
                break;
        }
    }

    private final void analyseDATE(String event, Fam currentRecord, String lineIndi) {
        String suffix = lineIndi.substring(6).trim();
        switch (event) {
            case "MARR":
                currentRecord.setDateMariage(convertDate(suffix));
                break;
        }
    }

    private final void analysePLAC(String event, Indi currentRecord, String lineIndi) {
        String suffix = lineIndi.substring(6).trim();
        switch (event) {
            case "BIRT":
                currentRecord.setLieuNaissance(suffix);
                break;
            case "DEAT":
                currentRecord.setLieuDeces(suffix);
                break;
        }
    }

    private final void analysePLAC(String event, Fam currentRecord, String lineIndi) {
        String suffix = lineIndi.substring(6).trim();
        switch (event) {
            case "MARR":
                currentRecord.setLieuMariage(suffix);
                break;
        }
    }

    private final static Pattern PATTERN_DATE = Pattern.compile("(\\d+)\\s(.+)\\s(\\d+)");

    private final static String convertDate(String date) {
        // yyyy-MM-dd => dd/MM/yyyy
        if (date != null) {
            Matcher macher = PATTERN_DATE.matcher(date);
            if (macher.matches()) {
                return macher.group(1) + "/" + convertMonth(macher.group(2)) + "/" + macher.group(3);
            } else {
                date = date.replace("BEF", "Avant ");
                date = date.replace("AFT", "Après ");
                date = date.replace("BET", "Entre ");
                date = date.replace("AND", "et ");
                date = date.replace("ABT", "Environ ");
                date = date.replace("EST", "Estimation ");
                date = date.replace(" JAN ", "/01/");
                date = date.replace(" FEB ", "/02/");
                date = date.replace(" MAR ", "/03/");
                date = date.replace(" APR ", "/04/");
                date = date.replace(" MAY ", "/05/");
                date = date.replace(" JUN ", "/06/");
                date = date.replace(" JUL ", "/07/");
                date = date.replace(" AUG ", "/08/");
                date = date.replace(" SEP ", "/09/");
                date = date.replace(" OCT ", "/10/");
                date = date.replace(" NOV ", "/11/");
                date = date.replace(" DEC ", "/12/");
                date = date.replace(" /", " "); // Cas d'une date sans jour
                date = date.replace("  ", " "); // Cas d'une date sans jour
            }
            return date;
        }
        return null;
    }

    private final static String convertMonth(String month) {
        switch (month) {
            case "JAN":
                return "01";
            case "FEB":
                return "02";
            case "MAR":
                return "03";
            case "APR":
                return "04";
            case "MAY":
                return "05";
            case "JUN":
                return "06";
            case "JUL":
                return "07";
            case "AUG":
                return "08";
            case "SEP":
                return "09";
            case "OCT":
                return "10";
            case "NOV":
                return "11";
            case "DEC":
                return "12";
        }
        return month;
    }


}
