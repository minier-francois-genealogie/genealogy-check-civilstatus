package com.fanch35000.genealogycheckcivilstatus.service;

import com.fanch35000.genealogycheckcivilstatus.entity.civilstatus.CivilStatus;
import com.fanch35000.genealogycheckcivilstatus.entity.ged.Fam;
import com.fanch35000.genealogycheckcivilstatus.entity.ged.Indi;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.io.Writer;
import java.util.HashMap;
import java.util.Scanner;
import java.util.regex.Matcher;

public class UpdateGedService extends GedService {

    public final void launch(String gedLocation, String newGedLocation, HashMap<Integer, CivilStatus> ancetres, HashMap<String, Indi> indis) {
        try {
            File fileGed = new File(gedLocation);
            Writer fileNewGed = new FileWriter(newGedLocation);

            completeFile(fileGed, fileNewGed, ancetres, indis);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private final void completeFile(File fileGed, Writer fileNewGed, HashMap<Integer, CivilStatus> ancetres, HashMap<String, Indi> indis) {
        try {
            Scanner scanner = new Scanner(fileGed);

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
                            currentIndi = indis.get(macher.group(1));
                        }
                        currentRecord = currentIndi;
                    } else if (line.contains("FAM")) {
                        currentRecord = new Fam();
                    } else {
                        currentRecord = null;
                    }
                    // copie de la ligne
                    fileNewGed.write(line + "\n");
                } else {
                    if (line.length() >= 6) {
                        String prefix = line.substring(0, 6);
                        if (currentRecord instanceof Indi) {
                            Indi currentIndi = (Indi) currentRecord;
                            if (prefix.equals("2 SOUR") || prefix.equals("2 NOTE")) {
                                // pas de copie de la ligne
                            } else {
                                // copie de la ligne
                                fileNewGed.write(line + "\n");
                                Integer sosa = CreateExcelService.getMinSosa(currentIndi.getSosas());
                                switch (prefix) {
                                    case "1 BIRT":
                                        if (ancetres.get(sosa) != null)
                                            writeSource(fileNewGed, "Acte de naissance", ancetres.get(sosa).getFichierActeNaissance());
                                        break;
                                    case "1 DEAT":
                                        //Création ligne si acte
                                        if (ancetres.get(sosa) != null)
                                            writeSource(fileNewGed, "Acte de décès", ancetres.get(sosa).getFichierActeDeces());
                                        break;
                                }
                            }
                        } else if (currentRecord instanceof Fam) {
                            if (prefix.equals("2 SOUR") || prefix.equals("2 NOTE")) {
                                // pas de copie de la ligne
                            } else {
                                // copie de la ligne
                                fileNewGed.write(line + "\n");

                                switch (prefix) {
                                    case "1 HUSB":
                                        //Création ligne si acte
                                        Integer sosa = null;
                                        Matcher macher = PATTERN_HUSB.matcher(line);
                                        if (macher.matches()) {
                                            sosa = CreateExcelService.getMinSosa(indis.get(macher.group(1).trim()).getSosas());
                                        }
                                        if (ancetres.get(sosa) != null)
                                            writeSource(fileNewGed, "Acte de mariage", ancetres.get(sosa).getFichierActeDeces());
                                        break;
                                }
                            }
                        } else {
                            // copie de la ligne
                            fileNewGed.write(line + "\n");
                        }
                    }
                }
            }

            scanner.close();
            fileNewGed.flush();
            fileNewGed.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    private final static void writeSource(Writer fileNewGed, String title, String acte) throws IOException {
        if (acte != null) {
            String link = "<a href=\"https://raw.githubusercontent.com/minier-francois-genealogie/data/main/actes/" + acte + "\">";
            link += title;
            link += "</a>";
            fileNewGed.write("2 SOUR " + link + "\n");
        }
    }

}

