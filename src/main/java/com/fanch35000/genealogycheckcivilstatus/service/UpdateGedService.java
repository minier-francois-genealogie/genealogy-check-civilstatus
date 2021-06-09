package com.fanch35000.genealogycheckcivilstatus.service;

import com.fanch35000.genealogycheckcivilstatus.entity.civilstatus.CivilStatus;
import com.fanch35000.genealogycheckcivilstatus.entity.ged.Fam;
import com.fanch35000.genealogycheckcivilstatus.entity.ged.Indi;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.io.Writer;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Scanner;
import java.util.Set;
import java.util.regex.Matcher;

public class UpdateGedService extends GedService {

    public Set<String> listFichier = new HashSet<>();

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

                boolean keepLine = !(line.startsWith("1 SOUR")
                        || line.startsWith("1 NOTE")
                        || line.startsWith("1 CONT")
                        || line.startsWith("2 SOUR")
                        || line.startsWith("2 NOTE")
                        || line.startsWith("2 CONT")
                        || line.startsWith("3 SOUR")
                        || line.startsWith("3 NOTE")
                        || line.startsWith("3 CONT"));

                if (keepLine) {

                    // copie de la ligne
                    fileNewGed.write(line + "\n");

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
                    } else {
                        if (line.length() >= 6) {
                            String prefix = line.substring(0, 6);
                            if (currentRecord instanceof Indi) {
                                CivilStatus civilStatus;
                                switch (prefix) {
                                    case "1 BIRT":
                                        civilStatus = ancetres.get(CreateExcelService.getMinSosa(((Indi) currentRecord).getSosas()));
                                        if (civilStatus != null)
                                            writeSource(fileNewGed, "Acte de naissance", civilStatus.getFichierActeNaissance());
                                        break;
                                    case "1 DEAT":
                                        civilStatus = ancetres.get(CreateExcelService.getMinSosa(((Indi) currentRecord).getSosas()));
                                        if (civilStatus != null)
                                            writeSource(fileNewGed, "Acte de décès", civilStatus.getFichierActeDeces());
                                        break;
                                }
                            } else if (currentRecord instanceof Fam) {
                                Integer sosa = null;
                                Matcher macher = PATTERN_HUSB.matcher(line);
                                if (macher.matches()) {
                                    sosa = CreateExcelService.getMinSosa(indis.get(macher.group(1).trim()).getSosas());
                                }
                                CivilStatus civilStatus;
                                switch (prefix) {
                                    case "1 HUSB":
                                        civilStatus = ancetres.get(sosa);
                                        if (civilStatus != null)
                                            writeSource(fileNewGed, "Acte de mariage", civilStatus.getFichierActeMariage());
                                        break;
                                }
                            }
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


    private final void writeSource(Writer fileNewGed, String title, String acte) throws IOException {
        if (acte != null) {
            listFichier.add(acte);
            String link = "<a href=\"https://raw.githubusercontent.com/minier-francois-genealogie/data/main/actes/" + acte + "\">";
            link += title;
            link += "</a>";
            fileNewGed.write("2 SOUR " + link + "\n");
        }
    }

}

