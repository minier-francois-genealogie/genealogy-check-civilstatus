package com.fanch35000.genealogycheckcivilstatus.service;

import com.fanch35000.genealogycheckcivilstatus.entity.civilstatus.CivilStatus;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

public class CivilStatusService {

    public Set<Integer> listSosa = new HashSet<>();
    public HashMap<Integer, CivilStatus> ancetres = new HashMap();

    public final void launch(String dirLocation) {
        List<File> files = listEtatCivil(dirLocation);
        files.forEach(f -> analyseFile(f));
    }

    private final static List<File> listEtatCivil(String dirLocation) {

        try {
            List<File> files = Files.list(Paths.get(dirLocation))
                    .filter(Files::isRegularFile)
                    .map(Path::toFile)
                    .collect(Collectors.toList());
            return files;
        } catch (IOException e) {
            e.printStackTrace();
        }
        return new ArrayList<>();
    }

    private final void analyseFile(File file) {
        String fileName = file.getName();

        String pattern = "(\\S*)_([^\\.]+)\\.sosa_(\\d+)\\.([\\d-]+)\\.(.{1})(\\..+){0,1}\\.jpg";
        // String pattern = ".*";

        // Create a Pattern object
        Pattern r = Pattern.compile(pattern);
        Matcher m = r.matcher(fileName);

        if (m.matches()) {

            String nom = m.group(1);
            String prenom = m.group(2);
            Integer sosa = Integer.parseInt(m.group(3));
            listSosa.add(sosa);
            String date = m.group(4);
            String type = m.group(5);
            String villeRegistre = m.group(6);
            if (villeRegistre != null) {
                villeRegistre = villeRegistre.replaceFirst("\\.", "");
            }

            CivilStatus individu = ancetres.get(sosa);
            if (individu == null) {
                individu = new CivilStatus();
                individu.setNom(nom);
                individu.setPrenom(prenom);
                individu.setSosa(sosa);
                ancetres.put(sosa, individu);
            }
            switch (type) {
                case "N":
                    individu.setDateNaissance(date);
                    individu.setVilleRegistreNaissance(villeRegistre);
                    individu.setFichierActeNaissance(fileName);
                    break;
                case "M":
                    individu.setDateMariage(date);
                    individu.setVilleRegistreMariage(villeRegistre);
                    individu.setFichierActeMariage(fileName);
                    break;
                case "D":
                    individu.setDateDeces(date);
                    individu.setVilleRegistreDeces(villeRegistre);
                    individu.setFichierActeDeces(fileName);
                    break;
            }
        } else {
            System.err.println(fileName);
        }

    }

}
