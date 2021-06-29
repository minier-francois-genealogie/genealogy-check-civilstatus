package com.fanch35000.genealogycheckcivilstatus;

import com.fanch35000.genealogycheckcivilstatus.entity.civilstatus.CivilStatus;
import com.fanch35000.genealogycheckcivilstatus.entity.ged.Indi;
import com.fanch35000.genealogycheckcivilstatus.service.CivilStatusService;
import com.fanch35000.genealogycheckcivilstatus.service.CreateExcelService;
import com.fanch35000.genealogycheckcivilstatus.service.GedService;
import com.fanch35000.genealogycheckcivilstatus.service.UpdateGedService;

import java.util.*;
import java.util.stream.Collectors;

public class GenealogyCheckcivilstatusApplication {

    public static void main(String[] args) {

        // List des fichiers
        String dirLocation = "/home/fminier/Perso/genealogie/data/actes";
        CivilStatusService civilStatusService = new CivilStatusService();
        civilStatusService.launch(dirLocation);

        String gedLocation = "/home/fminier/Perso/genealogie/data/ged/fminier-ASC.ged";
        HashMap<String, Indi> indis = new GedService().launch(gedLocation);

        String xlsLocation = "/home/fminier/Perso/genealogie/data/ancetres.xls";
        CreateExcelService createExcelService = new CreateExcelService();
        createExcelService.launch(xlsLocation, civilStatusService.ancetres, indis);

        String newGedLocation = "/home/fminier/Perso/genealogie/data/ged/fminier-ASC-new.ged";
        UpdateGedService updateGedService = new UpdateGedService();
        updateGedService.launch(gedLocation, newGedLocation, civilStatusService.ancetres, indis);

        /*********
         * CHECK *
         *********/

        // Les SOSA sur les actes papiers sont retrouvés dans les GED
        for (Integer sosa : civilStatusService.listSosa) {
            if (!createExcelService.listSosa.contains(sosa))
                System.err.println("SOSA dans fichier JPG non présent dans GED : " + sosa);
        }

        // Les actes papiers n'ont pas pu être ajouté dans le GED
        // Pas de rubrique BIRT DEAT ou MARR dans le GED
        final List<String> listFichierJpg = new ArrayList<>();
        civilStatusService.ancetres.values().stream().sorted((s1, s2) -> s1.getSosa().compareTo(s2.getSosa())).forEach(status -> addActe(status, listFichierJpg));
        listFichierJpg.removeAll(updateGedService.listFichier);
        System.err.println();
        System.err.println("********************************************************************");
        System.err.println("* Liste des fichiers JPG non ajouté(s) dans le GED (triés par sosa):");
        System.err.println("* ");
        for (String fichier : listFichierJpg) {
            System.err.println("* " + fichier);
        }
        System.err.println("* ");
        System.err.println("* TOTAL de date à ajouter : " + listFichierJpg.size());
        System.err.println("***************************");

    }

    private final static void addActe(CivilStatus actes, List<String> listFichier){
        if(actes.getFichierActeNaissance()!=null){
            listFichier.add(actes.getFichierActeNaissance());
        }
        if(actes.getFichierActeMariage()!=null){
            listFichier.add(actes.getFichierActeMariage());
        }
        if(actes.getFichierActeDeces()!=null){
            listFichier.add(actes.getFichierActeDeces());
        }
    }

}
