package com.fanch35000.genealogycheckcivilstatus.entity.ged;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.ArrayList;
import java.util.List;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class Indi {

    private String idGed;
    private List<Integer> sosas = new ArrayList<>();
    private String nom;
    private String prenom;
    private String sexe;
    private String dateNaissance;
    private String lieuNaissance;
    private String dateDeces;
    private String lieuDeces;
    private String profession;

    // Link
    private Indi pere;
    private Indi mere;
    private List<Fam> mariages = new ArrayList();

}
