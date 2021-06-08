package com.fanch35000.genealogycheckcivilstatus.entity.civilstatus;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class CivilStatus {

    // Server api specific data
    private Integer sosa;
    private String nom;
    private String prenom;
    private String dateNaissance;
    private String villeRegistreNaissance;
    private String fichierActeNaissance;
    private String dateMariage;
    private String villeRegistreMariage;
    private String fichierActeMariage;
    private String dateDeces;
    private String villeRegistreDeces;
    private String fichierActeDeces;

}
