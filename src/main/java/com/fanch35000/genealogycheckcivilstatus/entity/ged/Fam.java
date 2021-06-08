package com.fanch35000.genealogycheckcivilstatus.entity.ged;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.ArrayList;
import java.util.List;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class Fam {

    private Indi homme;
    private Indi femme;
    private List<Indi> enfants = new ArrayList();
    private String dateMariage;
    private String lieuMariage;
}
