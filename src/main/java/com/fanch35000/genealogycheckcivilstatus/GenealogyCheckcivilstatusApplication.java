package com.fanch35000.genealogycheckcivilstatus;

import com.fanch35000.genealogycheckcivilstatus.entity.civilstatus.CivilStatus;
import com.fanch35000.genealogycheckcivilstatus.entity.ged.Indi;
import com.fanch35000.genealogycheckcivilstatus.service.CivilStatusService;
import com.fanch35000.genealogycheckcivilstatus.service.CreateExcelService;
import com.fanch35000.genealogycheckcivilstatus.service.GedService;
import com.fanch35000.genealogycheckcivilstatus.service.UpdateGedService;

import java.util.HashMap;

public class GenealogyCheckcivilstatusApplication {

	public static void main (String[] args){

		// List des fichiers
		String dirLocation = "/home/fminier/Perso/genealogie/data/actes";
		HashMap<Integer, CivilStatus> ancetres = new CivilStatusService().launch(dirLocation);

		String gedLocation = "/home/fminier/Perso/genealogie/data/ged/fminier-ASC.ged";
		HashMap<String, Indi> indis = new GedService().launch(gedLocation);

		String xlsLocation = "/home/fminier/Perso/genealogie/data/ancetres.xls";
		new CreateExcelService().launch(xlsLocation, ancetres, indis);

		String newGedLocation = "/home/fminier/Perso/genealogie/data/ged/fminier-ASC-new.ged";
		new UpdateGedService().launch(gedLocation, newGedLocation, ancetres, indis);

	}

}
