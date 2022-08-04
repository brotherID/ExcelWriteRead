package org.manage.file.ManageFile.model;



import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data 
@AllArgsConstructor
@NoArgsConstructor
public class Cheque {
	private String identifantCheque;
	private String cmc7;
	private String endos;
	private double montant;

}
