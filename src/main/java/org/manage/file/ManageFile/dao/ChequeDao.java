package org.manage.file.ManageFile.dao;

import java.util.ArrayList;
import java.util.List;

import org.manage.file.ManageFile.model.Cheque;

public class ChequeDao {
	  public static List<Cheque> listCheques() {
	        List<Cheque> list = new ArrayList<Cheque>();

	        Cheque c1 = new Cheque("c01", "qwrew34", "OLK98",400.8);
	        Cheque c2 = new Cheque("c02", "hjjk23", "were56",500.9);
	        Cheque c3 = new Cheque("c03", "deer56", "mhgf12",300.8);
	        list.add(c1);
	        list.add(c2);
	        list.add(c3);
	        return list;
	    }

}
