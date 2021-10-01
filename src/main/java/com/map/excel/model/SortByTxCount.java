package com.map.excel.model;

import java.util.Comparator;

public class SortByTxCount implements Comparator<RowDataForCumalative> {

	@Override
	public int compare(RowDataForCumalative o1, RowDataForCumalative o2) {
		int returnValue = o2.getIndividualPercentageForTxCount().compareTo(o1.getIndividualPercentageForTxCount());
		if (returnValue == 0)
			return o1.getTrxCode().compareTo(o2.getTrxCode());
		return returnValue;
	}

}
