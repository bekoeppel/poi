package org.subtlelib.poiooxml.impl.style.system;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.subtlelib.poiooxml.api.style.AdditiveStyle;

public enum SystemCellWrapTextStyle implements AdditiveStyle {
	
	WRAP_TEXT;
	
	@Override
	public void enrich(XSSFWorkbook workbook, CellStyle style) {
		style.setWrapText(true);
	}

	@Override
	public Enum<SystemStyleType> getType() {
		return SystemStyleType.CELL_WRAP_TEXT;
	}

}
