package org.subtlelib.poi.impl.style.system;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.subtlelib.poi.api.style.AdditiveStyle;

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
