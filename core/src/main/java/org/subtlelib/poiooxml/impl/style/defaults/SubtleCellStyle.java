package org.subtlelib.poiooxml.impl.style.defaults;

import static org.apache.poi.ss.usermodel.CellStyle.BORDER_DOUBLE;
import static org.apache.poi.ss.usermodel.CellStyle.BORDER_NONE;
import static org.apache.poi.ss.usermodel.CellStyle.BORDER_THIN;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.subtlelib.poiooxml.api.style.AdditiveStyle;

public enum SubtleCellStyle implements AdditiveStyle {

	BORDERS_THIN_ALL(BORDER_THIN, BORDER_THIN, BORDER_THIN, BORDER_THIN),
	BORDERS_BOTTOM_THIN(BORDER_NONE, BORDER_NONE, BORDER_THIN, BORDER_NONE),
	BORDERS_TOP_BOTTOM_THIN(BORDER_THIN, BORDER_NONE, BORDER_THIN, BORDER_NONE),
	BORDERS_BOTTOM_DOUBLE(BORDER_NONE, BORDER_NONE, BORDER_DOUBLE, BORDER_NONE);

    private final short borderTop;
    private final short borderRight;
    private final short borderBottom;
    private final short borderLeft;

    private SubtleCellStyle(short borderTop, short borderRight, short borderBottom, short borderLeft) {
        this.borderTop = borderTop;
        this.borderRight = borderRight;
        this.borderBottom = borderBottom;
        this.borderLeft = borderLeft;
    }

	@Override
	public void enrich(XSSFWorkbook workbook, CellStyle style) {
		style.setBorderTop(borderTop);
        style.setBorderRight(borderRight);
        style.setBorderBottom(borderBottom);
        style.setBorderLeft(borderLeft);
    }

	@Override
	public Enum<StyleType> getType() {
		return StyleType.CELL;
	}

}
