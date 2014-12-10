package org.subtlelib.poiooxml.impl.style.defaults;

import static org.apache.poi.xssf.usermodel.XSSFFont.DEFAULT_FONT_NAME;
import static org.apache.poi.ss.usermodel.CellStyle.ALIGN_LEFT;
import static org.apache.poi.ss.usermodel.CellStyle.ALIGN_RIGHT;
import static org.apache.poi.ss.usermodel.CellStyle.VERTICAL_BOTTOM;
import static org.apache.poi.ss.usermodel.Font.BOLDWEIGHT_BOLD;
import static org.apache.poi.ss.usermodel.Font.BOLDWEIGHT_NORMAL;
import static org.apache.poi.ss.usermodel.Font.U_NONE;
import static org.apache.poi.ss.usermodel.Font.U_SINGLE;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.subtlelib.poiooxml.api.style.AdditiveStyle;

public enum FontStyle implements AdditiveStyle {
	NORMAL(DEFAULT_FONT_NAME, 10, BOLDWEIGHT_NORMAL, U_NONE, ALIGN_LEFT, VERTICAL_BOTTOM),
    ITALIC(DEFAULT_FONT_NAME, 10, BOLDWEIGHT_NORMAL, U_NONE, ALIGN_LEFT, VERTICAL_BOTTOM, true),
	NORMAL_RIGHT(DEFAULT_FONT_NAME, 10, BOLDWEIGHT_NORMAL, U_NONE, ALIGN_RIGHT, VERTICAL_BOTTOM),
	BOLD(DEFAULT_FONT_NAME, 10, BOLDWEIGHT_BOLD, U_NONE, ALIGN_LEFT, VERTICAL_BOTTOM),
	BOLD_RIGHT(DEFAULT_FONT_NAME, 10, BOLDWEIGHT_BOLD, U_NONE, ALIGN_RIGHT, VERTICAL_BOTTOM),

	COLUMN_HEADER(DEFAULT_FONT_NAME, 10, BOLDWEIGHT_BOLD, U_NONE, ALIGN_LEFT, VERTICAL_BOTTOM),
	COLUMN_HEADER_RIGHT(DEFAULT_FONT_NAME, 10, BOLDWEIGHT_BOLD, U_NONE, ALIGN_RIGHT, VERTICAL_BOTTOM),
	SECTION_HEADER(DEFAULT_FONT_NAME, 10, BOLDWEIGHT_BOLD, U_SINGLE, ALIGN_LEFT, VERTICAL_BOTTOM),
	SECTION_HEADER_RIGHT(DEFAULT_FONT_NAME, 10, BOLDWEIGHT_BOLD, U_SINGLE, ALIGN_RIGHT, VERTICAL_BOTTOM),
    PRIMARY_HEADER(DEFAULT_FONT_NAME, 12, BOLDWEIGHT_BOLD, U_SINGLE, ALIGN_LEFT, VERTICAL_BOTTOM),
    SECONDARY_HEADER(DEFAULT_FONT_NAME, 14, BOLDWEIGHT_BOLD, U_NONE, ALIGN_LEFT, VERTICAL_BOTTOM);

    private final String name;
    private final short height;
    private final short boldWeight;
    private final byte underline;
    private final short horizontalAlignment;
    private final short verticalAlignment;
    private final boolean italic;

	private FontStyle(String name, Integer height, short boldWeight, byte underline, short horizontalAlignment,
                      short verticalAlignment) {
        this(name, height, boldWeight, underline, horizontalAlignment, verticalAlignment, false);
	}

    private FontStyle(String name, Integer height, short boldWeight, byte underline, short horizontalAlignment,
                      short verticalAlignment, boolean italic) {
        this.name = name;
        this.height = height.shortValue();
        this.boldWeight = boldWeight;
        this.underline = underline;
        this.horizontalAlignment = horizontalAlignment;
        this.verticalAlignment = verticalAlignment;
        this.italic = italic;
    }

	@Override
	public void enrich(XSSFWorkbook workbook, CellStyle style) {
        XSSFFont font = workbook.createFont();
        font.setFontName(name);
        font.setFontHeightInPoints(height);
        font.setBoldweight(boldWeight);
        font.setUnderline(underline);
        font.setItalic(italic);
        style.setFont(font);

        style.setAlignment(horizontalAlignment);
        style.setVerticalAlignment(verticalAlignment);
	}

	@Override
	public Enum<StyleType> getType() {
		return StyleType.FONT;
	}
}
