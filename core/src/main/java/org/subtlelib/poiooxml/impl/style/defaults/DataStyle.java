package org.subtlelib.poiooxml.impl.style.defaults;

import java.util.Locale;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.DateFormatConverter;
import org.subtlelib.poiooxml.api.style.AdditiveStyle;

public enum DataStyle implements AdditiveStyle {

	INTEGER("0"),
    AMOUNT("#,##0.00"),
	PERCENTAGE("0.00000%"),

	DATE(DateFormatConverter.convert(Locale.US, "dd-mmm-yyyy")),
	DATE_NUMERIC(DateFormatConverter.convert(Locale.US, "dd-mm-yyyy"));

    private final String format;

    private DataStyle(String format) {
        this.format = format;
    }

	@Override
	public void enrich(XSSFWorkbook workbook, CellStyle style) {
        XSSFDataFormat dataFormat = workbook.createDataFormat();
        style.setDataFormat(dataFormat.getFormat(format));
	}

	@Override
	public Enum<StyleType> getType() {
		return StyleType.DATA;
	}
}
