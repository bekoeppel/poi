package org.subtlelib.poi.impl.style.defaults;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.subtlelib.poi.api.style.Style;

/**
 * An empty style that does nothing.
 * Created on 10/04/13
 * @author d.serdiuk
 */
public class EmptyStyle implements Style {
    @Override
    public void enrich(XSSFWorkbook workbook, CellStyle style) {}
    public static final Style instance = new EmptyStyle();
}
