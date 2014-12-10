package org.subtlelib.poiooxml.impl.workbook;

import static com.google.common.base.Preconditions.checkArgument;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.subtlelib.poiooxml.api.configuration.Configuration;
import org.subtlelib.poiooxml.api.sheet.SheetContext;
import org.subtlelib.poiooxml.api.style.Style;
import org.subtlelib.poiooxml.api.style.StyleConfiguration;
import org.subtlelib.poiooxml.api.workbook.WorkbookContext;
import org.subtlelib.poiooxml.impl.sheet.SheetContextImpl;
import org.subtlelib.poiooxml.impl.style.InheritableStyleConfiguration;

public class WorkbookContextImpl extends InheritableStyleConfiguration<WorkbookContext> implements WorkbookContext {

    private final XSSFWorkbook workbook;
    
    private final Map<Style, CellStyle> registeredStyles = new HashMap<Style, CellStyle>();
    
    private final Configuration configuration;
    
    protected WorkbookContextImpl(XSSFWorkbook workbook, StyleConfiguration styleConfiguration, Configuration configuration) {
    	super(styleConfiguration);
        this.workbook = workbook;
        this.configuration = configuration;
    }
    
    @Override
    public SheetContext createSheet(String sheetName) {
        return new SheetContextImpl(workbook.createSheet(sheetName), this);
    }
    
	@Override
	public SheetContext useSheet(String sheetName) {
		XSSFSheet sheet = workbook.getSheet(sheetName);
		checkArgument(sheet != null, "Sheet %s doesn't exist in workbook", sheetName);
		return new SheetContextImpl(sheet, this);
	}    
    
	@Override
	public CellStyle registerStyle(Style style) {
		checkArgument(style != null, "Style is null");

		CellStyle registeredStyle = registeredStyles.get(style);
		
		if (registeredStyle == null) {
			registeredStyle = workbook.createCellStyle();
			style.enrich(workbook, registeredStyle);
			registeredStyles.put(style, registeredStyle);
		}
		
		return registeredStyle;
	}
    
    @Override
    public byte[] toNativeBytes() {
    	try {
    		ByteArrayOutputStream baos = new ByteArrayOutputStream();
    		workbook.write(baos);
    		return baos.toByteArray();
    	} catch (IOException e) {
    		throw new RuntimeException("Quite unlikely case as we are working with an in-memory data. Wrap to avoid handling checked exception", e);
		}
    }

	@Override
	public XSSFWorkbook toNativeWorkbook() {
		return workbook;
	}

	@Override
	public Configuration getConfiguration() {
		return configuration;
	}


}
