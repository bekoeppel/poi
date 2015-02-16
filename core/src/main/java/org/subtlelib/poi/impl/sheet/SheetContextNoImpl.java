package org.subtlelib.poi.impl.sheet;

import org.apache.poi.ss.usermodel.Sheet;
import org.subtlelib.poi.api.configuration.Configuration;
import org.subtlelib.poi.api.filter.FilterDataRange;
import org.subtlelib.poi.api.row.RowContext;
import org.subtlelib.poi.api.sheet.SheetContext;
import org.subtlelib.poi.api.totals.ColumnTotalsDataRange;
import org.subtlelib.poi.api.workbook.WorkbookContext;
import org.subtlelib.poi.impl.row.RowContextNoImpl;
import org.subtlelib.poi.impl.style.InheritableStyleConfiguration;

public class SheetContextNoImpl extends InheritableStyleConfiguration<SheetContext> implements SheetContext {

	private final SheetContext parent;
	
	private final RowContext noImplRowContext = new RowContextNoImpl(this);
	
	public SheetContextNoImpl(SheetContext parent, WorkbookContext workbook) {
    	super(workbook);		
		this.parent = parent;
	}

	@Override
	public RowContext nextRow() {
		return noImplRowContext;
	}

	@Override
	public RowContext nextConditionalRow(boolean condition) {
		return noImplRowContext;
	}

	@Override
	public RowContext currentRow() {
		return noImplRowContext;
	}
	
	@Override
	public SheetContext skipRow() {
		return this;
	}

	@Override
	public SheetContext skipRows(int offset) {
		return this;
	}

    @Override
    public SheetContext stepOneRowBack() {
        return this;
    }

    @Override
	public SheetContext startConditionalBlock(boolean condition) {
		return this;
	}

	@Override
	public SheetContext endConditionalBlock() {
		return parent;
	}
	
	@Override
	public Sheet getNativeSheet() {
		throw new UnsupportedOperationException("Native sheet is not supported by SheetContextNoImpl");
	}

    @Override
    public SheetContext setColumnWidth(int columnNumber, int width) {
    	return this;
    }
	
	@Override
	public SheetContext setColumnWidths(int... multipliers) {
		return this;
	}
	
	@Override
	public SheetContext mergeCells(int startColumn, int endColumn) {
		return this;
	}

	@Override
	public SheetContext hideGrid() {
		return this;
	}
	
	@Override
	public SheetContext setDefaultRowIndent(int indent) {	
		return this;
	}

    @Override
    public SheetContext fitOnPagesByWidth(int pages) {
        return this;
    }

    @Override
    public SheetContext fitOnPagesByHeight(int pages) {
        return this;
    }

    @Override
    public Configuration getConfiguration() {
        throw new UnsupportedOperationException("Configuration is not supported by SheetContextNoImpl");
    }

    @Override
    public ColumnTotalsDataRange startColumnTotalsDataRangeFromNextRow() {
        throw new UnsupportedOperationException("Totals unsupported in conditional blocks");
    }

	@Override
	public FilterDataRange startFilterDataRangeFromCurrentCell() {
		throw new UnsupportedOperationException("Filters unsupported in conditional blocks");
	}
}
