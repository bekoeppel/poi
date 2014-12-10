package org.subtlelib.poiooxml.impl.sheet;

import static com.google.common.base.Preconditions.checkState;

import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.subtlelib.poiooxml.api.configuration.Configuration;
import org.subtlelib.poiooxml.api.filter.FilterDataRange;
import org.subtlelib.poiooxml.api.row.RowContext;
import org.subtlelib.poiooxml.api.sheet.SheetContext;
import org.subtlelib.poiooxml.api.totals.ColumnTotalsDataRange;
import org.subtlelib.poiooxml.api.workbook.WorkbookContext;
import org.subtlelib.poiooxml.impl.filter.FilterDataRangeImpl;
import org.subtlelib.poiooxml.impl.row.RowContextImpl;
import org.subtlelib.poiooxml.impl.row.RowContextNoImpl;
import org.subtlelib.poiooxml.impl.row.Rows;
import org.subtlelib.poiooxml.impl.style.InheritableStyleConfiguration;
import org.subtlelib.poiooxml.impl.totals.ColumnTotalsDataRangeImpl;

public class SheetContextImpl extends InheritableStyleConfiguration<SheetContext> implements SheetContext {

	private final WorkbookContext workbook;
    
    private final Sheet sheet;
    
    protected RowContext currentRow;
    protected int rowNo = -1;
    
    private int defaultRowIndent = 0;

    private final SheetContext noImplSheetContext;
    private final RowContext noImplRowContext;
    
    public SheetContextImpl(Sheet sheet, WorkbookContext workbook) {
    	super(workbook);
    	
        this.sheet = sheet;
        this.workbook = workbook;
        
        this.noImplSheetContext = new SheetContextNoImpl(this, workbook);
        this.noImplRowContext = new RowContextNoImpl(this);
    }
    
    @Override
	public SheetContext skipRow() {
        rowNo++;
        return this;
    }

    @Override
	public SheetContext skipRows(int offset) {
        this.rowNo += offset;
        return this;
    }

    @Override
    public SheetContext stepOneRowBack() {
        rowNo -= 1;
        currentRow = null;
        return this;
    }

    @Override
	public RowContext nextRow() {
    	++rowNo;
    	currentRow = null; // clear to force init of the currentRow
        return currentRow();
    }

    @Override
	public RowContext nextConditionalRow(boolean condition) {
        return condition ? nextRow() : nextRowNoImpl();
    }
    
    private RowContext nextRowNoImpl() {
        currentRow = noImplRowContext;
        return currentRow;
    }

    @Override
	public RowContext currentRow() {
    	if (rowNo == -1) {
            checkState(currentRow != null, "Current row doesn't exist. Use nextRow() to create a new row");
    	}
    	if (currentRow == null) {
    		currentRow = new RowContextImpl(Rows.getOrCreate(sheet, rowNo), this, workbook, defaultRowIndent);    		
    	}
        return currentRow;
    }

    @Override
	public Sheet getNativeSheet() {
        return sheet;
    }

    @Override
    public SheetContext setColumnWidth(int columnNumber, int width) {
        //TODO max col width in excel is 255 characters; check aforementioned condition with a test (whether our 1.1 doesn't reduce this upper boundary), add precondition check
    	sheet.setColumnWidth(columnNumber, (int) (getConfiguration().getColumnWidthBaseValue() * width));
    	return this;
    }
    
    @Override
	public SheetContext setColumnWidths(int... multipliers) {
    	for (int i = 0; i < multipliers.length; i++) {
            setColumnWidth(i, multipliers[i]);
        }
        return this;
    }
    
	@Override
	public SheetContext mergeCells(int startColumn, int endColumn) {
		sheet.addMergedRegion(new CellRangeAddress(rowNo, rowNo, startColumn, endColumn));
		return this;
	}

	@Override
	public SheetContext hideGrid() {
		sheet.setDisplayGridlines(false);
		return this;
	}
	
	@Override
	public SheetContext setDefaultRowIndent(int indent) {
		this.defaultRowIndent = indent;
		return this;
	}

    @Override
    public SheetContext fitOnPagesByWidth(int pages) {
        PrintSetup printSetup = sheet.getPrintSetup();
        sheet.setAutobreaks(true);
        printSetup.setFitWidth((short) pages);
        return this;
    }

    @Override
    public SheetContext fitOnPagesByHeight(int pages) {
        PrintSetup printSetup = sheet.getPrintSetup();
        sheet.setAutobreaks(true);
        printSetup.setFitHeight((short) pages);
        return this;
    }

    @Override
	public SheetContext startConditionalBlock(boolean condition) {
		return condition ? this : noImplSheetContext;
	}

	@Override
	public SheetContext endConditionalBlock() {
		return this;
	}

	@Override
    public Configuration getConfiguration() {
		return workbook.getConfiguration();
	}

    @Override
    public ColumnTotalsDataRange startColumnTotalsDataRangeFromNextRow() {
        return new ColumnTotalsDataRangeImpl(this);
    }

    public int getCurrentRowNo() {
        return rowNo;
    }

    @Override
    public FilterDataRange startFilterDataRangeFromCurrentCell() {
        return new FilterDataRangeImpl(this);
    }
}
