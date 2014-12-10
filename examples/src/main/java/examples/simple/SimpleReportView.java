package examples.simple;

import java.util.Collection;

import org.subtlelib.poiooxml.api.sheet.SheetContext;
import org.subtlelib.poiooxml.api.totals.ColumnTotalsDataRange;
import org.subtlelib.poiooxml.api.totals.Formula;
import org.subtlelib.poiooxml.api.workbook.WorkbookContext;
import org.subtlelib.poiooxml.impl.workbook.WorkbookContextFactory;

import common.Payment;

/**
 * This example shows the basics of subtlelib. First headers are rendered, then the data rows
 * (each row corresponds to 1 domain object).
 *
 * Created on 15/05/13
 * @author d.serdiuk
 */
public class SimpleReportView {
    public WorkbookContext render(Collection<Payment> payments) {
        WorkbookContext workbookCtx = WorkbookContextFactory.createWorkbook();
        SheetContext sheetCtx = workbookCtx.createSheet("Payments");

        // heading
        sheetCtx
            .nextRow()
                .skipCell()
                .header("Amount")
                .header("Currency")
                .header("Beneficiary").setColumnWidth(25)
                .header("Payee bank").setColumnWidth(35);

        ColumnTotalsDataRange totalsData = sheetCtx.startColumnTotalsDataRangeFromNextRow();

        // data
        for (Payment payment : payments) {
            sheetCtx
                .nextRow()
                    .skipCell()
                    .number(payment.getAmount())
                    .text(payment.getCurrency())
                    .text(payment.getBeneficiary())
                    .text(payment.getPayeeBank());
        }

        sheetCtx
            .nextRow().setTotalsDataRange(totalsData)
                .header("Total:")
                .total(Formula.SUM);

        return workbookCtx;
    }
}
