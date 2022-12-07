package dev.mflash.poi.dsl;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;

import java.time.LocalDateTime;
import java.util.Objects;

public final class PoiCellWriter {

	private final Cell cell;

	public PoiCellWriter(Sheet sheet, CellReference cellReference) {
		final Cell c = sheet.getRow(cellReference.getRow()).getCell(cellReference.getCol());
		this.cell = Objects.nonNull(c) ? c : sheet.getRow(cellReference.getRow()).createCell(cellReference.getCol());
	}

	public void stringValue(String value) {
		cell.setCellValue(value);
	}

	public void booleanValue(boolean value) {
		cell.setCellValue(value);
	}

	public void dateTimeValue(LocalDateTime value) {
		cell.setCellValue(value);
	}

	public void integerValue(int value) {
		cell.setCellValue(value);
	}

	public void longValue(long value) {
		cell.setCellValue(value);
	}

	public void doubleValue(double value) {
		cell.setCellValue(value);
	}
}
