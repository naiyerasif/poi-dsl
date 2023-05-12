package dev.mflash.poi.dsl;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.util.CellUtil;

import java.time.LocalDateTime;
import java.util.Map;
import java.util.function.Consumer;
import java.util.function.Supplier;

public final class PoiCellWriter {

	private final Cell cell;

	PoiCellWriter(Cell cell) {
		this.cell = cell;
	}

	public static PoiCellWriter create(Supplier<Cell> cellSupplier) {
		return new PoiCellWriter(cellSupplier.get());
	}

	public PoiCellWriter map(Consumer<Cell> cellMapper) {
		cellMapper.accept(cell);
		return this;
	}

	public PoiCellWriter styleProperty(String propKey, Object propValue) {
		CellUtil.setCellStyleProperty(cell, propKey, propValue);
		return this;
	}

	public PoiCellWriter styleProperties(Map<String, Object> props) {
		CellUtil.setCellStyleProperties(cell, props);
		return this;
	}

	public PoiCellWriter style(CellStyle cellStyle) {
		cell.setCellStyle(cellStyle);
		return this;
	}

	public PoiCellWriter stringValue(String value) {
		cell.setCellValue(value);
		return this;
	}

	public PoiCellWriter booleanValue(boolean value) {
		cell.setCellValue(value);
		return this;
	}

	public PoiCellWriter dateTimeValue(LocalDateTime value) {
		cell.setCellValue(value);
		return this;
	}

	public PoiCellWriter numericValue(double value) {
		cell.setCellValue(value);
		return this;
	}

	public PoiCellWriter blankValue() {
		cell.setBlank();
		return this;
	}

	public PoiCellWriter hyperlink(Hyperlink hyperlink) {
		cell.setHyperlink(hyperlink);
		return this;
	}

	public PoiCellWriter comment(Comment comment) {
		cell.setCellComment(comment);
		return this;
	}
}
