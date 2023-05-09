package dev.mflash.poi.dsl;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Hyperlink;

import java.time.LocalDateTime;
import java.util.function.Consumer;
import java.util.function.Supplier;

public class PoiCellWriter {

	private final Cell cell;

	PoiCellWriter(Supplier<Cell> cellSupplier) {
		this.cell = cellSupplier.get();
	}

	public static PoiCellWriter create(Supplier<Cell> cellSupplier) {
		return new PoiCellWriter(cellSupplier);
	}

	public PoiCellWriter style(CellStyle cellStyle) {
		cell.setCellStyle(cellStyle);
		return this;
	}

	public PoiCellWriter value(String value) {
		cell.setCellValue(value);
		return this;
	}

	public PoiCellWriter value(boolean value) {
		cell.setCellValue(value);
		return this;
	}

	public PoiCellWriter value(LocalDateTime value) {
		cell.setCellValue(value);
		return this;
	}

	public PoiCellWriter value(double value) {
		cell.setCellValue(value);
		return this;
	}

	public PoiCellWriter blank() {
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

	public PoiCellWriter compute(Consumer<Cell> cellConsumer) {
		cellConsumer.accept(cell);
		return this;
	}
}
