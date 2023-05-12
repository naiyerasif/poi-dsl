package dev.mflash.poi.dsl;

import org.apache.poi.ss.usermodel.Cell;

import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Collection;
import java.util.HashMap;
import java.util.Map;
import java.util.function.Function;
import java.util.function.Supplier;
import java.util.stream.Stream;

public final class PoiColumnReader {

	private final Collection<Cell> cells;

	PoiColumnReader(Collection<Cell> cells) {
		this.cells = cells;
	}

	public static PoiColumnReader create(Supplier<Collection<Cell>> columnSupplier) {
		return new PoiColumnReader(columnSupplier.get());
	}

	private Stream<PoiCellReader> readers() {
		return cells.stream().map(PoiCellReader::new);
	}

	public Stream<String> stringValues() {
		return readers().map(PoiCellReader::stringValue);
	}

	public Stream<String> stringValues(DateTimeFormatter formatter) {
		return readers().map(cellReader -> cellReader.stringValue(formatter));
	}

	public Stream<Boolean> booleanValues() {
		return readers().map(PoiCellReader::booleanValue);
	}

	public Stream<LocalDateTime> dateTimeValues() {
		return readers().map(PoiCellReader::dateTimeValue);
	}

	public Stream<BigDecimal> numericValues() {
		return readers().map(PoiCellReader::numericValue);
	}

	private <T> Map<Integer, T> valuesByRowNum(Function<PoiCellReader, T> valueSupplier) {
		final Map<Integer, T> values = new HashMap<>(cells.size());
		for (Cell cell : cells) {
			values.put(cell.getRowIndex(), valueSupplier.apply(new PoiCellReader(cell)));
		}
		return values;
	}

	public Map<Integer, String> stringValuesByRowNum() {
		return valuesByRowNum(PoiCellReader::stringValue);
	}

	public Map<Integer, String> stringValuesByRowNum(DateTimeFormatter formatter) {
		return valuesByRowNum(cellReader -> cellReader.stringValue(formatter));
	}

	public Map<Integer, Boolean> booleanValuesByRowNum() {
		return valuesByRowNum(PoiCellReader::booleanValue);
	}

	public Map<Integer, LocalDateTime> dateTimeValuesByRowNum() {
		return valuesByRowNum(PoiCellReader::dateTimeValue);
	}

	public Map<Integer, BigDecimal> numericValuesByRowNum() {
		return valuesByRowNum(PoiCellReader::numericValue);
	}
}
