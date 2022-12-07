package dev.mflash.poi.dsl;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;

import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.stream.Stream;

public final class PoiColumnReader {

	private final Sheet sheet;
	private final CellReference[] cellReferences;

	public PoiColumnReader(Sheet sheet, AreaReference areaReference) {
		this.sheet = sheet;
		this.cellReferences = areaReference.getAllReferencedCells();
	}

	public Stream<String> stringValues() {
		return readers().map(PoiCellReader::stringValue);
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

	private Stream<PoiCellReader> readers() {
		return Arrays.stream(cellReferences)
				.map(cellReference -> new PoiCellReader(sheet, cellReference));
	}
}
