package dev.mflash.poi.dsl;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;

import java.util.Arrays;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public final class PoiAreaReader {

	private final Map<Integer, Row> rowsByRowIndex;

	public PoiAreaReader(Sheet sheet, AreaReference areaReference) {
		this.rowsByRowIndex = Arrays.stream(areaReference.getAllReferencedCells())
				.map(CellReference::getRow).distinct()
				.collect(Collectors.toMap(rowIndex -> rowIndex, sheet::getRow));
	}

	public <T> Stream<T> values(PoiRowMapper<T, Row> mapper) {
		return this.rowsByRowIndex.values().stream().map(mapper::mapRow);
	}
}
