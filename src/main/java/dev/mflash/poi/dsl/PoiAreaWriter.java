package dev.mflash.poi.dsl;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;

import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

public final class PoiAreaWriter {

	private final Map<Integer, Row> rowsByRowIndex;

	public PoiAreaWriter(Sheet sheet, AreaReference areaReference) {
		this.rowsByRowIndex = Arrays.stream(areaReference.getAllReferencedCells())
				.map(CellReference::getRow).distinct()
				.collect(Collectors.toMap(rowIndex -> rowIndex, sheet::getRow));
	}

	public <T> void write(List<T> data, PoiTypeMapper<Row, T> mapper) {
		// TODO: implement
	}

	public <T> void write(Map<Integer, T> data, PoiTypeMapper<Row, T> mapper) {
		// TODO: implement
	}

	public <T> void rewrite(List<T> data, PoiTypeMapper<Row, T> mapper) {
		// TODO: implement
	}

	public <T> void rewrite(Map<Integer, T> data, PoiTypeMapper<Row, T> mapper) {
		// TODO: implement
	}

	private void purge() {
		// TODO: implement
	}
}
