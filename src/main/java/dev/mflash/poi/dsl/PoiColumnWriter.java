package dev.mflash.poi.dsl;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;

import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;

public final class PoiColumnWriter {

	private final Sheet sheet;
	private final CellReference[] cellReferences;

	public PoiColumnWriter(Sheet sheet, AreaReference areaReference) {
		this.sheet = sheet;
		this.cellReferences = areaReference.getAllReferencedCells();
	}

	public void stringValues(List<String> values) {
		for (int i = 0; i < Math.min(values.size(), cellReferences.length); i++) {
			final String value = values.get(i);
			if (Objects.nonNull(value)) {
				final PoiCellWriter writer = new PoiCellWriter(sheet, cellReferences[i]);
				writer.stringValue(value);
			}
		}
	}

	public void stringValues(Map<Integer, String> values) {
		final Map<Integer, CellReference> cellReferencesByRowIndex = cellReferencesByRowIndex();
		values.forEach((rowIndex, value) -> {
			if (cellReferencesByRowIndex.containsKey(rowIndex) && Objects.nonNull(value)) {
				final PoiCellWriter writer = new PoiCellWriter(sheet, cellReferencesByRowIndex.get(rowIndex));
				writer.stringValue(value);
			}
		});
	}

	public void booleanValues(List<Boolean> values) {
		for (int i = 0; i < Math.min(values.size(), cellReferences.length); i++) {
			final Boolean value = values.get(i);
			if (Objects.nonNull(value)) {
				final PoiCellWriter writer = new PoiCellWriter(sheet, cellReferences[i]);
				writer.booleanValue(value);
			}
		}
	}

	public void booleanValues(Map<Integer, Boolean> values) {
		final Map<Integer, CellReference> cellReferencesByRowIndex = cellReferencesByRowIndex();
		values.forEach((rowIndex, value) -> {
			if (cellReferencesByRowIndex.containsKey(rowIndex) && Objects.nonNull(value)) {
				final PoiCellWriter writer = new PoiCellWriter(sheet, cellReferencesByRowIndex.get(rowIndex));
				writer.booleanValue(value);
			}
		});
	}

	public void dateTimeValues(List<LocalDateTime> values) {
		for (int i = 0; i < Math.min(values.size(), cellReferences.length); i++) {
			final LocalDateTime value = values.get(i);
			if (Objects.nonNull(value)) {
				final PoiCellWriter writer = new PoiCellWriter(sheet, cellReferences[i]);
				writer.dateTimeValue(value);
			}
		}
	}

	public void dateTimeValues(Map<Integer, LocalDateTime> values) {
		final Map<Integer, CellReference> cellReferencesByRowIndex = cellReferencesByRowIndex();
		values.forEach((rowIndex, value) -> {
			if (cellReferencesByRowIndex.containsKey(rowIndex) && Objects.nonNull(value)) {
				final PoiCellWriter writer = new PoiCellWriter(sheet, cellReferencesByRowIndex.get(rowIndex));
				writer.dateTimeValue(value);
			}
		});
	}

	public void integerValues(List<Integer> values) {
		for (int i = 0; i < Math.min(values.size(), cellReferences.length); i++) {
			final Integer value = values.get(i);
			if (Objects.nonNull(value)) {
				final PoiCellWriter writer = new PoiCellWriter(sheet, cellReferences[i]);
				writer.integerValue(value);
			}
		}
	}

	public void integerValues(Map<Integer, Integer> values) {
		final Map<Integer, CellReference> cellReferencesByRowIndex = cellReferencesByRowIndex();
		values.forEach((rowIndex, value) -> {
			if (cellReferencesByRowIndex.containsKey(rowIndex) && Objects.nonNull(value)) {
				final PoiCellWriter writer = new PoiCellWriter(sheet, cellReferencesByRowIndex.get(rowIndex));
				writer.integerValue(value);
			}
		});
	}

	public void longValues(List<Long> values) {
		for (int i = 0; i < Math.min(values.size(), cellReferences.length); i++) {
			final Long value = values.get(i);
			if (Objects.nonNull(value)) {
				final PoiCellWriter writer = new PoiCellWriter(sheet, cellReferences[i]);
				writer.longValue(value);
			}
		}
	}

	public void longValues(Map<Integer, Long> values) {
		final Map<Integer, CellReference> cellReferencesByRowIndex = cellReferencesByRowIndex();
		values.forEach((rowIndex, value) -> {
			if (cellReferencesByRowIndex.containsKey(rowIndex) && Objects.nonNull(value)) {
				final PoiCellWriter writer = new PoiCellWriter(sheet, cellReferencesByRowIndex.get(rowIndex));
				writer.longValue(value);
			}
		});
	}

	public void doubleValues(List<Double> values) {
		for (int i = 0; i < Math.min(values.size(), cellReferences.length); i++) {
			final Double value = values.get(i);
			if (Objects.nonNull(value)) {
				final PoiCellWriter writer = new PoiCellWriter(sheet, cellReferences[i]);
				writer.doubleValue(value);
			}
		}
	}

	public void doubleValues(Map<Integer, Double> values) {
		final Map<Integer, CellReference> cellReferencesByRowIndex = cellReferencesByRowIndex();
		values.forEach((rowIndex, value) -> {
			if (cellReferencesByRowIndex.containsKey(rowIndex) && Objects.nonNull(value)) {
				final PoiCellWriter writer = new PoiCellWriter(sheet, cellReferencesByRowIndex.get(rowIndex));
				writer.doubleValue(value);
			}
		});
	}

	private Map<Integer, CellReference> cellReferencesByRowIndex() {
		return Arrays.stream(cellReferences)
				.collect(Collectors.toMap(CellReference::getRow, cellReference -> cellReference));
	}
}
