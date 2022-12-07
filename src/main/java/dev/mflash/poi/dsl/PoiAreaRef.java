package dev.mflash.poi.dsl;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;

public record PoiAreaRef(AreaReference areaReference) {

	private static final SpreadsheetVersion _spreadSheetVersion = SpreadsheetVersion.EXCEL2007;

	public PoiAreaRef(PoiAreaRefBuilder builder) {
		this(builder.areaReference);
	}

	public AreaReference columnReference() {
		if (firstReference().getCol() != lastReference().getCol()) {
			throw new IllegalReferenceException("not a column reference");
		}
		return areaReference;
	}

	public AreaReference columnReference(int columnIndex) {
		if (columnIndex < firstReference().getCol() || columnIndex > lastReference().getCol()) {
			throw new IllegalReferenceException("columnIndex '" + columnIndex + "' out of bound");
		}
		var start = new CellReference(firstReference().getRow(), columnIndex);
		var stop = new CellReference(lastReference().getRow(), columnIndex);
		return new AreaReference(start, stop, _spreadSheetVersion);
	}

	public CellReference firstReference() {
		return areaReference.getFirstCell();
	}

	public CellReference lastReference() {
		return areaReference.getLastCell();
	}

	public static PoiAreaRefBuilder builder() {
		return new PoiAreaRefBuilder();
	}

	public static final class PoiAreaRefBuilder {

		private AreaReference areaReference;

		private PoiAreaRefBuilder() {}

		public PoiAreaRefBuilder reference(CellReference start, CellReference stop) {
			this.areaReference = new AreaReference(start, stop, _spreadSheetVersion);
			return this;
		}

		public PoiAreaRefBuilder reference(String reference) {
			this.areaReference = new AreaReference(reference, _spreadSheetVersion);
			return this;
		}

		public PoiAreaRef build() {
			return new PoiAreaRef(this);
		}
	}
}
