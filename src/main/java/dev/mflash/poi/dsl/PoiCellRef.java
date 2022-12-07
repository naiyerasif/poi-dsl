package dev.mflash.poi.dsl;

import org.apache.poi.ss.util.CellReference;

public record PoiCellRef(CellReference cellReference) {

	public PoiCellRef(PoiCellRefBuilder builder) {
		this(builder.cellReference);
	}

	public int columnIndex() {
		return cellReference.getCol();
	}

	public int rowIndex() {
		return cellReference.getRow();
	}

	public PoiCellRef withColumnIndex(int columnIndex) {
		return builder().reference(cellReference.getRow(), columnIndex).build();
	}

	public PoiCellRef withRowIndex(int rowIndex) {
		return builder().reference(rowIndex, cellReference.getCol()).build();
	}

	public static PoiCellRefBuilder builder() {
		return new PoiCellRefBuilder();
	}

	public static final class PoiCellRefBuilder {

		private CellReference cellReference = new CellReference(0, 0);

		private PoiCellRefBuilder() {}

		public PoiCellRefBuilder reference(int rowIndex, int columnIndex) {
			this.cellReference = new CellReference(rowIndex, columnIndex);
			return this;
		}

		public PoiCellRefBuilder reference(String reference) {
			this.cellReference = new CellReference(reference);
			return this;
		}

		public PoiCellRef build() {
			return new PoiCellRef(this);
		}
	}
}
