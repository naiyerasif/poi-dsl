package dev.mflash.poi.dsl.value;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;

public record PoiCellAttributes(
		CellReference cellReference,
		CellType cellType,
		Row.MissingCellPolicy missingCellPolicy
) {

	PoiCellAttributes(PoiCellAttributesBuilder builder) {
		this(builder.cellReference, builder.cellType, builder.missingCellPolicy);
	}

	public int rowIndex() {
		return this.cellReference.getRow();
	}

	public int columnIndex() {
		return this.cellReference.getCol();
	}

	public static PoiCellAttributesBuilder builder() {
		return new PoiCellAttributesBuilder();
	}

	public static final class PoiCellAttributesBuilder {
		private CellReference cellReference = new CellReference(0, 0);
		private CellType cellType = CellType.BLANK;
		private Row.MissingCellPolicy missingCellPolicy;

		private PoiCellAttributesBuilder() {}

		public PoiCellAttributesBuilder reference(CellReference cellReference) {
			this.cellReference = cellReference;
			return this;
		}

		public PoiCellAttributesBuilder reference(String cellReference) {
			return reference(new CellReference(cellReference));
		}

		public PoiCellAttributesBuilder index(int rowIndex, int columnIndex) {
			return reference(new CellReference(rowIndex, columnIndex));
		}

		public PoiCellAttributesBuilder cellType(CellType cellType) {
			this.cellType = cellType;
			return this;
		}

		public PoiCellAttributesBuilder missingCellPolicy(Row.MissingCellPolicy missingCellPolicy) {
			this.missingCellPolicy = missingCellPolicy;
			return this;
		}

		public PoiCellAttributes build() {
			return new PoiCellAttributes(this);
		}
	}
}
