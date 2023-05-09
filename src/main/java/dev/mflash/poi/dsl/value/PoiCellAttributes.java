package dev.mflash.poi.dsl.value;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;

public record PoiCellAttributes(
		CellReference cellReference,
		CellType cellType,
		CellStyle cellStyle,
		Hyperlink hyperlink,
		Comment comment,
		Row.MissingCellPolicy missingCellPolicy
) {

	PoiCellAttributes(PoiCellAttributesBuilder builder) {
		this(
				builder.cellReference,
				builder.cellType,
				builder.cellStyle,
				builder.hyperlink,
				builder.comment,
				builder.missingCellPolicy
		);
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
		private CellStyle cellStyle;
		private Hyperlink hyperlink;
		private Comment comment;
		private Row.MissingCellPolicy missingCellPolicy;

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

		public PoiCellAttributesBuilder cellStyle(CellStyle cellStyle) {
			this.cellStyle = cellStyle;
			return this;
		}

		public PoiCellAttributesBuilder hyperlink(Hyperlink hyperlink) {
			this.hyperlink = hyperlink;
			return this;
		}

		public PoiCellAttributesBuilder comment(Comment comment) {
			this.comment = comment;
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
