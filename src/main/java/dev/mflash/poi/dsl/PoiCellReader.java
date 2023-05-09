package dev.mflash.poi.dsl;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;

import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.function.Supplier;

public class PoiCellReader {

	private static final DataFormatter _DATA_FORMATTER = new DataFormatter();

	private final Cell cell;

	PoiCellReader(Cell cell) {
		this.cell = cell;
	}

	public static PoiCellReader create(Supplier<Cell> cellSupplier) {
		return new PoiCellReader(cellSupplier.get());
	}

	public String stringValue() {
		return cell.getCellType().equals(CellType.STRING) ?
				cell.getStringCellValue() : _DATA_FORMATTER.formatCellValue(cell);
	}

	public String stringValue(DateTimeFormatter dateTimeFormatter) {
		return cell.getCellType().equals(CellType.NUMERIC) && DateUtil.isCellDateFormatted(cell) ?
				dateTimeValue().format(dateTimeFormatter) : stringValue();
	}

	public boolean booleanValue() {
		try {
			return cell.getBooleanCellValue();
		} catch (Exception __) {
			return Boolean.parseBoolean(stringValue().strip());
		}
	}

	public LocalDateTime dateTimeValue() {
		return cell.getLocalDateTimeCellValue();
	}

	public BigDecimal numericValue() {
		try {
			return BigDecimal.valueOf(cell.getNumericCellValue());
		} catch (Exception __) {
			return new BigDecimal(stringValue().strip());
		}
	}
}
