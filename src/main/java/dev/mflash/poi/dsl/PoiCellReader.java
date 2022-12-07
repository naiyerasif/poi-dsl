package dev.mflash.poi.dsl;

import com.github.sisyphsu.dateparser.DateParser;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;

import java.math.BigDecimal;
import java.time.LocalDateTime;

public final class PoiCellReader {

	private static final DataFormatter _dataFormatter = new DataFormatter();
	private static final DateParser _dateParser = DateParser.newBuilder().build();

	private final Cell cell;

	public PoiCellReader(Sheet sheet, CellReference cellReference) {
		this.cell = sheet.getRow(cellReference.getRow()).getCell(cellReference.getCol());
	}

	public Cell cell() {
		return cell;
	}

	public String stringValue() {
		return cell.getCellType().equals(CellType.STRING) ?
				cell.getStringCellValue() :
				_dataFormatter.formatCellValue(cell);
	}

	public boolean booleanValue() {
		try {
			return cell.getBooleanCellValue();
		} catch (Exception __) {
			return Boolean.parseBoolean(stringValue().strip());
		}
	}

	public LocalDateTime dateTimeValue() {
		try {
			return cell.getLocalDateTimeCellValue();
		} catch (Exception __) {
			return _dateParser.parseDateTime(stringValue().strip());
		}
	}

	public BigDecimal numericValue() {
		try {
			return BigDecimal.valueOf(cell.getNumericCellValue());
		} catch (Exception __) {
			return new BigDecimal(stringValue().strip());
		}
	}
}
