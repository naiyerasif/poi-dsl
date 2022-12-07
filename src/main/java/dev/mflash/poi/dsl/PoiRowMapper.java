package dev.mflash.poi.dsl;

import org.apache.poi.ss.usermodel.Row;

@FunctionalInterface
public interface PoiRowMapper<S, R extends Row> {

	S mapRow(Row row);
}
