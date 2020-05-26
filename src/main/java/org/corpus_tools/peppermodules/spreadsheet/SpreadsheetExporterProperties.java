package org.corpus_tools.peppermodules.spreadsheet;

import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.corpus_tools.pepper.modules.PepperModuleProperties;
import org.corpus_tools.pepper.modules.PepperModuleProperty;

public class SpreadsheetExporterProperties extends PepperModuleProperties {
	/** Sets the font is to be used in the spreadsheet table (this influences the covered unicode characters). */
	private static final String PROP_FONT_NAME = "font.name";
	/** Configure the order of columns as comma-separated annotation names (qualified). Partial configurations are allowed. If no configuration is provided for an annotation, the next free column is used. */
	private static final String PROP_COL_ORDER = "column.order";
	
	public SpreadsheetExporterProperties() {
		addProperty(PepperModuleProperty.create()
				.withName(PROP_FONT_NAME)
				.withType(String.class)
				.withDescription("Sets the font is to be used in the spreadsheet table (this influences the covered unicode characters).")
				.withDefaultValue("DejaVu Sans").build());
		addProperty(PepperModuleProperty.create()
				.withName(PROP_COL_ORDER)
				.withType(String.class)
				.withDescription("Configure the order of columns as comma-separated annotation names (qualified). Partial configurations are allowed. If no configuration is provided for an annotation, the next free column is used.")
				.isRequired(false)
				.build());
	}
	
	/**
	 * 
	 * @return configured font name as {@link String}.
	 */
	public String getFont() {
		return (String) getProperty(PROP_FONT_NAME).getValue();
	}
	
	public Map<String, Integer> getColumnOrder() {
		Object val = getProperty(PROP_COL_ORDER).getValue();
		if (val == null) {
			return Collections.<String, Integer>emptyMap();
		}
		Map<String, Integer> columnOrder = new HashMap<>();
		for (String qName : StringUtils.split((String) val)) {
			columnOrder.put(qName, columnOrder.values().size());
		}
		return columnOrder;
	}
}
