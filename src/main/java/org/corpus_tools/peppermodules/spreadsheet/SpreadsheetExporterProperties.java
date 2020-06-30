package org.corpus_tools.peppermodules.spreadsheet;

import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

import org.apache.commons.lang3.StringUtils;
import org.corpus_tools.pepper.modules.PepperModuleProperties;
import org.corpus_tools.pepper.modules.PepperModuleProperty;

public class SpreadsheetExporterProperties extends PepperModuleProperties {
	/** Sets the font is to be used in the spreadsheet table (this influences the covered unicode characters). */
	public static final String PROP_FONT_NAME = "font.name";
	/** Configure the order of columns as comma-separated annotation names (qualified). Partial configurations are allowed. If no configuration is provided for an annotation, the next free column is used. */
	public static final String PROP_COL_ORDER = "column.order";
	/** Remove trailing and leading whitespaces. */
	public static final String PROP_TRIM_VALUES = "trim.values";
	/** Annotation names (comma-separated), that are ignored when mapping span annotations. The provided names will still be imported as tokens, given there are order relations. */
	public static final String PROP_IGNORE_ANNO_NAMES = "ignore.anno.names";
	
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
		addProperty(PepperModuleProperty.create()
				.withName(PROP_TRIM_VALUES)
				.withType(Boolean.class)
				.withDescription("Remove trailing and leading whitespaces.")
				.withDefaultValue(false)
				.build());
		addProperty(PepperModuleProperty.create()
				.withName(PROP_IGNORE_ANNO_NAMES)
				.withType(String.class)
				.withDescription("Annotation names (comma-separated), that are ignored when mapping span annotations. The provided names will still be imported as tokens, given there are order relations.")
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
		for (String qName : StringUtils.split((String) val, ',')) {
			columnOrder.put(qName.trim(), columnOrder.values().size());
		}
		return columnOrder;
	}
	
	public boolean trimValues() {
		return (Boolean) getProperty(PROP_TRIM_VALUES).getValue();
	}
	
	public Set<String> ignoreAnnoNames() {
		Object o = getProperty(PROP_IGNORE_ANNO_NAMES).getValue();
		if (o == null) {
			return Collections.<String>emptySet();
		}
		return Arrays.stream(StringUtils.split((String) o, ","))
				.map((String s) -> s.trim())
				.collect(Collectors.toSet());
	}
}
