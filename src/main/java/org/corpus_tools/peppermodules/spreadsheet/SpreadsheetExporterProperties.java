/**
 * Copyright 2009 Humboldt-Universität zu Berlin, INRIA.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *       http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 *
 */
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
	/** TODO This property defines the preferred vertical alignment. */
	public static final String PROP_VERTICAL_ALIGNMENT = "text.align.vertical";
	/** TODO This property defines the preferred horizontal alignment. */
	public static final String PROP_HORIZONTAL_ALIGNMENT = "text.align.horizontal";
	/** Creates a column containing the document name if a String header value is provided. **/
	public static final String PROP_DOC_COL_TITLE = "document.column.title";
	
	public SpreadsheetExporterProperties() {
		addProperty(PepperModuleProperty.create()
				.withName(PROP_FONT_NAME)
				.withType(String.class)
				.withDescription("Sets the font is to be used in the spreadsheet table (this influences the covered unicode characters).")
				.isRequired(false)
				.build());
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
		addProperty(PepperModuleProperty.create()
				.withName(PROP_VERTICAL_ALIGNMENT)
				.withType(AlignmentValue.class)
				.withDescription("")
				.withDefaultValue(AlignmentValue.top)
				.build());
		addProperty(PepperModuleProperty.create()
				.withName(PROP_HORIZONTAL_ALIGNMENT)
				.withType(AlignmentValue.class)
				.withDescription("")
				.withDefaultValue(AlignmentValue.left)
				.build());
		addProperty(PepperModuleProperty.create()
				.withName(PROP_DOC_COL_TITLE)
				.withType(String.class)
				.withDescription("Creates a column containing the document name if a String header value is provided.")
				.build());
	}
	
	/**
	 * 
	 * @return configured font name as {@link String}.
	 */
	public String getFont() {
		Object value = getProperty(PROP_FONT_NAME).getValue();
		return value == null? null : (String) value;
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
	
	public AlignmentValue getVerticalTextAlignment() {
		return (AlignmentValue) getProperty(PROP_VERTICAL_ALIGNMENT).getValue();
	}

	public AlignmentValue getHorizontalTextAlignment() {
		return (AlignmentValue) getProperty(PROP_HORIZONTAL_ALIGNMENT).getValue();
	}
	
	public static enum AlignmentValue {
		bottom, mid, top, left, center, right;
	}
	
	public String getDocumentColumnTitle() {
		Object value = getProperty(PROP_DOC_COL_TITLE).getValue();
		return value == null? null : (String) value;
	}
}
