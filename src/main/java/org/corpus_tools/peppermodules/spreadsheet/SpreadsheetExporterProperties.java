package org.corpus_tools.peppermodules.spreadsheet;

import org.corpus_tools.pepper.modules.PepperModuleProperties;
import org.corpus_tools.pepper.modules.PepperModuleProperty;

public class SpreadsheetExporterProperties extends PepperModuleProperties {
	/** Sets the font is to be used in the spreadsheet table (this influences the covered unicode characters). */
	private static final String PROP_FONT_NAME = "font.name";
	
	public SpreadsheetExporterProperties() {
		addProperty(PepperModuleProperty.create()
				.withName(PROP_FONT_NAME)
				.withType(String.class)
				.withDescription("Sets the font is to be used in the spreadsheet table (this influences the covered unicode characters).")
				.withDefaultValue("DejaVu Sans").build());
	}
	
	/**
	 * 
	 * @return configured font name as {@link String}.
	 */
	public String getFont() {
		return (String) getProperty(PROP_FONT_NAME).getValue();
	}
}
