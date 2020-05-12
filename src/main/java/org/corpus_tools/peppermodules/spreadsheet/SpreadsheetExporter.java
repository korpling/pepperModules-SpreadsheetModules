package org.corpus_tools.peppermodules.spreadsheet;

import org.corpus_tools.pepper.impl.PepperExporterImpl;
import org.corpus_tools.pepper.modules.PepperExporter;
import org.osgi.service.component.annotations.Component;

@Component(name="SpreadsheetExporterComponent", factory="PepperExporterComponentFactory")
public class SpreadsheetExporter extends PepperExporterImpl implements PepperExporter {

}
