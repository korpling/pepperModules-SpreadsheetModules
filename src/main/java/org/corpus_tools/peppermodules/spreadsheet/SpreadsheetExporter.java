package org.corpus_tools.peppermodules.spreadsheet;

import org.corpus_tools.pepper.common.PepperConfiguration;
import org.corpus_tools.pepper.impl.PepperExporterImpl;
import org.corpus_tools.pepper.modules.PepperExporter;
import org.corpus_tools.pepper.modules.PepperMapper;
import org.corpus_tools.salt.common.SDocument;
import org.corpus_tools.salt.graph.Identifier;
import org.eclipse.emf.common.util.URI;
import org.osgi.service.component.annotations.Component;

@Component(name="SpreadsheetExporterComponent", factory="PepperExporterComponentFactory")
public class SpreadsheetExporter extends PepperExporterImpl implements PepperExporter {
	
	private static final String NAME = "SpreadsheetExporter";

	public SpreadsheetExporter() {
		super();
		setSupplierContact(URI.createFileURI(PepperConfiguration.EMAIL));
		setName(NAME);
		addSupportedFormat("xlsx", "2007+", null);
		setIsMultithreaded(true);
		setProperties(new SpreadsheetExporterProperties());
		setDesc("This exporter transforms a Salt model into a spreadsheet.");
		setDocumentEnding("xlsx");
		setExportMode(EXPORT_MODE.DOCUMENTS_IN_FILES);
	}
	
	@Override
	public PepperMapper createPepperMapper(Identifier identifier) {
		Salt2SpreadsheetMapper mapper = new Salt2SpreadsheetMapper();
		if (identifier.getIdentifiableElement() != null && identifier.getIdentifiableElement() instanceof SDocument) {
			URI resource = getIdentifier2ResourceTable().get(identifier);
			mapper.setResourceURI(resource);
		}
		return mapper;
	}
}
