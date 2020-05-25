package org.corpus_tools.peppermodules.spreadsheet;

import org.corpus_tools.pepper.common.DOCUMENT_STATUS;
import org.corpus_tools.pepper.impl.PepperMapperImpl;
import org.corpus_tools.pepper.modules.PepperMapper;

public class Salt2SpreadsheetMapper extends PepperMapperImpl implements PepperMapper {
	@Override
	public DOCUMENT_STATUS mapSDocument() {
		return DOCUMENT_STATUS.COMPLETED;
	}
}
