package org.eugene.any2word;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

public class CustomXmlUtils {
	/**
	 * We need the item id of the custom xml part.
	 *
	 * There are several strategies we could use to find it, including searching
	 * the docx for a bind element, but here, we simply look in the xpaths part.
	 *
	 * @param wordMLPackage
	 * @return
	 */
	public static String getCustomXmlItemId(
			WordprocessingMLPackage wordMLPackage) throws Docx4JException {
		MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
		if (documentPart.getXPathsPart() == null) {
			throw new Docx4JException("OpenDoPE XPaths part missing");
		}
		org.opendope.xpaths.Xpaths xPaths = documentPart
				.getXPathsPart().getJaxbElement();
		return xPaths.getXpath().get(0).getDataBinding().getStoreItemID();
	}

}
