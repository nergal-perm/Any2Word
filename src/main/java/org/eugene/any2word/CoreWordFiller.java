package org.eugene.any2word;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.docx4j.Docx4J;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.CustomXmlDataStoragePart;
import org.docx4j.samples.CreateDocx;

public class CoreWordFiller {
	private WordprocessingMLPackage wordMLPackage;

	public CoreWordFiller(WordprocessingMLPackage wordDoc, File xmlFile) throws FileNotFoundException, Docx4JException {
		setWordDocument(wordDoc);
		Docx4J.bind(wordMLPackage, new FileInputStream(xmlFile), Docx4J.FLAG_BIND_INSERT_XML);
		// add XML file as part of the package
		String itemID = CustomXmlUtils.getCustomXmlItemId(wordMLPackage);
		CustomXmlDataStoragePart xmlPart = (CustomXmlDataStoragePart) wordMLPackage.getCustomXmlDataStorageParts().get(itemID);
		if (xmlPart==null) {
			throw new Docx4JException("Couldn't find Custom XML Part in a document");
		}
		xmlPart.getData().setDocument(new FileInputStream(xmlFile));
		
		
		
	}

	public WordprocessingMLPackage getWordDocument() {
		return wordMLPackage;
	}

	public void setWordDocument(WordprocessingMLPackage wordMLPackage) {
		this.wordMLPackage = wordMLPackage;
	}
	
	

}