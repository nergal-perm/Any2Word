package org.eugene.any2word;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;

import org.docx4j.XmlUtils;
import org.docx4j.customXmlProperties.DatastoreItem;
import org.docx4j.model.datastorage.CustomXmlDataStorage;
import org.docx4j.model.datastorage.CustomXmlDataStorageImpl;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.CustomXmlDataStoragePart;
import org.docx4j.openpackaging.parts.CustomXmlDataStoragePropertiesPart;
import org.docx4j.openpackaging.parts.Parts;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

public class CoreWordFiller {
	private WordprocessingMLPackage wordMLPackage;

	/**
	 * Creates in-memory representation of a Word template, adds provided
	 * XML file to it, manages data binding to ContentControls inside the
	 * template
	 * 
	 * @param wordDoc - Word template file
	 * @param xmlFile - XML file with custom data to use with Word template
	 * @throws FileNotFoundException
	 * @throws Docx4JException
	 */
	public CoreWordFiller(WordprocessingMLPackage wordDoc, File xmlFile)
			throws FileNotFoundException, Docx4JException {
		setWordDocument(wordDoc);
		
		// Create custom XML part infrastructure
		CustomXmlDataStoragePart xmlPart = injectCustomXmlDataStoragePart(
				wordDoc.getMainDocumentPart(), wordDoc.getParts(), xmlFile);
		if (xmlPart == null) {
			throw new Docx4JException(
					"Couldn't find Custom XML Part in a document");
		}
		
		// Populate custom XML part with data provided
		// xmlPart.getData().setDocument(new FileInputStream(xmlFile));
		addProperties(xmlPart);
		CustomXmlDataStorage customXmlDataStorage = xmlPart.getData();
		// DEBUG only
		System.out.println(XmlUtils.w3CDomNodeToString(customXmlDataStorage.getDocument()));		
		
		// Bind ContentControls to data in XML file
	}

	/**
	 * 
	 * @param base
	 * @param parts
	 * @return
	 */
	private CustomXmlDataStoragePart injectCustomXmlDataStoragePart(
			MainDocumentPart base, Parts parts, File xmlFile) {

		CustomXmlDataStoragePart customXmlDataStoragePart = null;
		try {

			customXmlDataStoragePart = new CustomXmlDataStoragePart(parts);
			CustomXmlDataStorage data = new CustomXmlDataStorageImpl();
			data.setDocument(createCustomXmlDocument(xmlFile));
			customXmlDataStoragePart.setData(data);
			base.addTargetPart(customXmlDataStoragePart);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return customXmlDataStoragePart;
	}

	private InputStream createCustomXmlDocument(File xmlFile) {
		try {
			return new FileInputStream(xmlFile);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			return null;
		}
	}
	
    public static void addProperties(CustomXmlDataStoragePart customXmlDataStoragePart) throws InvalidFormatException {
        CustomXmlDataStoragePropertiesPart part = new CustomXmlDataStoragePropertiesPart();
        org.docx4j.customXmlProperties.ObjectFactory of = new org.docx4j.customXmlProperties.ObjectFactory();
        DatastoreItem dsi = of.createDatastoreItem();
        String newItemId = "{DD6E220C-54BC-47B3-8AE8-A0A61D4934FF}".toLowerCase();                              
        dsi.setItemID(newItemId);
        part.setJaxbElement(dsi );
        customXmlDataStoragePart.addTargetPart(part);
    }



	public WordprocessingMLPackage getWordDocument() {
		return wordMLPackage;
	}

	public void setWordDocument(WordprocessingMLPackage wordMLPackage) {
		this.wordMLPackage = wordMLPackage;
	}

}