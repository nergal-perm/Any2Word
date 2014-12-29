package org.eugene.any2word;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

public class CoreWordFiller {
	private WordprocessingMLPackage wordMLPackage;

	public CoreWordFiller(WordprocessingMLPackage wordDoc) {
		setWordDocument(wordDoc);
	}

	public WordprocessingMLPackage getWordDocument() {
		return wordMLPackage;
	}

	public void setWordDocument(WordprocessingMLPackage wordMLPackage) {
		this.wordMLPackage = wordMLPackage;
	}
	
	

}