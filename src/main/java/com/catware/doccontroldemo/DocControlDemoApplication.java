package com.catware.doccontroldemo;

import org.docx4j.Docx4J;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.SdtRun;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import javax.xml.bind.JAXBContext;
import javax.xml.bind.JAXBElement;
import java.io.File;
import java.io.FileInputStream;
import java.util.List;

@SpringBootApplication
public class DocControlDemoApplication implements CommandLineRunner {

    public static JAXBContext context = org.docx4j.jaxb.Context.jc;

    private final static boolean DEBUG = true;
    private final static boolean SAVE = true;

    public static void main(String[] args) {
        SpringApplication.run(DocControlDemoApplication.class, args);
    }

    @Override
    public void run(String... args) throws Exception {
        // Without Saxon, you are restricted to XPath 1.0
        boolean USE_SAXON = true; // set this to true; add Saxon to your classpath, and uncomment below
//		if (USE_SAXON) XPathFactoryUtil.setxPathFactory(
//				new net.sf.saxon.xpath.XPathFactoryImpl());

        // the docx 'template'
//        String input_DOCX = "binding-simple.docx";
//        String input_DOCX = "Brev.docx";
        String input_DOCX = "OUT_BrevXML.docx";

        // the instance data
//        String input_XML = "binding-simple-data.xml";

        // resulting docx
//        String OUTPUT_DOCX = "OUT_ContentControlsMergeXML.docx";
        String OUTPUT_DOCX = "OUT_BrevXML.docx";

        // Load input_template.docx
        WordprocessingMLPackage wordMLPackage = Docx4J.load(new File(input_DOCX));

        // Open the xml stream
//        FileInputStream xmlStream = new FileInputStream(new File(input_XML));

        List<Object> objVars = wordMLPackage.getMainDocumentPart().getJAXBNodesViaXPath("//w:sdt", false);
        for (Object objVar : objVars) {
            JAXBElement jaxbVar = (JAXBElement) objVar;
            SdtRun var = (SdtRun) jaxbVar.getValue();
            if ("FulltNavn".equals(var.getSdtPr().getTag().getVal())) {
                ((javax.xml.bind.JAXBElement) ((org.docx4j.wml.R) var.getSdtContent().getContent().get(0)).getContent().get(0)).setValue("HELLO THOMAS!");
//            System.out.println(objVar.getClass());
            }
        }

        // Do the binding:
        // FLAG_NONE means that all the steps of the binding will be done,
        // otherwise you could pass a combination of the following flags:
        // FLAG_BIND_INSERT_XML: inject the passed XML into the document
        // FLAG_BIND_BIND_XML: bind the document and the xml (including any OpenDope handling)
        // FLAG_BIND_REMOVE_SDT: remove the content controls from the document (only the content remains)
        // FLAG_BIND_REMOVE_XML: remove the custom xml parts from the document

        //Docx4J.bind(wordMLPackage, xmlStream, Docx4J.FLAG_NONE);
        //If a document doesn't include the Opendope definitions, eg. the XPathPart,
        //then the only thing you can do is insert the xml
        //the example document binding-simple.docx doesn't have an XPathPart....

//        Docx4J.bind(wordMLPackage, xmlStream, Docx4J.FLAG_BIND_INSERT_XML | Docx4J.FLAG_BIND_BIND_XML);

        //Save the document
        Docx4J.save(wordMLPackage, new File(OUTPUT_DOCX), Docx4J.FLAG_NONE);
//        System.out.println("Saved: " + OUTPUT_DOCX);
    }
}
