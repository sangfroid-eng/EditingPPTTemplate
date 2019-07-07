mport java.awt.Dimension;
import java.awt.Toolkit;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;

import javax.swing.Icon;
import javax.swing.ImageIcon;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

class CreatePDF {

	public static void editPPT() throws IOException, SAXException, ParserConfigurationException {
		File fXmlFile = new File("C:\\Desktop\\aspose\\ppt_to_video\\CustomerDetails.xml");
		DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
		DocumentBuilder dbuilder = dbFactory.newDocumentBuilder();
		Document doc = dbuilder.parse(fXmlFile);
		doc.getDocumentElement().normalize();

		NodeList nList = doc.getElementsByTagName("customer");

		for (int temp = 0; temp < nList.getLength(); temp++) {
			Node nNode = nList.item(temp);
			if (nNode.getNodeType() == Node.ELEMENT_NODE) {
				Element eElement = (Element) nNode;
				File file = new File("C:\\Desktop\\aspose\\ppt_to_video\\a.pptx");
				XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(file));

				XSLFSlide[] slide = ppt.getSlides();
				for (int i = 0; i < slide.length; i++)

				{
					XSLFTextShape[] sh = slide[i].getPlaceholders();
					for (int j = 0; j < sh.length; j++) {
						if (sh[j].getText().equals("PH_Name")) {
							sh[j].setText("Tina Rawat"/*
															 * eElement.getElementsByTagName("firstname").item(0).
															 * getTextContent() + " " +
															 * eElement.getElementsByTagName("lastname").item(0).
															 * getTextContent()
															 */);
						}
						if (sh[j].getText().equals("PH_Acc")) {
							sh[j].setText(eElement.getElementsByTagName("accountNo").item(0).getTextContent());
						}
						if (sh[j].getText().equals("PH_StartDate") || sh[j].getText().equals("PH_FromDT")) {
							sh[j].setText(eElement.getElementsByTagName("startingFrom").item(0).getTextContent());
						}
						if (sh[j].getText().equals("PH_EndDate") || sh[j].getText().equals("PH_ToDt")) {
							sh[j].setText(eElement.getElementsByTagName("startingTo").item(0).getTextContent());
						}
						if (sh[j].getText().equals("PH_Place")) {
							sh[j].setText(eElement.getElementsByTagName("place").item(0).getTextContent());
						}
						if (sh[j].getText().equals("PH_Phone")) {
							sh[j].setText(eElement.getElementsByTagName("phone").item(0).getTextContent());
						}
						if (sh[j].getText().equals("PH_Email")) {
							sh[j].setText(eElement.getElementsByTagName("email_id").item(0).getTextContent());
						}
						if (sh[j].getText().equals("PH_Expenditure")) {
							sh[j].setText(eElement.getElementsByTagName("totalExpenditure").item(0).getTextContent());
						}
						if (sh[j].getText().equals("PH_OpeningBal")) {
							sh[j].setText(eElement.getElementsByTagName("openingBalance").item(0).getTextContent());
						}
						if (sh[j].getText().equals("PH_ClosingBal")) {
							sh[j].setText(eElement.getElementsByTagName("closingBalance").item(0).getTextContent());
						}
					}
				}
				File outFile = new File("C:\\Desktop\\aspose\\ppt_to_video\\new.pptx");
				FileOutputStream out = new FileOutputStream(outFile);
				ppt.write(out);
				out.close();
			}
		}
	}

	public static void main(String arg[]) throws Exception {

		editPPT();
}
}