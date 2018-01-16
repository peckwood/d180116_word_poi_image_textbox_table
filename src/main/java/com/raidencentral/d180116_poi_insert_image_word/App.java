package com.raidencentral.d180116_poi_insert_image_word;

import java.awt.Dimension;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigInteger;

import javax.imageio.ImageIO;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPicture;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTxbxContent;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STOnOff;
import org.w3c.dom.Node;

import com.microsoft.schemas.vml.CTGroup;
import com.microsoft.schemas.vml.CTShape;

/**
 * Hello world!
 *
 */
public class App {
	public static void main(String[] args) throws Exception {
		XWPFDocument document = new XWPFDocument();

		XWPFParagraph title = document.createParagraph();
		XWPFRun run0 = title.createRun();
		run0.setText("Fig.1 A Natural Scene");
		run0.setBold(true);
		title.setAlignment(ParagraphAlignment.CENTER);

		String imgFile1 = "D:/d/flower.jpg";
		FileInputStream is1 = new FileInputStream(imgFile1);
		String imgFile2 = "D:/d/horse.jpg";
		FileInputStream is2 = new FileInputStream(imgFile2);
		run0.addBreak();
		BufferedImage unresizedImage1 = ImageIO.read(new File(imgFile1));
		BufferedImage unresizedImage2 = ImageIO.read(new File(imgFile2));

		Dimension dimension1 = getResizeTargetDimension(unresizedImage1);
		Dimension dimension2 = getResizeTargetDimension(unresizedImage2);
		run0.addPicture(is1, XWPFDocument.PICTURE_TYPE_JPEG, imgFile1, Units.toEMU(dimension1.getWidth()),
				Units.toEMU(dimension1.getHeight()));
		run0.addPicture(is2, XWPFDocument.PICTURE_TYPE_JPEG, imgFile2, Units.toEMU(dimension2.getWidth()),
				Units.toEMU(dimension2.getHeight())); // pixels

		is1.close();
		is2.close();
		run0.addCarriageReturn();

		// textbox
		XWPFParagraph paragraph2 = document.createParagraph();
		CTGroup ctGroup = CTGroup.Factory.newInstance();
		CTShape ctShape = ctGroup.addNewShape();
		ctShape.setStyle("width:100pt;height:24pt");
		CTTxbxContent ctTxbxContent = ctShape.addNewTextbox().addNewTxbxContent();
		ctTxbxContent.addNewP().addNewR().addNewT()
				.setStringValue("1The TextBox text 2The TextBox text 3The TextBox text");
		Node ctGroupNode = ctGroup.getDomNode();
		CTPicture ctPicture = CTPicture.Factory.parse(ctGroupNode);
		XWPFRun run2 = paragraph2.createRun();
		CTR cTR = run2.getCTR();
		cTR.addNewPict();
		cTR.setPictArray(0, ctPicture);

		// create table
		XWPFTable table = document.createTable();
		
		table.setCellMargins(100, 100, 100, 100);//设置单元格margin
		CTTblPr pr = table.getCTTbl().getTblPr();
		pr.addNewTblW().setW(new BigInteger("100"));//设置表格宽度
		CTJc jc = pr.addNewJc();        
		jc.setVal(STJc.CENTER);//居中表格
		pr.setJc(jc);
		pr.unsetTblBorders();//隐藏表格border
		
		// create first row
		XWPFTableRow tableRowOne = table.getRow(0);
		addImageToCell(tableRowOne.getCell(0), "D:/d/flower.jpg");
		addImageToCell(tableRowOne.addNewTableCell(), "D:/d/horse.jpg");
		addImageToCell(tableRowOne.addNewTableCell(), "D:/d/flower.jpg");

		// create second row
		XWPFTableRow tableRowTwo = table.createRow();
		tableRowTwo.getCell(0).getParagraphArray(0).setAlignment(ParagraphAlignment.LEFT);
		tableRowTwo.getCell(1).getParagraphArray(0).setAlignment(ParagraphAlignment.LEFT);
		tableRowTwo.getCell(2).getParagraphArray(0).setAlignment(ParagraphAlignment.LEFT);
		XWPFRun run3 = tableRowTwo.getCell(0).getParagraphArray(0).createRun();
		run3.setText("Some Text");
		run3.addBreak();
		run3.setText("Some Text");
		
		XWPFRun run4 = tableRowTwo.getCell(1).getParagraphArray(0).createRun();
		run4.setText("Some Text");
		run4.addBreak();
		run4.setText("Some Text");
		
		XWPFRun run5 = tableRowTwo.getCell(2).getParagraphArray(0).createRun();
		run5.setText("Some Text");
		run5.addBreak();
		run5.setText("Some Text");
		
		FileOutputStream fos = new FileOutputStream("test4.docx");
		document.write(fos);
		fos.close();
		document.close();
		System.out.println("finished");
	}

	public static Dimension getResizeTargetDimension(BufferedImage image) throws Exception {
		int resultHeight = 100;
		Dimension IMG_MAX_SIZE = new Dimension(10000, resultHeight);
		Dimension targetDemension = getScaledDimension(new Dimension(image.getWidth(), image.getHeight()),
				IMG_MAX_SIZE);
		System.out.println("target width: " + targetDemension.getWidth());
		System.out.println("target height: " + targetDemension.getHeight());
		return targetDemension;
		// ImageIO.write(image, "jpg", new File("D:/d/flower1.jpg"));
	}

	public static Dimension getScaledDimension(Dimension imgSize, Dimension boundary) {

		int original_width = imgSize.width;
		int original_height = imgSize.height;
		int bound_width = boundary.width;
		int bound_height = boundary.height;
		int new_width = original_width;
		int new_height = original_height;

		// first check if we need to scale width
		if (original_width != bound_width) {
			// scale width to fit
			new_width = bound_width;
			// scale height to maintain aspect ratio
			new_height = (new_width * original_height) / original_width;
		}

		// then check if we need to scale even with the new height
		if (new_height > bound_height) {
			// scale height to fit instead
			new_height = bound_height;
			// scale width to maintain aspect ratio
			new_width = (new_height * original_width) / original_height;
		}

		return new Dimension(new_width, new_height);
	}

	public static void addImageToCell(XWPFTableCell cell, String imageLocation) throws Exception {
		
		XWPFParagraph cellParagraph = cell.getParagraphArray(0);
		XWPFRun run0 = cellParagraph.createRun();
		cellParagraph.setAlignment(ParagraphAlignment.CENTER);

		FileInputStream is1 = new FileInputStream(imageLocation);
		BufferedImage unresizedImage1 = ImageIO.read(new File(imageLocation));

		Dimension dimension1 = getResizeTargetDimension(unresizedImage1);
		run0.addPicture(is1, XWPFDocument.PICTURE_TYPE_JPEG, imageLocation, Units.toEMU(dimension1.getWidth()),
				Units.toEMU(dimension1.getHeight()));
		is1.close();
		cell.addParagraph(cellParagraph);
	}
}
