package org.pptx4j.samples;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;

import org.apache.commons.io.FileUtils;
import org.docx4j.XmlUtils;
import org.docx4j.dml.TextFont;
import org.docx4j.dml.chart.CTChartSpace;
import org.docx4j.dml.chart.CTNumVal;
import org.docx4j.dml.chart.CTStrVal;
import org.docx4j.dml.chart.CTUnsignedInt;
import org.docx4j.dml.chart.ListSer;
import org.docx4j.dml.chart.SerContent;
import org.docx4j.openpackaging.contenttype.ContentTypes;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.exceptions.PartUnrecognisedException;
import org.docx4j.openpackaging.io.SaveToZipFile;
import org.docx4j.openpackaging.packages.PresentationMLPackage;
import org.docx4j.openpackaging.packages.SpreadsheetMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.ThemePart;
import org.docx4j.openpackaging.parts.DrawingML.Chart;
import org.docx4j.openpackaging.parts.PresentationML.MainPresentationPart;
import org.docx4j.openpackaging.parts.PresentationML.SlideLayoutPart;
import org.docx4j.openpackaging.parts.PresentationML.SlideMasterPart;
import org.docx4j.openpackaging.parts.PresentationML.SlidePart;
import org.docx4j.openpackaging.parts.PresentationML.ViewPropertiesPart;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorksheetPart;
import org.docx4j.openpackaging.parts.WordprocessingML.EmbeddedPackagePart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.relationships.Relationship;
import org.docx4j.utils.BufferUtil;
import org.pptx4j.jaxb.Context;
import org.pptx4j.pml.CTGraphicalObjectFrame;
import org.pptx4j.pml.Sld;
import org.pptx4j.pml.SldLayout;
import org.pptx4j.pml.ViewPr;
import org.xlsx4j.sml.Cell;
import org.xlsx4j.sml.Row;
import org.xlsx4j.sml.STCellType;

public class Docx4jChartUtility {
	private static final String PPT_TEMPLATE_FILE = "D:/exportTemplates/blank_template.pptx";
    private static final String OUTPUT_FILE = "D:/ppt/Docx4j.pptx";
	private static final String PPT_PIECHART_XML = "/ppt/charts/piechart.xml";
    private static final String PPT_BARCHART_XML = "/ppt/charts/barchart.xml";
    private static final String BARCHART_XML_FILEPATH = "D:/exportTemplates/charts/barchart.xml";
    private static final String PIECHART_XML_FILEPATH = "D:/exportTemplates/charts/piechart.xml";
    private static final String SLIDELAYOUT_XML_FILEPATH = "D:/exportTemplates/layout/slideLayout7.xml";
    private static final String TABLE_FRAME_XML_FILEPATH = "D:/exportTemplates/table/tableGraphicFrame.xml";
    private static final String CHART_FRAME_XML_FILEPATH = "D:/exportTemplates/charts/chartGraphicFrame.xml";
    private static final String SLIDE_PART_XMLPATH = "D:/exportTemplates/slidePart.xml";
    private static final String VIEWPROPS_XMLPATH = "D:/exportTemplates/viewProps.xml";

    public enum ChartType {
        BAR, PIE, COLUMN
    };

    public static void main(String[] args) throws Exception {
    	List<Object[]> exportData = getExportData();

    	Docx4jChartUtility docx4jChartUtility = new Docx4jChartUtility() ;
    	long startTime = System.currentTimeMillis();
    	docx4jChartUtility.createPowerpoint(exportData, 10, Arrays.asList("BAR", "PIE"//,"BAR", "PIE","BAR", "PIE","BAR", "PIE"
//    			,"BAR", "PIE","BAR", "PIE","BAR", "PIE"
//    			,"BAR", "PIE","BAR", "PIE","BAR", "PIE"
//    			,"BAR", "PIE","BAR", "PIE","BAR", "PIE"
//    			,"BAR", "PIE","BAR", "PIE","BAR", "PIE"
//    			,"BAR", "PIE","BAR", "PIE","BAR", "PIE"
//    			,"BAR", "PIE","BAR", "PIE","BAR", "PIE"
//    			,"BAR", "PIE","BAR", "PIE","BAR", "PIE"
    			));
    	System.out.println("Total seconds taken:" +(System.currentTimeMillis()-startTime)/1000.0);
	}

    public void createPowerpoint(List<Object[]> exportData, long numberOfDataPoints, List<String> chartTypes)
            throws Exception {
        PresentationMLPackage presentationMLPackage = PresentationMLPackage.load(new File(PPT_TEMPLATE_FILE));
        MainPresentationPart mainPresentationPart = (MainPresentationPart) presentationMLPackage.getParts().getParts()
                .get(new PartName("/ppt/presentation.xml"));

        /*ViewPropertiesPart viewPropertiesPart = new ViewPropertiesPart(new PartName("/ppt/viewProps.xml"));
        viewPropertiesPart.setJaxbElement((ViewPr) unmarshal(VIEWPROPS_XMLPATH));
        mainPresentationPart.addTargetPart(viewPropertiesPart);
         */
        updateTheme(presentationMLPackage);

        SlideLayoutPart layoutPart = createSlideLayout(presentationMLPackage);

        String chartDataExcelName = null;
        ChartType chartType = null;
        SlidePart slidePart = null;
        Chart chart = null;
        Relationship relationship = null;
        EmbeddedPackagePart embPackage = null;

        List<String> horAxisLabels = getHorizontalAxisLabels(exportData);
        List<String> chartData = getChartData(exportData);

        int slideNumber = 1;
        for (String chartTypeString : chartTypes) {
            chartType = getChartType(chartTypeString);

            slidePart = addSlide(mainPresentationPart, layoutPart, slideNumber,
					chartTypeString);

            chart = getChart(chartType, slideNumber);
            relationship = slidePart.addTargetPart(chart);

            chartDataExcelName = "/ppt/embeddings/" + chartTypeString.toLowerCase() + "chart"+ slideNumber +".xlsx";
            embPackage = getExcelPartForChart(presentationMLPackage, chartDataExcelName);

            slidePart.getJaxbElement().getCSld().getSpTree().getSpOrGrpSpOrGraphicFrame()
                    .add(createChartFrame(relationship, chart.getJaxbElement()));

            chart.addTargetPart(embPackage);
            presentationMLPackage.addTargetPart(embPackage);
            createChart(exportData, numberOfDataPoints, chartType, presentationMLPackage, slideNumber, horAxisLabels, chartData);
            slideNumber++;
        }
        CTGraphicalObjectFrame graphicFrame =  unmarshalToCTGraphicalObjectFrame(TABLE_FRAME_XML_FILEPATH);

        slidePart = PresentationMLPackage.createSlidePart(mainPresentationPart, layoutPart, new PartName("/ppt/slides/slide_table.xml"));
        slidePart.getJaxbElement().getCSld().getSpTree().getSpOrGrpSpOrGraphicFrame().add( graphicFrame );

        presentationMLPackage.setTitle("PptTitle");
        mainPresentationPart.removeSlide(0);

        presentationMLPackage.save(new java.io.File(OUTPUT_FILE));
        System.out.println("\n\n done .. saved " + OUTPUT_FILE);

        System.out.println("Total number of slides: "+(slideNumber -1));
    }

    private void createChart(List<Object[]> exportData, long numberOfDataPoints, ChartType chartType,
            PresentationMLPackage presentationMLPackage, int slideNumber, List<String> horAxisLabels, List<String> chartData) throws Docx4JException, JAXBException {

        String chartPartNameString = getChartPartName(chartType, slideNumber);

        PartName chartPartName = new PartName(chartPartNameString);
        Chart chart = (Chart) presentationMLPackage.getParts().get(chartPartName);

        List<Object> objects = chart.getJaxbElement().getChart().getPlotArea().getAreaChartOrArea3DChartOrLineChart();

        updateChartXml(numberOfDataPoints, horAxisLabels, chartData, objects, chartType);

        updateChartDataInExcel(presentationMLPackage, numberOfDataPoints, horAxisLabels, chartData, objects, chartType, slideNumber);
    }

    private void updateChartXml(long numberOfDataPoints, List<String> horAxisLabels, List<String> chartData,
            List<Object> objects, ChartType chartType) {
        CTUnsignedInt dataPointSize = new CTUnsignedInt();
        dataPointSize.setVal(numberOfDataPoints);

        for (Object object : objects) {

                List<SerContent> ctSers = ((ListSer) object).getSer();

                for (SerContent ctSerContent : ctSers) {

                	ctSerContent.getTx().getStrRef().getStrCache().getPt().get(0).setV("Custom title");

                    List<CTStrVal> axisLabels = ctSerContent.getCat().getStrRef().getStrCache().getPt();
                    int horAxisLabelCounter = 0;
                    for (CTStrVal axis : axisLabels) {
                        if (horAxisLabelCounter > numberOfDataPoints) {
                            break;
                        }
                        axis.setV(horAxisLabels.get(horAxisLabelCounter++));
                    }
                    // if there are more labels than in the template
                    while (horAxisLabelCounter < horAxisLabels.size()) {
                        CTStrVal axis = new CTStrVal();
                        axis.setV(horAxisLabels.get(horAxisLabelCounter++));
                        axisLabels.add(axis);
                    }
                    ctSerContent.getCat().getStrRef().getStrCache().setPtCount(dataPointSize);

                    List<CTNumVal> ctNumVals = ctSerContent.getVal().getNumRef().getNumCache().getPt();
                    int chartDataCounter = 0;
                    for (CTNumVal ctNumVal : ctNumVals) {
                        if (chartDataCounter > numberOfDataPoints) {
                            break;
                        }
                        ctNumVal.setV(chartData.get(chartDataCounter++));
                    }

                    // if there are more data points than in the template
                    while (chartDataCounter < numberOfDataPoints) {
                        CTNumVal ctNumVal = new CTNumVal();
                        ctNumVal.setV(chartData.get(chartDataCounter++));
                        ctNumVals.add(ctNumVal);
                    }
                    ctSerContent.getVal().getNumRef().getNumCache().setPtCount(dataPointSize);
                }
            }
    }

    private void updateChartDataInExcel(PresentationMLPackage ppt, long numberOfDataPoints, List<String> horAxisLabels,
            List<String> chartData, List<Object> objects, ChartType chartType, int slideNumber) throws Docx4JException {
    	String xlsPartName ="/ppt/embeddings/" + chartType.toString().toLowerCase() + "chart"+ slideNumber +".xlsx";
        EmbeddedPackagePart epp = (EmbeddedPackagePart) ppt.getParts().get(new PartName(xlsPartName));

        if (epp == null) {
            throw new Docx4JException("Could find EmbeddedPackagePart: " + xlsPartName);
        }

        InputStream is = BufferUtil.newInputStream(epp.getBuffer());

        SpreadsheetMLPackage spreadSheet = (SpreadsheetMLPackage) SpreadsheetMLPackage.load(is);

        Map<PartName, Part> partsMap = spreadSheet.getParts().getParts();
        Iterator<Entry<PartName, Part>> it = partsMap.entrySet().iterator();

        while (it.hasNext()) {
            Map.Entry<PartName, Part> pairs = it.next();

            if (partsMap.get(pairs.getKey()) instanceof WorksheetPart) {

                WorksheetPart wsp = (WorksheetPart) partsMap.get(pairs.getKey());

                List<Row> rows = wsp.getJaxbElement().getSheetData().getRow();

                int index = 1;
                for (; index < rows.size(); index++) {
                    Row row = rows.get(index);
                    List<Cell> cells = row.getC();
                    for (Cell cell : cells) {
                        if (index > numberOfDataPoints) {
                            break;
                        }

                        if (cell.getR().equals("A" + (index + 1))) {
                            cell.setT(STCellType.STR);
                            cell.setV(horAxisLabels.get(index - 1));
                        } else if (cell.getR().equals("B" + (index + 1))) {
                            cell.setT(STCellType.N);
                            cell.setV(chartData.get(index - 1));
                        }
//                        System.out.println(cell.getV());
                    }
                }

                while (index <= numberOfDataPoints) {
                    Row row = new Row();

                    Cell cell = new Cell();
                    cell.setR("A" + (index + 1));
                    cell.setT(STCellType.STR);
                    cell.setV(horAxisLabels.get(index - 1));

//                    System.out.println(cell.getV());

                    row.getC().add(cell);

                    cell = new Cell();
                    cell.setR("B" + (index + 1));
                    cell.setT(STCellType.N);
                    cell.setV(chartData.get(index - 1));

//                    System.out.println(cell.getV());

                    row.getC().add(cell);

                    rows.add(row);
                    index++;
                }
            }
        }

        /*
         * Convert the Spreadsheet to a binary format, set it on the
         * EmbeddedPackagePart, add it back onto the deck and save to a file.
         */
        ByteArrayOutputStream baos = new ByteArrayOutputStream();

        SaveToZipFile saver = new SaveToZipFile(spreadSheet);

        saver.save(baos);
        epp.setBinaryData(baos.toByteArray());
    }

    private CTGraphicalObjectFrame createChartFrame(Relationship relationship, CTChartSpace chartSpace)
            throws Exception {
        CTGraphicalObjectFrame graphicFrame = unmarshalToCTGraphicalObjectFrame(CHART_FRAME_XML_FILEPATH);
        return graphicFrame;
    }

    private EmbeddedPackagePart getExcelPartForChart(PresentationMLPackage presentationMLPackage,
            String chartDataExcelName) throws InvalidFormatException, PartUnrecognisedException, FileNotFoundException {

        Relationship rel = new Relationship();
        rel.setType(Namespaces.EMBEDDED_PKG);
        EmbeddedPackagePart embPackage = (EmbeddedPackagePart) presentationMLPackage.getContentTypeManager()
                .newPartForContentType(ContentTypes.SPREADSHEETML_WORKBOOK, chartDataExcelName, rel);
        org.docx4j.relationships.Relationship sourceRel = new org.docx4j.relationships.ObjectFactory()
                .createRelationship();
        rel.setType(Namespaces.EMBEDDED_PKG);
        rel.setTarget("D:/exportTemplates/embeddings/Chart.xlsx");
        rel.setTargetMode("External");
        // TODO: remove hard coding
        rel.setId("rId3");

        InputStream fis = new FileInputStream(new File("D:/exportTemplates/embeddings/Chart.xlsx"));
        embPackage.setBinaryData(fis);

        embPackage.setSourceRelationship(sourceRel);
        return embPackage;
    }

    private Chart getBarChartPart(int slideNumber) throws Exception {
        CTChartSpace el = ((JAXBElement<CTChartSpace>) unmarshal(BARCHART_XML_FILEPATH)).getValue();
        PartName chartPartName = new PartName(PPT_BARCHART_XML.replace(".", "" + slideNumber + "."));
        Chart chart = new Chart(chartPartName);
        chart.setJaxbElement(el);

        return chart;
    }

    private Chart getPieChartPart(int slideNumber) throws Exception {
        CTChartSpace el = ((JAXBElement<CTChartSpace>) unmarshal(PIECHART_XML_FILEPATH)).getValue();
        PartName chartPartName = new PartName(PPT_PIECHART_XML.replace(".", "" + slideNumber + "."));
        Chart chart = new Chart(chartPartName);
        chart.setJaxbElement(el);

        return chart;
    }


    private Chart getChart(ChartType chartType, int slideNumber) throws Exception {
        switch (chartType) {
        case BAR:
            return getBarChartPart(slideNumber);
        case PIE:
            return getPieChartPart(slideNumber);
        default:
            throw new RuntimeException("Invalid chart type");
        }
    }

    private String getChartPartName(ChartType chartType, int slideNumber) {
        String chartPartNameString = null;
        switch (chartType) {
        case BAR:
            chartPartNameString = PPT_BARCHART_XML.replace(".", "" + slideNumber + ".");
            break;
        case PIE:
            chartPartNameString = PPT_PIECHART_XML.replace(".", "" + slideNumber + ".");
            break;
        default:
            break;
        }
        return chartPartNameString;
    }

    private ChartType getChartType(String type) {
        if (type.equalsIgnoreCase("BAR")) {
            return ChartType.BAR;
        } else if (type.equalsIgnoreCase("PIE")) {
            return ChartType.PIE;
        } else {
            throw new RuntimeException("Invalid chart type");
        }
    }

    private CTGraphicalObjectFrame unmarshalToCTGraphicalObjectFrame(String filePath) throws Exception {
    	String tableXml = FileUtils.readFileToString(new File(filePath));
        return (CTGraphicalObjectFrame) XmlUtils.unmarshalString(tableXml, Context.jcPML, CTGraphicalObjectFrame.class);
    }

    private Object unmarshal(String filePath) throws Exception {
    	InputStream fis = new FileInputStream(new File(filePath));
        return XmlUtils.unmarshal(fis, Context.jcPML);
    }

    private SlideLayoutPart createSlideLayout(
			PresentationMLPackage presentationMLPackage)
			throws InvalidFormatException, Exception {
		/*SlideMasterPart masterPart = (SlideMasterPart) presentationMLPackage.getParts().getParts().get(new PartName("/ppt/slideMasters/slideMaster1.xml"));


        SlideLayoutPart layoutPart = new SlideLayoutPart(new PartName("/ppt/slideLayouts/slideLayout7.xml"));
        SldLayout sldLayout = (SldLayout) unmarshal(SLIDELAYOUT_XML_FILEPATH);
        layoutPart.setJaxbElement( sldLayout );
        masterPart.addSlideLayoutIdListEntry(layoutPart);
        layoutPart.addTargetPart(masterPart);

		return layoutPart;*/
		return (SlideLayoutPart)presentationMLPackage.getParts().getParts().get(
				new PartName("/ppt/slideLayouts/slideLayout2.xml"));
	}

	private SlidePart addSlide(MainPresentationPart mainPresentationPart,
			SlideLayoutPart layoutPart, int slideNumber, String chartTypeString)
			throws InvalidFormatException, Exception {
		SlidePart slidePart;
		slidePart = new SlidePart(new PartName("/ppt/slides/slide_"+ chartTypeString + slideNumber + ".xml"));
		mainPresentationPart.addSlideIdListEntry(slidePart);

		String slidePartString = FileUtils.readFileToString(new File(SLIDE_PART_XMLPATH));
		slidePartString = slidePartString.replace("Bar Chart", chartTypeString+" Chart");
		slidePartString = slidePartString.replace("CustomFooter", "CustomFooter for "+chartTypeString+" Chart");

		slidePart.setJaxbElement((Sld) XmlUtils.unmarshalString(slidePartString, Context.jcPML));
		// Slide layout part
		slidePart.addTargetPart(layoutPart);
		return slidePart;
	}

	private ThemePart updateTheme(PresentationMLPackage presentationMLPackage)
			throws InvalidFormatException {
		ThemePart theme = (ThemePart)presentationMLPackage.getParts().getParts().get(new PartName("/ppt/theme/theme1.xml"));
        // set the font family
        TextFont tx = new TextFont();
        tx.setTypeface("Jokerman");
        theme.getFontScheme().getMajorFont().setLatin(tx);
        theme.getFontScheme().getMinorFont().setLatin(tx);
        return theme;
	}

    // data conversion methods
    private List<String> getHorizontalAxisLabels(List<Object[]> exportData) {
        List<String> labels = new ArrayList<String>();
        for (Object[] objectArray : exportData) {
            labels.add((String) objectArray[2]);
        }
        return labels;
    }

    private List<String> getChartData(List<Object[]> exportData) {
        List<String> chartData = new ArrayList<String>();
        for (Object[] objectArray : exportData) {
            chartData.add(objectArray[0].toString());
        }
        return chartData;
    }

    private static List<Object[]> getExportData() throws Exception {
    	List<Object[]> exportData = new ArrayList<Object[]>();
		SpreadsheetMLPackage pkg = SpreadsheetMLPackage.load(new java.io.File("D:/ppt/ChartDataSource.xlsx"));
		WorksheetPart wsp = pkg.getWorkbookPart().getWorksheet(0);

		List<Row> rows = wsp.getJaxbElement().getSheetData().getRow();

        int index = 1;

        Object[] objectArray = null;
        for (; index < rows.size(); index++) {
        	exportData.add(index-1, new Object[3]);
            Row row = rows.get(index);
            List<Cell> cells = row.getC();
            for (Cell cell : cells) {
            	objectArray = exportData.get(index-1);
                if (cell.getR().equals("C" + (index + 1))) {
                	objectArray[0] =  cell.getV();
                } else if (cell.getR().equals("B" + (index + 1))) {
                	objectArray[1] = cell.getV();
                }else if (cell.getR().equals("A" + (index + 1))) {
                	objectArray[2] = "Category" + cell.getV();
                }
            }
        }
        return exportData;
    }
}
