/************************************************************************
 * 
 * DO NOT ALTER OR REMOVE COPYRIGHT NOTICES OR THIS FILE HEADER
 * 
 * Copyright 2011 IBM. All rights reserved.
 * 
 * Use is subject to license terms.
 * 
 * Licensed under the Apache License, Version 2.0 (the "License"); you may not
 * use this file except in compliance with the License. You may obtain a copy of
 * the License at http://www.apache.org/licenses/LICENSE-2.0. You can also
 * obtain a copy of the License at http://odftoolkit.org/docs/license.txt
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS, WITHOUT
 * WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * 
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * 
 ************************************************************************/
package odftoolkit;

import java.awt.Rectangle;
import java.util.Iterator;

import org.odftoolkit.odfdom.type.CellRangeAddressList;
import org.odftoolkit.simple.PresentationDocument;
import org.odftoolkit.simple.SpreadsheetDocument;
import org.odftoolkit.simple.TextDocument;
import org.odftoolkit.simple.chart.Chart;
import org.odftoolkit.simple.chart.ChartType;
import org.odftoolkit.simple.draw.Textbox;
import org.odftoolkit.simple.presentation.Slide;
import org.odftoolkit.simple.presentation.Slide.SlideLayout;
import org.odftoolkit.simple.table.Cell;
import org.odftoolkit.simple.table.Table;
import org.odftoolkit.simple.text.list.List;
import org.odftoolkit.simple.text.list.ListItem;

public class DocumentConstructor {

	public static void main(String[] args) {
		generatePresentationChart();
		generateTextDocument();
		generateSpreadsheetDocument();
	}

	private static void generatePresentationChart() {
		try {
			PresentationDocument presentationDoc = PresentationDocument
					.newPresentationDocument();
			SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument
					.loadDocument("demo9_data.ods");

			// create cover page
			Slide slide = presentationDoc.newSlide(0, "Slide1",	SlideLayout.TITLE_SUBTITLE);
			setSlideTextContent(slide, "Simple Website Analytics Report", "2011-04-27~2011-05-27");

			// create visitors overview page 1
			slide = presentationDoc.newSlide(1, "Slide2", SlideLayout.TITLE_PLUS_3_OBJECT);
			setSlideTextContent(slide, "Visitors Overview");
			Table tableA = spreadsheetDoc.getTableByName("A");
			convertFromTableToList(tableA, slide.addList(), 4, 17, 5, 20);
			CellRangeAddressList cellRange = CellRangeAddressList.valueOf("A.A1:A.B3");
			Chart chart = slide.createChart(
					"New Visitor VS. Returning Visitor", spreadsheetDoc, cellRange, true, true, false, null);
			chart.setChartType(ChartType.PIE);
			cellRange = CellRangeAddressList.valueOf("A.A6:A.B37");
			chart = slide.createChart("Daily Visit", spreadsheetDoc, cellRange,	true, true, false, null);
			chart.setChartType(ChartType.LINE);

			// create visitors overview page 2
			slide = presentationDoc.newSlide(2, "Slide3", SlideLayout.TITLE_PLUS_2_CHART);
			setSlideTextContent(slide, "Visitors Overview");
			cellRange = CellRangeAddressList.valueOf("A.E1:A.G14");
			chart = slide.createChart("Count of Visits", spreadsheetDoc, cellRange, true, true, false, null);
			chart.setChartType(ChartType.BAR);
			cellRange = CellRangeAddressList.valueOf("A.I1:A.K8");
			chart = slide.createChart("Visit Duration", spreadsheetDoc,	cellRange, true, true, false, null);

			// create traffic sources overview page
			slide = presentationDoc.newSlide(3, "Slide4", SlideLayout.TITLE_PLUS_4_OBJECT);
			setSlideTextContent(slide, "Traffic Sources Overview");
			Table tableB = spreadsheetDoc.getTableByName("B");
			convertFromTableToList(tableB, slide.addList(), 0, 2, 1, 4);
			cellRange = CellRangeAddressList.valueOf("B.A2:B.C5");
			chart = slide.createChart("Traffic Sources Type", spreadsheetDoc, cellRange, true, true, false, null);
			chart.setChartType(ChartType.PIE);
			cellRange = CellRangeAddressList.valueOf("B.A9:B.C19");
			chart = slide.createChart("Referral Traffic", spreadsheetDoc, cellRange, true, true, false, null);
			chart.setChartType(ChartType.PIE);
			cellRange = CellRangeAddressList.valueOf("B.E2:B.G8");
			chart = slide.createChart("Search Keyword", spreadsheetDoc,	cellRange, true, true, false, null);
			chart.setChartType(ChartType.PIE);

			// create content overview page
			slide = presentationDoc.newSlide(4, "Slide5",
					SlideLayout.TITLE_PLUS_CHART);
			setSlideTextContent(slide, "Content Overview");
			cellRange = CellRangeAddressList.valueOf("C.A1:C.C8");
			chart = slide.createChart("Page Visit", spreadsheetDoc, cellRange, true, true, false, null);
			chart.setChartType(ChartType.BAR);

			spreadsheetDoc.close();
			presentationDoc.save("demo9p.odp");
			presentationDoc.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private static void generateTextDocument() {
		try {
			TextDocument textDoc = TextDocument.newTextDocument();
			SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.loadDocument("demo9_data.ods");

			// create cover page
			textDoc.addParagraph("Simple Website Analytics Report(2011-04-27~2011-05-27)");
			textDoc.addParagraph("Visitors Overview");
			// create visitors overview page 1
			CellRangeAddressList cellRange = CellRangeAddressList.valueOf("A.A1:A.B3");
			Rectangle rect = new Rectangle();
			rect.width = 14000;
			rect.height = 8000;
			Chart chart = textDoc.createChart("New Visitor VS. Returning Visitor", spreadsheetDoc, cellRange, true, true, false, rect);
			chart.setChartType(ChartType.PIE);
			cellRange = CellRangeAddressList.valueOf("A.A6:A.B37");
			chart = textDoc.createChart("Daily Visit", spreadsheetDoc,cellRange, true, true, false, rect);
			chart.setChartType(ChartType.LINE);
			cellRange = CellRangeAddressList.valueOf("A.E1:A.G14");
			chart = textDoc.createChart("Count of Visits", spreadsheetDoc, cellRange, true, true, false, rect);
			chart.setChartType(ChartType.BAR);
			cellRange = CellRangeAddressList.valueOf("A.I1:A.K8");
			chart = textDoc.createChart("Visit Duration", spreadsheetDoc, cellRange, true, true, false, rect);

			// create traffic sources overview page
			textDoc.addParagraph("Traffic Sources Overview");
			cellRange = CellRangeAddressList.valueOf("B.A2:B.C5");
			chart = textDoc.createChart("Traffic Sources Type", spreadsheetDoc,	cellRange, true, true, false, rect);
			chart.setChartType(ChartType.PIE);
			cellRange = CellRangeAddressList.valueOf("B.A9:B.C19");
			chart = textDoc.createChart("Referral Traffic", spreadsheetDoc,	cellRange, true, true, false, rect);
			chart.setChartType(ChartType.PIE);
			cellRange = CellRangeAddressList.valueOf("B.E2:B.G8");
			chart = textDoc.createChart("Search Keyword", spreadsheetDoc, cellRange, true, true, false, rect);
			chart.setChartType(ChartType.PIE);

			// create content overview page
			textDoc.addParagraph("Content Overview");
			cellRange = CellRangeAddressList.valueOf("C.A1:C.C8");
			chart = textDoc.createChart("Page Visit", spreadsheetDoc, cellRange, true, true, false, rect);
			chart.setChartType(ChartType.BAR);
			
			spreadsheetDoc.close();
			textDoc.save("demo9t.odt");
			textDoc.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	private static void generateSpreadsheetDocument() {
		try {
			SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.loadDocument("demo9_data.ods");
			// create visitors overview page 1
			CellRangeAddressList cellRange = CellRangeAddressList.valueOf("A.A1:A.B3");
			Rectangle rect = new Rectangle();
			rect.width = 15000;
			rect.height = 8000;
			Cell positionCell = spreadsheetDoc.getTableByName("B").getCellByPosition("E1");
			spreadsheetDoc.createChart("Page Visit", spreadsheetDoc, cellRange,	true, true, false, rect, positionCell);
			spreadsheetDoc.save("demo9s.ods");
			spreadsheetDoc.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// set slide text content
	private static void setSlideTextContent(Slide slide, String... texts) {
		int i = 0;
		Iterator<Textbox> textBoxIterator = slide.getTextboxIterator();
		while (textBoxIterator.hasNext()) {
			Textbox title = textBoxIterator.next();
			title.setTextContent(texts[i++]);
		}
	}

	// convert from table to list
	private static void convertFromTableToList(Table table, List list,
			int startColumn, int startRow, int endColumn, int endRow) {
		while (startRow <= endRow) {
			Cell cell = table.getCellByPosition(startColumn, startRow);
			int rowSpannedNumber = cell.getRowSpannedNumber();
			String cellText = cell.getDisplayText();
			if (!"".equals(cellText)) {
				ListItem item = list.addItem(cellText);
				int columnSpannedNumber = cell.getColumnSpannedNumber();
				int newStartColumn = startColumn + columnSpannedNumber;
				if (newStartColumn <= endColumn) {
					if (rowSpannedNumber > 1) {
						List subList = item.addList();
						convertFromTableToList(table, subList, newStartColumn,
								startRow, endColumn, startRow
										+ rowSpannedNumber - 1);
					} else {
						int tmpStartColumn = newStartColumn;
						while (tmpStartColumn <= endColumn) {
							cell = table.getCellByPosition(tmpStartColumn,
									startRow);
							cellText = cell.getDisplayText();
							if (!"".equals(cellText)) {
								item.setTextContent(item.getTextContent()
										+ "    " + cellText);
							}
							tmpStartColumn += cell.getColumnSpannedNumber();
						}
					}
				}
			}
			startRow += rowSpannedNumber;
		}
	}
}
