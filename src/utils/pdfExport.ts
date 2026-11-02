/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { jsPDF } from 'jspdf';
import html2canvas from 'html2canvas';

export interface PdfExportOptions {
  reportTitle?: string;
  orientation?: 'portrait' | 'landscape';
  paperSize?: 'a4' | 'letter';
  colorMode?: 'light' | 'dark';
  authorName?: string;
  authorEmail?: string;
  datasetName?: string;
  filename?: string;
  activeSheetName?: string;
  rowCount?: number;
  includeTimestamp?: boolean;
}

/**
 * Captures an HTML DOM element and exports it as a multi-page or scaled high-resolution PDF report.
 */
export async function exportElementToPdf(
  element: HTMLElement,
  options: PdfExportOptions = {}
): Promise<{ success: boolean; filename: string; pages: number }> {
  const {
    reportTitle = 'Executive Performance & KPI Report',
    orientation = 'landscape',
    paperSize = 'a4',
    authorName = 'Ammar Thaqif',
    authorEmail = 'ammarthaqif.ar@gmail.com',
    datasetName = 'XLSM Workbook',
    filename: sourceFilename = 'Dashboard',
    activeSheetName = 'Sheet1',
    colorMode = 'light'
  } = options;

  // Background color matching the requested report mode
  const bgColor = colorMode === 'light' ? '#ffffff' : '#020617';

  // High-DPI canvas capture (2x scale for print sharpness)
  const canvas = await html2canvas(element, {
    scale: 2,
    useCORS: true,
    logging: false,
    backgroundColor: bgColor,
    windowWidth: element.scrollWidth,
    windowHeight: element.scrollHeight
  });

  const imgData = canvas.toDataURL('image/png');

  // Initialize jsPDF
  const doc = new jsPDF({
    orientation,
    unit: 'pt',
    format: paperSize
  });

  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();

  const marginX = 28;
  const headerSpace = 38;
  const footerSpace = 28;
  const printableWidth = pageWidth - marginX * 2;
  const printableHeight = pageHeight - headerSpace - footerSpace;

  const imgWidth = printableWidth;
  const imgHeight = (canvas.height * printableWidth) / canvas.width;

  let heightLeft = imgHeight;
  let positionY = headerSpace;
  let pageNumber = 1;

  const totalPages = Math.max(1, Math.ceil(imgHeight / printableHeight));

  const drawHeaderAndFooter = (pageNum: number) => {
    // Header banner
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(8.5);
    doc.setTextColor(colorMode === 'light' ? 30 : 220, colorMode === 'light' ? 41 : 230, colorMode === 'light' ? 59 : 242);
    doc.text(reportTitle.toUpperCase(), marginX, 22);

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7.5);
    doc.setTextColor(colorMode === 'light' ? 100 : 160, colorMode === 'light' ? 116 : 175, colorMode === 'light' ? 139 : 200);
    const metaText = `${datasetName} · Sheet: "${activeSheetName}" · Ingested from ${sourceFilename}`;
    doc.text(metaText, marginX, 32);

    const dateStr = new Date().toLocaleDateString(undefined, {
      year: 'numeric',
      month: 'short',
      day: 'numeric'
    });
    doc.text(dateStr, pageWidth - marginX - doc.getTextWidth(dateStr), 22);

    // Subtle header divider
    doc.setDrawColor(colorMode === 'light' ? 226 : 51, colorMode === 'light' ? 232 : 65, colorMode === 'light' ? 240 : 85);
    doc.setLineWidth(0.5);
    doc.line(marginX, 36, pageWidth - marginX, 36);

    // Footer divider
    doc.line(marginX, pageHeight - footerSpace + 8, pageWidth - marginX, pageHeight - footerSpace + 8);

    // Footer
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7.5);
    doc.setTextColor(colorMode === 'light' ? 148 : 100, colorMode === 'light' ? 163 : 116, colorMode === 'light' ? 184 : 139);
    doc.text(
      `XLSM Cloud Studio · Developed by ${authorName} (${authorEmail}) · Confidential & Proprietary`,
      marginX,
      pageHeight - 10
    );

    const pageStr = `Page ${pageNum} of ${totalPages}`;
    doc.text(pageStr, pageWidth - marginX - doc.getTextWidth(pageStr), pageHeight - 10);
  };

  // First page
  drawHeaderAndFooter(pageNumber);
  doc.addImage(imgData, 'PNG', marginX, positionY, imgWidth, imgHeight, undefined, 'FAST');
  heightLeft -= printableHeight;

  // Additional pages if layout exceeds single page
  while (heightLeft > 0) {
    positionY = headerSpace - (imgHeight - heightLeft);
    doc.addPage(paperSize, orientation);
    pageNumber++;
    drawHeaderAndFooter(pageNumber);
    doc.addImage(imgData, 'PNG', marginX, positionY, imgWidth, imgHeight, undefined, 'FAST');
    heightLeft -= printableHeight;
  }

  // Generate safe filename
  const cleanDatasetName = datasetName.replace(/[^a-zA-Z0-9_-]/g, '_');
  const dateTag = new Date().toISOString().split('T')[0];
  const outFilename = `${cleanDatasetName}_Executive_Report_${dateTag}.pdf`;

  doc.save(outFilename);

  return {
    success: true,
    filename: outFilename,
    pages: pageNumber
  };
}
