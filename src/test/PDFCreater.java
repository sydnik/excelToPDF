package test;


import com.itextpdf.text.*;
import com.itextpdf.text.Document;
import com.itextpdf.text.Element;
import com.itextpdf.text.pdf.*;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;


import java.io.IOException;
import java.util.ArrayList;



public class PDFCreater {
    private Document document;
    private float[] columnWidths;
    private ReadExcel readExcel;
    private ArrayList<PdfPCell> listCell = new ArrayList<>();
    private ArrayList<DeleteCellAddress> listForDeleteCellXAndY = new ArrayList<>();
    private int[] columnZeroAndRowOneAndColumnLastTwoAndRowLastThree = null;
    private int countRegions = 0;

    public PDFCreater(Document document, ReadExcel readExcel) {
        this.document = document;
        this.readExcel = readExcel;
        setColumnWidths();
    }

    public void createdPDF() {
        try {
            document.add(createFirstTable(readExcel));



        } catch (DocumentException e) {
            e.printStackTrace();
        }


    }

    public PdfPTable createFirstTable(ReadExcel readExcel) {
        PdfPTable table = new PdfPTable(columnWidths);
        try {
            table.setTotalWidth(columnWidths);

        } catch (DocumentException e) {
            e.printStackTrace();
        }
        table.setLockedWidth(true);

        FontFactory.registerDirectories();

        PdfPCell cell = null;
        BaseFont bf = null;
        String string = "";

        Font font = null;
        int k =0;
        float kp =0;

        for (int y = 0; y < readExcel.getEndPage()[0]; y++) {
            for (int x = 0; x < readExcel.getNumberColumnList(); x++) {
                columnZeroAndRowOneAndColumnLastTwoAndRowLastThree = readExcel.getFirstCell(countRegions);
                if (readExcel.getNumberColumnList() - 1 < columnZeroAndRowOneAndColumnLastTwoAndRowLastThree[0]) {
                    countRegions++;
                }
                if(!readExcel.isNotHidden(x,y)){
                    if (y == columnZeroAndRowOneAndColumnLastTwoAndRowLastThree[1] && x == columnZeroAndRowOneAndColumnLastTwoAndRowLastThree[0]){countRegions++;}

                }
                else if(listForDeleteCellXAndY.contains(new DeleteCellAddress(x,y))){
                }
                else {
                    string = readExcel.readCell(x, y);
                    font = getStyleFontName(x, y);
                    cell = new PdfPCell(new Phrase(string, font));
                    cell.setFixedHeight(readExcel.getSheet().getRow(y).getHeight()/15);
                    if(y==0&&x==5){
                        System.out.println(readExcel.getSheet().getRow(y).getHeight()/15);
                        System.out.println(((float)readExcel.getSheet().getColumnWidth(x)) / 42.6667f);
                        System.out.println(readExcel.getSheet().getRow(y).getCell(x).getStringCellValue());
                    }
                    cell = getPDFPCellStyle(cell, x, y);
                    cell = setWidthAndHeight(cell,x,y);
                    try {
                        if (!readExcel.getSheet().getRow(y).getCell(x).getCellStyle().getWrapText()) {
                            cell.setNoWrap(true);
                        }
                    }catch (Exception e){}
                    cell.setPadding(0.5F);
                    setPositionTextInCell(cell, x, y);

                    listCell.add(cell);
                }

            }
        }

        for(int i=0;i<listCell.size();i++){
            table.addCell(listCell.get(i));
        }
        writePictureToPDF(table);
        return table;
    }

    public void writePictureToPDF(PdfPTable table) {

        Image image = null;
        for (Picture picture : readExcel.getPictures().keySet())
        try {
            image = Image.getInstance(picture.getPicture().getData());
            double summPexelX=readExcel.getPictures().get(picture).getStartPixelinCellX();
            double summPexelY = readExcel.getPictures().get(picture).getStartPixelinCellY();
            double summPixelXPic = readExcel.getPictures().get(picture).getEndPixelCellX();
            double summPixelYPic = readExcel.getPictures().get(picture).getEndPixelCellY();
            for(int i=0;i<readExcel.getPictures().get(picture).getStartCellX();i++){
                if(readExcel.isNotHiddenX(i)) {
                    summPexelX = summPexelX + table.getAbsoluteWidths()[i];
                }
            }
            for(int i=0;i<readExcel.getPictures().get(picture).getStartCellY();i++){
                if(readExcel.isNotHiddenY(i)) {
                    summPexelY = summPexelY + table.getRowHeight(i);
                }
            }
            for(int i = (int) readExcel.getPictures().get(picture).getStartCellX(); i<readExcel.getPictures().get(picture).getEndCellX(); i++){
                if(readExcel.isNotHiddenX(i)) {
                    summPixelXPic = summPixelXPic + table.getAbsoluteWidths()[i];
                }
            }
            for(int i = (int) readExcel.getPictures().get(picture).getStartCellY(); i<readExcel.getPictures().get(picture).getEndCellY(); i++){
                if(readExcel.isNotHiddenY(i)) {
                    summPixelYPic = summPixelYPic + table.getRowHeight(i);
                }
            }
            summPexelY = summPexelY + summPixelYPic;
            image.setAbsolutePosition((float) summPexelX+25,
                    document.getPageSize().getHeight()-25  -(float) summPexelY);
            image.scaleAbsolute((float) summPixelXPic, (float) summPixelYPic);
            document.add(image);


        } catch (BadElementException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (DocumentException e) {
            e.printStackTrace();
        }

    }//записываю картинки(пока не готово)

    public Font getStyleFontName(int x, int y) {
        try {
            XSSFCellStyle style = readExcel.getStyleCell(x, y);

            if(style.getFont().getBold()){
//                if(readExcel.getSheet().getRow(y).getHeightInPoints()-1-style.getFont().getFontHeightInPoints()-2<=2) {
//                    return FontFactory.getFont(style.getFont().getFontName(), "windows-1251", style.getFont().getFontHeightInPoints() - 2, Font.BOLD);
//                }
                return FontFactory.getFont(style.getFont().getFontName(), "windows-1251", style.getFont().getFontHeightInPoints(), Font.BOLD);
            }
            return FontFactory.getFont(style.getFont().getFontName(), "windows-1251", style.getFont().getFontHeightInPoints());

        } catch (Exception e) {
            return FontFactory.getFont("",10);

        }
    } // Получаю стиль ячейки(если нету, по умолчанию)
    public PdfPCell getPDFPCellStyle(PdfPCell cell, int x, int y) {

        try {
            readExcel.getStyleCell(x, y).getBorderBottomEnum().getCode();
//            cell.setCellEvent((short)readExcel.getStyleCell(x, y).getBorderBottomEnum().getCode());
            cell.setBorderWidthBottom(readExcel.getStyleCell(x, y).getBorderBottomEnum().getCode());
            cell.setBorderWidthLeft(readExcel.getStyleCell(x, y).getBorderLeftEnum().getCode());
            cell.setBorderWidthRight(readExcel.getStyleCell(x, y).getBorderRightEnum().getCode());
            cell.setBorderWidthTop(readExcel.getStyleCell(x, y).getBorderTopEnum().getCode());
//                    cell.setCellEvent(new SolidBorder(PdfPCell.BOX));

        } catch (NullPointerException e) {
            cell.setBorderWidth(PdfPCell.NO_BORDER);
        }


        return cell;
    }//Ширина бортиков


    private void setColumnWidths() {
        columnWidths = readExcel.getArrayWidthColumn();
    }//Устанавливаю ширину колонок
    private void setPositionTextInCell(PdfPCell cell, int x, int y) {
        try {


            XSSFCellStyle style = readExcel.getStyleCell(x, y);
            switch (style.getAlignmentEnum().getCode()) {
                case 1:
                    cell.setHorizontalAlignment(Element.ALIGN_LEFT);
                    break;
                case 2:
                    cell.setHorizontalAlignment(Element.ALIGN_CENTER);
                    break;
                case 3:
                    cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
                    break;

            }
            switch (style.getVerticalAlignmentEnum().ordinal()) {
                case 0:
                    cell.setVerticalAlignment(Element.ALIGN_TOP);
                    break;
                case 1:
                    cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
                    break;
                case 2:
                    cell.setVerticalAlignment(Element.ALIGN_BOTTOM);
                    break;

            }
            if(style.getShrinkToFit()) {
                cell.setNoWrap(true);
            }
        }
        catch(Exception e){

            } //ПОзиция текста в ячейки


            abstract class CustomBorder implements PdfPCellEvent {
                private int border = 0;

                public CustomBorder(int border) {
                    this.border = border;
                }

                public void cellLayout(PdfPCell cell, Rectangle position,
                                       PdfContentByte[] canvases) {
                    PdfContentByte canvas = canvases[PdfPTable.BASECANVAS];
                    canvas.saveState();
                    setLineDash(canvas);
                    if ((border & PdfPCell.TOP) == PdfPCell.TOP) {
                        canvas.moveTo(position.getRight(), position.getTop());
                        canvas.lineTo(position.getLeft(), position.getTop());
                    }
                    if ((border & PdfPCell.BOTTOM) == PdfPCell.BOTTOM) {
                        canvas.moveTo(position.getRight(), position.getBottom());
                        canvas.lineTo(position.getLeft(), position.getBottom());
                    }
                    if ((border & PdfPCell.RIGHT) == PdfPCell.RIGHT) {
                        canvas.moveTo(position.getRight(), position.getTop());
                        canvas.lineTo(position.getRight(), position.getBottom());
                    }
                    if ((border & PdfPCell.LEFT) == PdfPCell.LEFT) {
                        canvas.moveTo(position.getLeft(), position.getTop());
                        canvas.lineTo(position.getLeft(), position.getBottom());
                    }
                    canvas.stroke();
                    canvas.restoreState();
                }

                public abstract void setLineDash(PdfContentByte canvas);
            }
            class SolidBorder extends CustomBorder {
                public SolidBorder(int border) {
                    super(border);
                }

                public void setLineDash(PdfContentByte canvas) {
                }
            }

        }//Устанавливает позицию ячейки
    public PdfPCell setWidthAndHeight (PdfPCell cell,int x,int y){
        if (y == columnZeroAndRowOneAndColumnLastTwoAndRowLastThree[1] && x == columnZeroAndRowOneAndColumnLastTwoAndRowLastThree[0]) {
            int countX = 1;
            for( int i=1;i<=columnZeroAndRowOneAndColumnLastTwoAndRowLastThree[2] - columnZeroAndRowOneAndColumnLastTwoAndRowLastThree[0] ;i++){
                if(readExcel.isNotHiddenX(x+i)){
                    countX++;
                }
            }
            int countY = 1;
            for( int i=1;i<=columnZeroAndRowOneAndColumnLastTwoAndRowLastThree[3] - columnZeroAndRowOneAndColumnLastTwoAndRowLastThree[1] ;i++){
                if(readExcel.isNotHiddenY(y+i)){
                    countY++;
                }
            }
            cell.setColspan(countX);
            cell.setRowspan(countY);
            for(int deleteY =0;deleteY<=columnZeroAndRowOneAndColumnLastTwoAndRowLastThree[3] - columnZeroAndRowOneAndColumnLastTwoAndRowLastThree[1];deleteY++ ){
                for(int deleteX = 0;deleteX<=columnZeroAndRowOneAndColumnLastTwoAndRowLastThree[2] - columnZeroAndRowOneAndColumnLastTwoAndRowLastThree[0];deleteX++){
                    listForDeleteCellXAndY.add(new DeleteCellAddress(x + deleteX,y + deleteY));
                }
            }

            countRegions++;
        }
        return cell;
    }//Вычесляет какие ячейки надо пропускать и устанавливает ширину и высоту(Количетсво объедененных ячеек) ячеек



}