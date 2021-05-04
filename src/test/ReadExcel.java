package test;



import com.itextpdf.text.PageSize;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import org.apache.poi.ss.usermodel.PictureData;

import java.text.DecimalFormat;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;


public class ReadExcel {
    private double width1PixelPDF;
    private XSSFSheet sheet;
    private HashMap<Picture,AddressForPicture> pictures =new HashMap<>();
    private ArrayList<Integer> columnsNumbersNotHidden = new ArrayList<>();//СТолбики на печать
    private ArrayList<Integer> rowsNumbersNotHidden = new ArrayList<>();//Строки на печать
    private int[] numberRow = {};//последние строки на страницу
    private int finishNumberOfColumns;//Последний столбик для печати(С учетом скрытых)
    private int finishNumberOfRow;//Последняя строка для печати (С учетом скрытых)
    private int startNumberOfColumns;//первый столбик для печати
    private int startNumberOfRow;// первая строка для печати
    private ArrayList<CellRangeAddress> cellRangeAddressList;



    public ReadExcel(XSSFSheet sheet) {
        this.sheet = sheet;
        setNumberOfColumnsAndRow();
        columnsNumbersNotHiddenAndRowsNumbersNotHidden();
       readImage();
        getMergedRegions();
        setListMergedRegion();

    }

    private void setNumberOfColumnsAndRow(){
        String string = sheet.getWorkbook().getPrintArea(sheet.getWorkbook().getActiveSheetIndex());
        String[] row = string.split("\\$");
        finishNumberOfRow = Integer.parseInt(row[row.length-1])-1;
        finishNumberOfColumns = (row[row.length-2].charAt(0)-64);
        startNumberOfColumns = row[row.length-4].charAt(0)-65;
        startNumberOfRow = Integer.parseInt(row[row.length-3].replace(":",""))-1;
    }// устанавливает область печати
    private void columnsNumbersNotHiddenAndRowsNumbersNotHidden(){
        for(int i=0;i<getNumberColumnList();i++){
            if(!sheet.isColumnHidden(i)){
                columnsNumbersNotHidden.add(i);
            }
        }
        for(int i=0;i<getNumberRowList();i++){
            try {

                if (!sheet.getRow(i).getZeroHeight()) {
                    rowsNumbersNotHidden.add(i);
                }
            }
            catch (Exception e){
                rowsNumbersNotHidden.add(i);
            }
        }
    }// записывает в лист индексы столбцов и строк которые не скрыты
    private void getMergedRegions (){
        for (int i= 0;i<sheet.getMergedRegions().size();i++) {
        }
    }

    public void readImage(){// чтение картинок и место их расположения
        XSSFClientAnchor clientAnchor = null;
        XSSFPicture inpPic = null;
        XSSFDrawing dp = sheet.createDrawingPatriarch();
        List<XSSFShape> pics = dp.getShapes();
        for(int i=0;i<pics.size();i++) {
            inpPic = (XSSFPicture) pics.get(i);
            clientAnchor = inpPic.getClientAnchor();
            String name = inpPic.getShapeName();
            PictureData pictureData = inpPic.getPictureData();
            int x = clientAnchor.getCol1();// эти числа скорее всего не правильны
            int y = clientAnchor.getRow1();// эти числа скорее всего не правильны
            Picture picture = new Picture(name,pictureData,x,y);
            pictures.put(picture,new AddressForPicture(clientAnchor.getCol1(),clientAnchor.getRow1(),clientAnchor.getDx1(),clientAnchor.getDy1(),
                    clientAnchor.getCol2(),clientAnchor.getRow2(),clientAnchor.getDx2(),clientAnchor.getDy2()));// тут пока под вопрос где находится карттинка
        }

    }//запись все картинок

    public void setWidth1PixelPDF(int sumPoint){
        width1PixelPDF = ((double)sumPoint) / ((double)PageSize.A4.getWidth());
    }

    public int[] getEndPage (){
        if(numberRow.length==0) {
            numberRow = sheet.getRowBreaks();
        }
        return numberRow;

    }//последние строки на странице
    public int sumPage() {
        return numberRow.length + 1;
    }
    public int getNumberRowList(){

        return finishNumberOfRow-startNumberOfRow;
    }//Количество строк для печати без учета скрытых
    public int getNumberRowListWithoutHidden(){
        int sum =0;
        for(int i=startNumberOfRow;i<finishNumberOfRow;i++){
            if(!sheet.getRow(i).getZeroHeight()){
                sum++;
            }

        }
        return sum;
    }//Количество строк для печати c учета скрытых
    public int getNumberColumnList(){

        return finishNumberOfColumns-startNumberOfColumns;
    }//Количество столбцов для печати без учета скрытых
    public int getNumberColumnListWithoutHidden(){
        int sum =0;
        for(int i=startNumberOfColumns;i<finishNumberOfColumns;i++){
            if(!sheet.isColumnHidden(i)){
                sum++;
            }

        }
        return sum;
    }//Количество столбцов для печати c учета скрытых
    public XSSFSheet getSheet() {
        return sheet;
    }

    public ArrayList<Integer> getColumnsNumbersNotHidden() {
        return columnsNumbersNotHidden;
    }

    public ArrayList<Integer> getRowsNumbersNotHidden() {
        return rowsNumbersNotHidden;
    }

    public int widthExcel() {
        XSSFRow row = null;
        int max = 0;
        for (int rn = 0; rn < sheet.getLastRowNum(); rn++) {
            row = sheet.getRow(rn);
            if (row.getLastCellNum() > max) {
                max = row.getLastCellNum();
            }
        }
        int summ = 0;
        for (int cn = 0; cn < max; cn++) {
            int a = sheet.getColumnWidth(cn);
            summ = summ + a;
                    // узнаем ширину листа
        }
        return summ;
    }//Ширина листа в точках
    public String  readCell (int x,int y){
            XSSFRow row = sheet.getRow(y);
            if (row != null||!row.getZeroHeight()) {

                    if(!sheet.isColumnHidden(x)){

                    XSSFCell cell = row.getCell(x);
                    if (cell == null) {

                        return " ";
                    } else {
                        if (cell.getCellTypeEnum() == CellType.STRING) {
                            return cell.getStringCellValue();
                        }
                        else if(cell.getCellTypeEnum() == CellType.FORMULA){
                            try {
                                Double.parseDouble(cell.getCellStyle().getDataFormatString());
                                return new DecimalFormat("#"+ cell.getCellStyle().getDataFormatString()).format(cell.getNumericCellValue());
                            }
                            catch (Exception e){
                                DateTimeFormatter dateFmt = null;
//                                if (cell.getCellStyle().getDataFormat() == 14) {
//                                    dateFmt = DateTimeFormatter.ofPattern("dd.MM.yyyy");
//                                } else { //other data formats with explicit formatting
//                                    dateFmt = DateTimeFormatter.ofPattern (cell.getCellStyle().getDataFormatString());
                                    dateFmt = DateTimeFormatter.ofPattern("dd.MM.yyyy");
//                                }
                                LocalDate date = (cell.getDateCellValue().toInstant().atZone(ZoneId.systemDefault())
                                        .toLocalDate());

                                return date.format(dateFmt);
                            }



                        }
                        else if(cell.getCellTypeEnum()==CellType.NUMERIC){
                            return ""+ (int)cell.getNumericCellValue();
                        }

                        CellStyle style = cell.getCellStyle();
                        XSSFCellStyle St = cell.getCellStyle();
                        style.getRotation();
//                        System.out.print(St.getFontIndex() + " ");

//                        System.out.print(x + " " + y + " Адрес");
//                        System.out.print(" " + sheet.getColumnWidth(x) + " Ширина");
//                        System.out.println(" " + row.getHeight() + " Высота");  // тута точки
                        }
                    }

        }
        return " ";
    }//Возврощает значие ячейки
    public XSSFCellStyle getStyleCell(int x,int y){
        XSSFRow row = sheet.getRow(y);
        return row.getCell(x).getCellStyle();
    } //Узнаю стиль ячейки

    public HashMap<Picture,AddressForPicture> getPictures() {
        return pictures;
    }

    public float[] getArrayWidthColumn(){
        float[] floats = new float[columnsNumbersNotHidden.size()];
        for(int i=0;i<columnsNumbersNotHidden.size();i++){
            floats[i] = sheet.getColumnWidth(columnsNumbersNotHidden.get(i))/42.666f;
        }
        return floats;
    }//узнаю общую ширину столбцов в пикселях пропустив скрытые столбцы
    public float getWidthPage(){
        float f = 0;
        for (int i :columnsNumbersNotHidden){
            if(i==finishNumberOfColumns)break;
            f = f + sheet.getColumnWidth(i)/42.666f;
        }
        return f;
    } //Ширина страницы в пикселях
    public float getHeightPage(){
        float f =0;
        for (int i :rowsNumbersNotHidden){
            if(i==getEndPage()[0]) break;
            try {
                f = f + sheet.getRow(i).getHeight()/15;
            }
            catch (Exception e){}
        }
        return f;
    }
    public int[] getFirstCell(int countRegion){
        int[] value = new int[4];
        value[0] = cellRangeAddressList.get(countRegion).getFirstColumn();
        value[1] = cellRangeAddressList.get(countRegion).getFirstRow();
        value[2] = cellRangeAddressList.get(countRegion).getLastColumn();
        value[3] = cellRangeAddressList.get(countRegion).getLastRow();
        return value;
    }
    public void setListMergedRegion(){
        cellRangeAddressList = new ArrayList<>(sheet.getMergedRegions());
        Collections.sort(cellRangeAddressList, new Comparator<CellRangeAddress>() {
            @Override
            public int compare(CellRangeAddress o1, CellRangeAddress o2) {
                if(o2.getFirstRow()>o1.getFirstRow()){
                    return -1;
                }
                else if(o2.getFirstRow()<o1.getFirstRow())return 1;
                else {
                    if(o2.getFirstColumn()>o1.getFirstColumn()){
                        return -1;
                    }
                    if(o2.getFirstColumn()<o1.getFirstColumn()){
                        return 1;
                    }
                    return 0;
                }
        }
        });

    }//Объедененые ячейки
    public boolean isNotHidden(int x,int y){
        if(rowsNumbersNotHidden.contains(y)&&columnsNumbersNotHidden.contains(x)){
        return true;
        }
        return false;
    }//Возращает true если строчка и столбик не скрыт
    public boolean isNotHiddenX(int x){
        if(columnsNumbersNotHidden.contains(x)){
            return true;
        }
        return false;
    }
    public boolean isNotHiddenY(int y){
        if(rowsNumbersNotHidden.contains(y)){
            return true;
        }
        return false;
    }

}

