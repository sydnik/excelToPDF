package test;

import org.apache.poi.ss.usermodel.PictureData;


public class Picture {
    private double width;
    private double heigth;
    private PictureData picture;
    private String name;

    public Picture(String name,PictureData picture,double width, double heigth) {
        this.width = width;
        this.heigth = heigth;
        this.picture = picture;
        this.name = name;
    }

    public double getWidth() {
        return width;
    }

    public double getHeigth() {
        return heigth;
    }

    public PictureData getPicture() {
        return picture;
    }
}
