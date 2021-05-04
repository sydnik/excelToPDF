package test;

public class AddressForPicture {
    private double startCellX;
    private double startCellY;
    private double startPixelinCellX;
    private double startPixelinCellY;
    private double endCellX;
    private double endCellY;
    private double endPixelCellX;
    private double endPixelCellY;
    private double start;

    public AddressForPicture(double startCellX, double startCellY, double startPixelinCellX, double startPixelinCellY, double endCellX, double endCellY, double endPixelCellX, double endPixelCellY) {
        this.startCellX = startCellX;
        this.startCellY = startCellY;
        this.startPixelinCellX = startPixelinCellX/12700D;
        this.startPixelinCellY = startPixelinCellY/12700D;
        this.endCellX = endCellX;
        this.endCellY = endCellY;
        this.endPixelCellX = endPixelCellX/12700D;
        this.endPixelCellY = endPixelCellY/12700D;
    }

    public double getStartCellX() {
        return startCellX;
    }

    public double getStartCellY() {
        return startCellY;
    }

    public double getStartPixelinCellX() {
        return startPixelinCellX;
    }

    public double getStartPixelinCellY() {
        return startPixelinCellY;
    }

    public double getEndCellX() {
        return endCellX;
    }

    public double getEndCellY() {
        return endCellY;
    }

    public double getEndPixelCellX() {
        return endPixelCellX;
    }

    public double getEndPixelCellY() {
        return endPixelCellY;
    }
}
