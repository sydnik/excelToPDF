package test;


public class DeleteCellAddress {
    private int x;
    private int y;

    public DeleteCellAddress(int x, int y) {
        this.x = x;
        this.y = y;
    }

    public int getX() {
        return x;
    }

    public int getY() {
        return y;
    }

    @Override
    public int hashCode() {
        return x+y;
    }

    public boolean equals(Object obj) {
        DeleteCellAddress obj2 = (DeleteCellAddress) obj;
        if(x==obj2.getX()&&y==obj2.getY()){
            return true;
        }
        return false;
    }
}
