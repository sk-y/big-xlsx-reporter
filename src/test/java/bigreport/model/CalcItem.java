package bigreport.model;

public class CalcItem {
    private double price;
    private int count;
    private String name;

    public CalcItem(String name, double price, int count) {
        this.name=name;
        this.price = price;
        this.count = count;
    }

    public double getSum(){
        return price*count;
    }

    public double getPrice() {
        return price;
    }

    public void setPrice(long price) {
        this.price = price;
    }

    public int getCount() {
        return count;
    }

    public void setCount(int count) {
        this.count = count;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }
}
