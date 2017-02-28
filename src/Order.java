/**
 * Created by pavellsda on 07.02.17.
 */
public class Order {
    private String name;
    private int count = 0;
    private String attribute=null;
    private String systemName;

    public Order(String name, int count){
        this.name = name;
        this.count = count;
    }

    public int getCount() {
        return count;
    }

    public void incCount(int count) {

        this.count+=count;
    }

    public void setCount(int count) {
        this.count = count;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getName() {
        return name;
    }

    public String getAttribute() {
        return attribute;
    }

    public void setAttribute(String attribute) {
        this.attribute = attribute;
    }

    public void setSystemName(String systemName) {
        this.systemName = systemName;
    }

    public String getSystemName() {
        return systemName;
    }

    public String getAtName() {

        return name;
    }

    public int getAtCount() {

        return count;
    }


}
