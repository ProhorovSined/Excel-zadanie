import java.util.ArrayList;

public class Record {
    private final ArrayList<Object> data = new ArrayList<>();

    public String getStringValue(int index) {
        return String.valueOf(data.get(index));
    }

    public int getIntegerValue(int index) {
        return Integer.parseInt(data.get(index).toString());
    }

    public void addValue(Object value) {
        data.add(value);
    }

    @Override
    public String toString() {
        StringBuilder stringBuilder = new StringBuilder();
        for (Object o : data) {
            stringBuilder.append(String.format(o.toString() + " "));
        }
        return stringBuilder.toString();
    }
}
