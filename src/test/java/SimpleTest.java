import java.io.IOException;
import java.util.ArrayList;

public class SimpleTest {
    public static void main(String[] args) throws IOException {
        ArrayList<String> dataVelue = dataDriven.getData("Purchase");
        for (String data:dataVelue
             ) {
            System.out.println(dataVelue);
        }

    }
}
