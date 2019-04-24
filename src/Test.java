
import java.net.MalformedURLException;

/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
/**
 *
 * @author carlogon
 */
public class Test {

    public static void main(String[] args) throws MalformedURLException {
       

        if (Convert.transformXlsx2Xls("C:/Users/carlogon/Desktop/test/archivoTest.xlsx")) {
            System.out.println("Funciono");
        } else {
            System.out.println("Error");
        }

    }
}
