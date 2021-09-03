import java.util.ArrayList;
import java.util.List;

public class Report1Row implements ReportRow {
    private ArrayList<Object> body;
    public static final ArrayList<String> headers = new ArrayList<>(List.of(
            "Adı",
            "Soyadı",
            "Atasının adı",
            "FİN kod",
            "Şəhər",
            "Rayon",
            "Ünvan",
            "Təvəllüd",
            "Telefon",
            "İstifadəçi kateqoriyası",
            "Qeydiyyat tarixi",
            "Cari reytinqi",
            "Peşə",
            "Spesifikasiya",
            "Təcrübə (il)",
            "İstənilən əmək haqqı",
            "İcra olunmuş sifarişlər sayı"
    ));

    Report1Row(){}

    Report1Row(String name, String surname, String dadName, String fin, String city, String district, String address,
               String birthdate, String telephone, String userCategory, String registrationDate, double rating, String profession,
               String specification, int experience, String salary, int numberOfCompletedOrders){

        ArrayList<Object> body = new ArrayList<>();
        body.add(0, name);
        body.add(1, surname);
        body.add(2, dadName);
        body.add(3, fin);
        body.add(4, city);
        body.add(5, district);
        body.add(6, address);
        body.add(7, birthdate);
        body.add(8, telephone);
        body.add(9, userCategory);
        body.add(10, registrationDate);
        body.add(11, rating);
        body.add(12, profession);
        body.add(13, specification);
        body.add(14, experience);
        body.add(15, salary);
        body.add(16, numberOfCompletedOrders);
        this.body = body;
    }

    public ArrayList<Object> getBody() {
        return body;
    }

    public void setBody(ArrayList<Object> body) {
        this.body = body;
    }
}
