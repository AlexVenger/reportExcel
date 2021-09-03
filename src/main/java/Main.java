import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class Main {
    public static void main(String[] args) throws IOException {
//        Customer customer1 = new Customer();
//        Customer customer2 = new Customer();
//        List<Customer> customers = new ArrayList<>();
//        customers.add(customer1);
//        customers.add(customer2);
//
//        IReport report = new IReport();
//
//        report.getRows().addAll(customers);
//        report.generate();
//        Report1Row row1 = new Report1Row("Chungus", "Yeetus", "Brutus", "696969", "Wacky City", "Pingas", "IDK st. h.31 ap.69",
//                "01/02/1666", "(666) 420-69-69", "E", "01/09/2021", 6.9,
//                "Yeeter", "BDSM", 69, "10000000", 666);
//        Report1Row row2 = new Report1Row("Chingus", "Dongus", "Yootus", "000000", "PP City", "TT District", "Yeet st. h.0 ap.-1",
//                "00/00/1666", "(666) 420-14-88", "O", "05/04/2023", 3.6,
//                "Yanker", "ABC", 69, "-20", 777);
//        ArrayList<ReportRow> report = new ArrayList<>();
//        report.add(row1);
//        report.add(row2);

        ArrayList<String> headers = new ArrayList<>(List.of(
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
        ArrayList<Object> row1 = new ArrayList<>(List.of(
                "Chungus", "Yeetus", "Brutus", "696969", "Wacky City", "Pingas", "IDK st. h.31 ap.69",
                "01/02/1666", "(666) 420-69-69", "E", "01/09/2021", 6.9,
                "Yeeter", "BDSM", 69, "10000000", 666
        ));
        ArrayList<Object> row2 = new ArrayList<>(List.of(
                "Chingus", "Dongus", "Yootus", "000000", "PP City", "TT District", "Yeet st. h.0 ap.-1",
                "00/00/1666", "(666) 420-14-88", "O", "05/04/2023", 3.6,
                "Yanker", "ABC", 69, "-20", 777
        ));
        ArrayList<ArrayList<Object>> bodies = new ArrayList<>(List.of(row1, row2));

        byte[] chungus = ExcelMaker.createExcel(headers, bodies);
        File file = new File("C:\\Users\\Igor\\Desktop\\JavaProjects\\ReportExcelGenerationAttempt\\src\\main\\chungus.xlsx");
        InputStream in = new ByteArrayInputStream(chungus);
        OutputStream out = new FileOutputStream(file);
        int len;
        while((len = in.read(chungus)) != -1) {
            out.write(chungus, 0, len);
        }
        out.flush();
        in.close();
        out.close();
    }
}
