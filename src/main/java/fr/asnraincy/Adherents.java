package fr.asnraincy;

import com.github.mustachejava.DefaultMustacheFactory;
import com.github.mustachejava.Mustache;
import com.github.mustachejava.MustacheFactory;
import com.google.common.collect.ImmutableList;
import com.google.common.collect.ImmutableMap;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.pdf.AcroFields;
import com.itextpdf.text.pdf.PdfCopy;
import com.itextpdf.text.pdf.PdfName;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfSmartCopy;
import com.itextpdf.text.pdf.PdfStamper;
import org.apache.commons.cli.*;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.commons.io.ByteOrderMark;
import org.apache.commons.io.input.BOMInputStream;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.*;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import java.awt.*;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.time.temporal.ChronoUnit;
import java.util.*;
import java.util.List;

/**
 * Created by cjolif on 08/06/2016.
 */
public class Adherents {

    private static ClassLoader CLASS_LOADER = Adherents.class.getClassLoader();

    private static Map<String, String> SWIM_GROUPS = ImmutableMap.<String, String>builder()
            .put("1", "Ecole de Natation 1 (lundi) (230€)")
            .put("2", "Ecole de Natation 1 (vendredi) (230€)")
            .put("3", "Ecole de Natation 2 (lundi) (230€)")
            .put("4", "Ecole de Natation 2 (vendredi) (230€)")
            .put("5", "Ecole de Natation 3 (230€)")
            .put("6", "Avenirs (260€)")
            .put("7", "Jeunes (260€)")
            .put("8", "Juniors (260€)")
            .put("9", "PréAdo (230€)")
            .put("10", "Ado (230€)")
            .put("11", "Prépa bac (230€)")
            .put("12", "Adulte (230€)")
            .put("13", "Maitre (260€)")
            .put("14", "Officiel (15€)")
            .build();

    // « Avenirs » (6), « Jeunes » (7), « Juniors » (8), « Maitres » (13)
    private static ImmutableList<String> COMPET_GROUPS = ImmutableList.<String>builder()
        .add("6", "7", "8", "13").build();

    private static Map<String, Integer> PRICE = ImmutableMap.<String, Integer>builder()
            .put("1", 230)
            .put("2", 230)
            .put("3", 230)
            .put("4", 230)
            .put("5", 230)
            .put("6", 260)
            .put("7", 260)
            .put("8", 260)
            .put("9", 230)
            .put("10", 230)
            .put("11", 230)
            .put("12", 230)
            .put("13", 230)
            .put("14", 15)
            .build();

    private static Map<String, String> GROUPS = ImmutableMap.<String, String>builder()
            .put("0", "Jardin aquatique")
            .put("1", "Avenir 1")
            .put("2", "Avenir 2")
            .put("3", "Poussin 1")
            .put("4", "Poussin 2")
            .put("5", "Benjamin")
            .put("6", "Compétition")
            .put("7", "Jeune")
            .put("8", "Ado")
            .put("9", "Prépa bac")
            .put("10", "Adulte")
            .put("11", "Handicap")
            .put("12", "Ecole de nage")
            .put("13", "Aquaphobie")
            .build();

    private static Map<String, String> CITIES = ImmutableMap.<String, String>builder()
            .put("93600", "Aulnay-sous-Bois")
            .put("93390", "Clichy-sous-Bois")
            .put("93370", "Montfermeil")
            .put("93340", "Le Raincy")
            .put("93320", "Les Pavillons-sous-Bois")
            .put("93250", "Villemomble")
            .put("93220", "Gagny")
            .put("93190", "Livry-Gargan")
            .put("93140", "Bondy")
            .put("93110", "Rosny-sous-Bois")
            .build();

    private static String getCity(String postcode, String city) {
        return CITIES.get(postcode) != null ? CITIES.get(postcode) : city;
    }

    static class RegistrationWriter extends Thread
    {
        PipedOutputStream pos;
        CSVRecord row;

        public RegistrationWriter(PipedOutputStream pos, CSVRecord row)
        {
            this.pos = pos;
            this.row = row;
        }

        public void run() {
            try {
                PdfReader reader = new PdfReader("formulaire_inscription.pdf");
                PdfStamper stamper = new PdfStamper(reader, pos, '\0', true);
                AcroFields form = stamper.getAcroFields();
                form.setField("Groupe", SWIM_GROUPS.get(row.get("swim_group")));
                form.setField("Adh_Nom", row.get("lastname"), row.get("lastname"), true);
                form.setField("Adh_Prenom", row.get("firstname"));
                form.setField("Sexe", row.get("gender").equals("male") ? "2" : "1");
                form.setField("Adh_Date_de_naissance",
                    row.get("birth_day") + "/" + row.get("birth_month") + "/" + row.get("birth_year"));
                form.setField("Adh_Nationalite", row.get("country"));
                form.setField("Adh_Adresse", row.get("address"));
                form.setField("Adh_Code_Postal", row.get("postcode"));
                form.setField("Adh_Ville", getCity(row.get("postcode"), row.get("city")));
                form.setField("Adh_Tel_Dom", fixPhone(row.get("phone")));
                form.setField("Adh_Mobile", fixPhone(row.get("mobile")));
                form.setField("Adh_Mail", row.get("email"));
                form.setField("Leg_Nom", row.get("lastname_leg"));
                form.setField("Leg_Prenom", row.get("firstname_leg"));
                form.setField("Leg_Lien_Parente", row.get("parent_leg"));
                form.setField("Leg_Adresse", row.get("address_leg"));
                form.setField("Leg_Code_Postal", row.get("postcode_leg"));
                form.setField("Leg_Ville", getCity(row.get("postcode_leg"), row.get("city_leg")));
                form.setField("Leg_Tel_Dom", fixPhone(row.get("phone_leg")));
                form.setField("Leg_Mobile", fixPhone(row.get("mobile_leg")));
                form.setField("Leg_Mail", row.get("email_leg"));
                form.setField("Urg_Nom", row.get("name_urg"));
                form.setField("Urg_Tel1", fixPhone(row.get("phone_urg")));
                form.setField("Urg_Tel2", fixPhone(row.get("mobile_urg")));
                if (!row.get("lastname_leg").equals("")) {
                    form.setField("Sousigne", row.get("lastname_leg") + " " + row.get("firstname_leg"));
                } else {
                    form.setField("Sousigne", row.get("lastname") + " " + row.get("firstname"));
                }
                form.setField("Autorisation_intervention", "autorise");
                int price = PRICE.get(row.get("swim_group")).intValue();
                form.setField("Adhesion", String.valueOf(price));
                form.setField("Prix", String.valueOf(price));
                stamper.close();
                reader.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    static class LicenceWriter extends Thread
    {
        PipedOutputStream pos;
        CSVRecord row;
        long years;
        boolean renew;

        public LicenceWriter(PipedOutputStream pos, CSVRecord row, long years, boolean renew)
        {
            this.pos = pos;
            this.row = row;
            this.years = years;
            this.renew = renew;
        }

        public void run() {
            try {
                PdfReader reader = new PdfReader(CLASS_LOADER.getResource("formulaire_licence.pdf"));
                if (renew) {
                    reader.selectPages(ImmutableList.of(1, 6, 7));
                } else {
                    reader.selectPages("1");
                }
                PdfStamper stamper = new PdfStamper(reader, pos);
                AcroFields form = stamper.getAcroFields();
                if (renew) {
                    form.setField("2", "Oui");
                } else {
                    form.setField("1", "Oui");
                }
                form.setField("nom", row.get("lastname"));
                form.setField("Prenom", row.get("firstname"));
                form.setField("Nationalité", row.get("country"));
                form.setField("h/f", row.get("gender").equals("male") ? "H" : "F");
                form.setField("Date de naissance", row.get("birth_day") + '/'
                    + row.get("birth_month") + '/' + row.get("birth_year"));
                form.setField("adresse", row.get("address"));
                form.setField("code postal", row.get("postcode"));
                form.setField("Ville", getCity(row.get("postcode"), row.get("city")));
                // split mail
                String email = row.get("email");
                int ar = email.lastIndexOf("@");
                form.setField("Email", email.substring(0, ar));
                form.setField("mail 2", email.substring(ar + 1));
                form.setField("Tel 1", row.get("phone"));
                form.setField("Text18", row.get("mobile"));
                // don't want additional info
                form.setField("5", "Oui");
                if (years > 6) {
                    if (COMPET_GROUPS.contains(row.get("swim_group"))) {
                        form.setField("6", "Oui");
                    } else {
                        // natation
                        form.setField("12", "Oui");
                    }
                } else {
                    // eveil
                    form.setField("19", "Oui");
                }
                // assurance de base
                form.setField("44", "Oui");
                // assurance étendue => non (case 48 = non)
                form.setField("48", "Oui");
                // autorise les prélèvements
                form.setField("42", "Oui");
                stamper.close();
                reader.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    public static void main(String[] args) throws IOException, DocumentException, ParseException, UnsupportedFlavorException {
        Options options = new Options();
        options.addOption("f", "file", true, "specify CSV file to be read");
        options.addOption("m", "mail", false, "specify notification mails must be sent");
        options.addOption("p", "properties", true, "specify a mail config property file");
        CommandLineParser cliparser = new DefaultParser();
        CommandLine cmd = cliparser.parse(options, args);

        Properties config = new Properties();
        InputStream is = CLASS_LOADER.getResourceAsStream("config.properties");
        if (is != null) {
            config.load(is);
        } else {
            throw new FileNotFoundException("property file config.properties not found in the classpath");
        }
        if (cmd.hasOption("p")) {
            config.load(new FileReader(cmd.getOptionValue("p")));
        }
        // get the property value and print it out
        final String username = config.getProperty("mail.user");
        final String password = config.getProperty("mail.password");

        config.put("mail.user", "");
        config.put("mail.password", "");

        Session session = Session.getInstance(config,
                new Authenticator() {
                    protected PasswordAuthentication getPasswordAuthentication() {
                        return new PasswordAuthentication(username, password);
                    }
                });

        boolean file = cmd.hasOption("f");
        CSVParser parser;
        if (!file) {
            Toolkit toolkit = Toolkit.getDefaultToolkit();
            Clipboard clipboard = toolkit.getSystemClipboard();
            String result = (String) clipboard.getData(DataFlavor.stringFlavor);
            parser = CSVParser.parse(result, CSVFormat.DEFAULT.withFirstRecordAsHeader());
        } else {
            InputStream inputStream = new FileInputStream(new File(cmd.getOptionValue("f")));
            BOMInputStream bOMInputStream = new BOMInputStream(inputStream);
            ByteOrderMark bom = bOMInputStream.getBOM();
            String charsetName = bom == null ? "UTF-8" : bom.getCharsetName();
            parser = new CSVParser(new InputStreamReader(new BufferedInputStream(bOMInputStream), charsetName),
                    CSVFormat.DEFAULT.withFirstRecordAsHeader());
        }
        for (CSVRecord row : parser) {
            generatePDFs(row, cmd.hasOption("m"), session);
        }
    }

    private static String fixPhone(String phone) {
        String r = phone.replaceAll("\\s","");
        if (r.indexOf("+33") == 0) {
            return "0" + r.substring(3);
        }
        if (r.indexOf("0033") == 0) {
            return "0" + r.substring(4);
        }
        return r;
    }

    public static void generatePDFs(CSVRecord row, boolean mail, Session session) throws IOException, DocumentException {
        // Unique name
        final String uname = row.get("lastname") + "-" + row.get("firstname") + "-" + row.get("id");

        // Age
        LocalDate birthdate = LocalDate.of(Integer.parseInt(row.get("birth_year")),
                Integer.parseInt(row.get("birth_month")),
                Integer.parseInt(row.get("birth_day")));
        LocalDate ref = LocalDate.of(LocalDate.now().getYear(), 9, 15);
        long years = ChronoUnit.YEARS.between(birthdate, ref);

        // Inscription
        PipedInputStream registrationIn = new PipedInputStream();
        final PipedOutputStream registrationOut = new PipedOutputStream(registrationIn);
        Thread registationThread = new RegistrationWriter(registrationOut, row);
        registationThread.start();

        // Licence
        boolean renew = row.get("renew").equals("1");
        PipedInputStream licenceIn = new PipedInputStream();
        final PipedOutputStream licenceOut = new PipedOutputStream(licenceIn);
        Thread licenceThread = new LicenceWriter(licenceOut, row, years, renew);
        licenceThread.start();

        Document doc = new Document();
        String filename = "FI-" + uname + ".pdf";
        PdfCopy copy = new PdfCopy(doc, new FileOutputStream(filename));
        copy.setMergeFields();
        copy.setFullCompression();
        doc.open();
        PdfReader registration = new PdfReader(registrationIn);
        registration.selectPages("1");
        PdfReader licence = new PdfReader(licenceIn);
        PdfReader rules = new PdfReader("formulaire_inscription.pdf");
        rules.selectPages("2");
        copy.addDocument(registration);
        copy.addDocument(licence);
        copy.addDocument(rules);
        copy.flush();
        doc.close();
        copy.close();

        // Send mail if asked for
        if (mail) {
            try {
                Message mm = new MimeMessage(session);
                List<Address> rec = new ArrayList<>();
                rec.add(new InternetAddress(row.get("email")));
                if (!row.get("email_leg").equals("")) {
                    rec.add(new InternetAddress(row.get("email_leg")));
                }
                mm.setFrom(new InternetAddress("asnr@gmail.com"));
                Address[] recArray = new Address[rec.size()];
                mm.setRecipients(Message.RecipientType.TO, rec.toArray(recArray));
                mm.setSubject("Documents d'inscription de " + row.get("firstname") + " à l'ASNR");

                // Create the message part
                BodyPart messageBodyPart = new MimeBodyPart();

                // Now set the actual message
                // First read the mail template and substitute
                InputStream template;
                if (renew) {
                    template = CLASS_LOADER.getResourceAsStream("mail-renew.txt");
                } else {
                    template = CLASS_LOADER.getResourceAsStream("mail.txt");
                }
                Writer writer = new StringWriter();
                MustacheFactory mf = new DefaultMustacheFactory();
                Mustache mustache = mf.compile(new InputStreamReader(template, StandardCharsets.UTF_8), "mail");
                mustache.execute(writer, ImmutableMap.of("firstname", row.get("firstname")));
                writer.flush();

                messageBodyPart.setText(writer.toString());

                // Create a multipart message
                Multipart multipart = new MimeMultipart();

                // Set text message part
                multipart.addBodyPart(messageBodyPart);

                // Part two is attachment
                messageBodyPart = new MimeBodyPart();
                DataSource source = new FileDataSource(filename);
                messageBodyPart.setDataHandler(new DataHandler(source));
                messageBodyPart.setFileName(filename);
                multipart.addBodyPart(messageBodyPart);

                // Send the complete message parts
                mm.setContent(multipart);

                // Send message
                Transport.send(mm);

                System.out.println("mail sent to: " + Arrays.toString(mm.getRecipients(Message.RecipientType.TO)));
            } catch (MessagingException ex) {
                ex.printStackTrace();
            }

        }
    }
}
