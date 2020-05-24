package fr.asnraincy;

import com.github.mustachejava.DefaultMustacheFactory;
import com.github.mustachejava.Mustache;
import com.github.mustachejava.MustacheFactory;
import com.google.common.collect.ImmutableList;
import com.google.common.collect.ImmutableMap;
import com.itextpdf.forms.PdfAcroForm;
import com.itextpdf.forms.PdfPageFormCopier;
import com.itextpdf.io.font.FontConstants;
import com.itextpdf.io.font.constants.StandardFonts;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.pdf.PdfDictionary;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfName;
import com.itextpdf.kernel.pdf.PdfReader;
import com.itextpdf.kernel.pdf.PdfWriter;
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
                PdfDocument pdf = new PdfDocument(
                    new PdfReader("formulaire_inscription.pdf"), new PdfWriter(pos));
                PdfAcroForm form = PdfAcroForm.getAcroForm(pdf, false);
                form.getField("Groupe").setValue(SWIM_GROUPS.get(row.get("swim_group")));
                form.getField("Adh_Nom").setValue(row.get("lastname"));
                form.getField("Adh_Prenom").setValue(row.get("firstname"));
                form.getField("Sexe").setValue(row.get("gender").equals("male") ? "2" : "1");
                form.getField("Adh_Date_de_naissance").setValue(
                    row.get("birth_day") + "/" + row.get("birth_month") + "/" + row.get("birth_year"));
                form.getField("Adh_Nationalite").setValue(row.get("country"));
                form.getField("Adh_Adresse").setValue(row.get("address"));
                form.getField("Adh_Code_Postal").setValue(row.get("postcode"));
                form.getField("Adh_Ville").setValue(getCity(row.get("postcode"), row.get("city")));
                form.getField("Adh_Tel_Dom").setValue(fixPhone(row.get("phone")));
                form.getField("Adh_Mobile").setValue(fixPhone(row.get("mobile")));
                form.getField("Adh_Mail").setValue(row.get("email"));
                form.getField("Leg_Nom").setValue(row.get("lastname_leg"));
                form.getField("Leg_Prenom").setValue(row.get("firstname_leg"));
                form.getField("Leg_Lien_Parente").setValue(row.get("parent_leg"));
                form.getField("Leg_Adresse").setValue(row.get("address_leg"));
                form.getField("Leg_Code_Postal").setValue(row.get("postcode_leg"));
                form.getField("Leg_Ville").setValue(getCity(row.get("postcode_leg"), row.get("city_leg")));
                form.getField("Leg_Tel_Dom").setValue(fixPhone(row.get("phone_leg")));
                form.getField("Leg_Mobile").setValue(fixPhone(row.get("mobile_leg")));
                form.getField("Leg_Mail").setValue(row.get("email_leg"));
                form.getField("Urg_Nom").setValue(row.get("name_urg"));
                form.getField("Urg_Tel1").setValue(fixPhone(row.get("phone_urg")));
                form.getField("Urg_Tel2").setValue(fixPhone(row.get("mobile_urg")));
                if (!row.get("lastname_leg").equals("")) {
                    form.getField("Sousigne").setValue(row.get("lastname_leg") + " " + row.get("firstname_leg"));
                } else {
                    form.getField("Sousigne").setValue(row.get("lastname") + " " + row.get("firstname"));
                }
                form.getField("Autorisation_intervention").setValue("autorise");
                int price = PRICE.get(row.get("swim_group")).intValue();
                form.getField("Adhesion").setValue(String.valueOf(price));
                form.getField("Prix").setValue(String.valueOf(price));
                pdf.close();
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
                PdfReader reader = new PdfReader("formulaire_licence.pdf");
                PdfDocument source = new PdfDocument(reader);
                PdfDocument pdf = new PdfDocument(new PdfWriter(pos));
                if (renew) {
                    source.copyPagesTo(ImmutableList.of(1, 6, 7), pdf, new PdfPageFormCopier());
                } else {
                    source.copyPagesTo(1, 1, pdf, new PdfPageFormCopier());
                }
                source.close();
                PdfAcroForm form = PdfAcroForm.getAcroForm(pdf, false);
                if (renew) {
                    form.getField("2").setValue("Oui");
                } else {
                    form.getField("1").setValue("Oui");
                }
                form.getField("nom").setValue(row.get("lastname"));
                form.getField("Prenom").setValue(row.get("firstname"));
                form.getField("Nationalité").setValue(row.get("country"));
                form.getField("h/f").setValue(row.get("gender").equals("male") ? "H" : "F");
                form.getField("Date de naissance").setValue(row.get("birth_day") + '/'
                    + row.get("birth_month") + '/' + row.get("birth_year"));
                form.getField("adresse").setValue(row.get("address"));
                form.getField("code postal").setValue(row.get("postcode"));
                form.getField("Ville").setValue(getCity(row.get("postcode"), row.get("city")));
                // split mail
                String email = row.get("email");
                int ar = email.lastIndexOf("@");
                form.getField("Email").setValue(email.substring(0, ar));
                form.getField("mail 2").setValue(email.substring(ar + 1));
                form.getField("Tel 1").setValue(row.get("phone"));
                form.getField("Text18").setValue(row.get("mobile"));
                // don't want additional info
                form.getField("5").setValue("Oui");
                if (years > 6) {
                    if (COMPET_GROUPS.contains(row.get("swim_group"))) {
                        form.getField("6").setValue("Oui");
                    } else {
                        // natation
                        form.getField("12").setValue("Oui");
                    }
                } else {
                    // eveil
                    form.getField("19").setValue("Oui");
                }
                // assurance de base
                form.getField("44").setValue("Oui");
                // assurance étendue => non (case 48 = non)
                form.getField("48").setValue("Oui");
                // autorise les prélèvements
                // form.getField("42").setValue("Oui");
                if (renew) {
                    // attestation
                    form.getField("nomAttest").setValue(row.get("lastname"));
                    form.getField("PrenomAttest").setValue(row.get("firstname"));
                    form.getField("adresseAttest").setValue(row.get("address"));
                    form.getField("code postalAttest").setValue(row.get("postcode"));
                    form.getField("VilleAttest").setValue(getCity(row.get("postcode"), row.get("city")));
                }
                pdf.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    public static PdfFont findFontInForm(PdfDocument pdfDoc, PdfName fontName) {
        PdfDictionary acroForm = pdfDoc.getCatalog().getPdfObject().getAsDictionary(PdfName.AcroForm);
        if (acroForm == null) {
            return null;
        }
        PdfDictionary dr = acroForm.getAsDictionary(PdfName.DR);
        if (dr == null) {
            return null;
        }
        PdfDictionary font = dr.getAsDictionary(PdfName.Font);
        if (font == null) {
            return null;
        }
        for (PdfName key : font.keySet()) {
            if (key.equals(fontName)) {
                return PdfFontFactory.createFont(font.getAsDictionary(key));
            }
        }
        return null;
    }

    public static void main(String[] args) throws IOException, ParseException, UnsupportedFlavorException {
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

    public static void generatePDFs(CSVRecord row, boolean mail, Session session) throws IOException {
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

        String filename = "FI-" + uname + ".pdf";
        PdfWriter pdfWriter = new PdfWriter(filename);
        pdfWriter.setSmartMode(true);
        PdfDocument doc = new PdfDocument(pdfWriter);
        PdfAcroForm form = PdfAcroForm.getAcroForm(doc, true);
        PdfPageFormCopier copier = new PdfPageFormCopier();
        PdfDocument licence = new PdfDocument(new PdfReader(licenceIn));
        licence.copyPagesTo(1, licence.getNumberOfPages(), doc, copier);
        licence.close();
        PdfDocument registration = new PdfDocument(new PdfReader(registrationIn));
        registration.copyPagesTo(1, 1, doc, 1, copier);
        registration.close();
        PdfDocument rules = new PdfDocument(new PdfReader("formulaire_inscription.pdf"));
        rules.copyPagesTo(2, 2, doc, copier);
        rules.close();
        doc.close();

        // Send mail if asked for
        if (mail) {
            sendMail(row, session, renew, filename);
        }
    }

    private static void sendMail(CSVRecord row, Session session, boolean renew, String filename) {
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

            // Since 2020-2021 we send the reglement interieur separately
            messageBodyPart = new MimeBodyPart();
            source = new URLDataSource(CLASS_LOADER.getResource("reglement_interieur.pdf"));
            messageBodyPart.setDataHandler(new DataHandler(source));
            messageBodyPart.setFileName("Règlement intérieur ASNR.pdf");
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
