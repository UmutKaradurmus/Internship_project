package source;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.SQLException;
import java.util.Properties;
import java.util.Timer;
import java.util.TimerTask;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.Part;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Store;
import javax.mail.internet.MimeBodyPart;

public class Gmail_Driver extends Sevkiyat{

    public void gmail() throws MessagingException, FileNotFoundException, IOException, SQLException {
        String host = "pop.gmail.com";
        final String user = "huaweiproit@gmail.com";
        final String password = "*****"; 

        Properties properties = System.getProperties();
        properties.setProperty("mail.smtp.host", host);
        properties.put("mail.smtp.auth", "true");

        Session session = Session.getDefaultInstance(properties,
                new javax.mail.Authenticator() {
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(user, password);
            }
        });

        Store store = session.getStore("pop3s");
        store.connect(host, user, password);

        Folder folder = store.getFolder("INBOX");

        folder.open(Folder.READ_WRITE);

        Message[] message = folder.getMessages();
        int a = 1;
        for (int i = message.length - 1; i > message.length-20; i--) {
            if (i<0) {
                break;
            }
            if (message[i].getFrom()[0].toString().equals("=?UTF-8?Q?umut_Karadurmu=C5=9F?= <huaweiproit@gmail.com>")
                    && (message[i].getSubject().equals("Sevkiyat"))) {
                System.out.println("Message " + (a));
                System.out.println("From : " + message[i].getFrom()[0]);
                System.out.println("Subject : " + message[i].getSubject());
                System.out.println("Sent Date : " + message[i].getSentDate());

                System.out.println();
                a++;
                Multipart multipart = (Multipart) message[i].getContent();

                for (int j = 0; j < multipart.getCount(); j++) {
                    MimeBodyPart part = (MimeBodyPart) multipart.getBodyPart(j);
                    if (Part.ATTACHMENT.equalsIgnoreCase(part.getDisposition())) {
                       System.out.println("Dosya Yükleniyor");
                        String destFilePath = "C:\\excel_test\\" + part.getFileName();
                      
                        FileOutputStream output = new FileOutputStream(destFilePath);

                        InputStream input = part.getInputStream();

                        byte[] buffer = new byte[4096];

                        int byteRead;

                        while ((byteRead = input.read(buffer)) != -1) {
                            output.write(buffer, 0, byteRead);
                        }
                          System.out.println("Dosya Yükllendi");
                        output.close();

                    }
                }

            } else {
                a++;
            }

        }

        folder.close(true);
        store.close();
        Sevkiyat s = new Sevkiyat();
        System.out.println("Klasör İşlemleri Başlıyor");
        s.basla();
    }

    public static void main(String[] args) throws MessagingException, IOException, FileNotFoundException, SQLException {
     
                 Timer myTimer=new Timer();
             TimerTask gorev =new TimerTask() {
 
                    @Override
                    public void run() {
                       Gmail_Driver j5 = new Gmail_Driver();
                        try {  
                            j5.gmail();
                        } catch (MessagingException | IOException | SQLException ex) {
                            Logger.getLogger(Gmail_Driver.class.getName()).log(Level.SEVERE, null, ex);
                        }
                          
                                 
                    }
             };
 
             myTimer.schedule(gorev,0,600000);
       }
    
    }


    

