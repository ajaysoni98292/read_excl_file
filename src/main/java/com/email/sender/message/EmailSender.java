package com.email.sender.message;

import com.sendgrid.SendGrid;
import com.sendgrid.SendGridException;

/**
 *
 * @author ajay
 */
public class EmailSender {

    public static void main(String[] args) {

        /* The below line used to pass username and password for the send mail from send grid
        *  Firs argument is the username and second one is password.
        * */
        SendGrid sendgrid = new SendGrid("username", "password");
        SendGrid.Email email = new SendGrid.Email();


        /* From this we can set the email. From to subject and body. SO its very easy to pass email body of any body
        *  in from attribute and receiver will think that sender is abc...
        * */
        email.addTo("abc@gmail.com");
        email.setFrom("test@gmail.com");
        email.setSubject("Hello World");
        email.setText("My first email with SendGrid Java!");
        try {
            SendGrid.Response response = sendgrid.send(email);
            System.out.println(response.getMessage());
        }
        catch (SendGridException e) {
            System.err.println(e);
        }
    }
}
