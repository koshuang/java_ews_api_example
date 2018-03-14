package com.mycompany.app;

import java.net.URI;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.service.ConflictResolutionMode;
import microsoft.exchange.webservices.data.core.enumeration.service.DeleteMode;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.service.folder.CalendarFolder;
import microsoft.exchange.webservices.data.core.service.item.Appointment;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.AppointmentSchema;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.ItemId;
import microsoft.exchange.webservices.data.search.CalendarView;
import microsoft.exchange.webservices.data.search.FindItemsResults;
/**
 * Hello world!
 *
 */
public class App
{
    private static ExchangeService service = null;

    public static void main( String[] args ) throws Exception
    {
        System.out.println( "Hello World2!" );

        service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
        service.setUrl(new URI("https://webmail12.sherwebcloud.com/ews/Exchange.asmx"));
        ExchangeCredentials credentials = new WebCredentials("kosh@test.com", "test");
        service.setCredentials(credentials);
        getAppointment();
    }

    public static void getAppointment() throws Exception {
        String appointmentItemId = "AAMkAGIyMjgxOTBhLTBlNDktNDBjMS04MzhmLTk1ZmZjZDMwNTdlZABGAAAAAAB9kX0vybIWQ49bVQ4OMCu3BwC4okUlG3aRSIhpm7wBBJLhAAAAAAENAAC4okUlG3aRSIhpm7wBBJLhAACxtdz1AAA=";
        Item item = Item.bind(service, new ItemId(appointmentItemId));
        // Map appointmentData = new HashMap();
        Appointment appointment = (Appointment)item;
        // appointment.setSubject("test2");
        // appointmentData = readAppointment(appointment);
        // appointment.update(ConflictResolutionMode.AutoResolve);
        // System.out.println(appointmentData.get("appointmentSubject"));
        appointment.delete(DeleteMode.HardDelete);
    }

    public static List readAppointments() {
        List apntmtDataList = new ArrayList();
        Calendar now = Calendar.getInstance();
        Date startDate = Calendar.getInstance().getTime();
        now.add(Calendar.DATE, 30);
        Date endDate = now.getTime();
        try {
            CalendarFolder calendarFolder = CalendarFolder.bind(service, WellKnownFolderName.Calendar, new PropertySet());
            CalendarView cView = new CalendarView(startDate, endDate, 5);
            cView.setPropertySet(new PropertySet(AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End));// we can set other properties
            // as well depending upon our need.
            FindItemsResults appointments = calendarFolder.findAppointments(cView);
            int i = 1;
            List<Appointment> appList = appointments.getItems();
            for (Appointment appointment : appList) {
                System.out.println("\nAPPOINTMENT #" + (i++) + ":");
                Map appointmentData = new HashMap();
                appointmentData = readAppointment(appointment);

                System.out.println("subject : " + appointmentData);
                System.out.println("subject : " + appointmentData.get("appointmentSubject"));
                System.out.println("On : " + appointmentData.get("appointmentStartTime"));
                apntmtDataList.add(appointmentData);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return apntmtDataList;
    }

    /**
     * Reading one appointment at a time. Using Appointment ID of the email.
     * Creating a message data map as a return value.
     */
    public static Map readAppointment(Appointment appointment) {
        Map appointmentData = new HashMap();
        try {
            appointmentData.put("appointmentItemId", appointment.getId().toString());
            appointmentData.put("appointmentSubject", appointment.getSubject());
            appointmentData.put("appointmentStartTime", appointment.getStart() + "");
            appointmentData.put("appointmentEndTime", appointment.getEnd() + "");
            //appointmentData.put("appointmentBody", appointment.getBody().toString());
        } catch (ServiceLocalException e) {
            e.printStackTrace();
        }
        return appointmentData;
    }
}
