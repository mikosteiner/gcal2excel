/* 
 * GCal2Excel.java
 * @author anupom
 */
 
import com.google.gdata.client.Query;
import com.google.gdata.client.calendar.CalendarQuery;
import com.google.gdata.client.calendar.CalendarService;
import com.google.gdata.data.DateTime;
import com.google.gdata.data.calendar.CalendarEventEntry;
import com.google.gdata.data.calendar.CalendarEventFeed;
import com.google.gdata.data.extensions.When;
import com.google.gdata.util.ServiceException;

import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.Vector;
import java.io.File;
import java.util.Date;
import java.util.Iterator;
import java.util.ListIterator;

import jxl.*;
import jxl.write.*; 
 
/**
 * Creates an excel file from google calendar
 * using Google Calendar API and JExcel API
 */
public class Converter 
{
  // The base URL for a user's calendar metafeed (needs a username appended).
  private final String METAFEED_URL_BASE = 
      "http://www.google.com/calendar/feeds/";
  
  private final String SINGLE_FEED_URL_SUFFIX = "/private/full";
  
  // The URL for the event feed of the specified user's primary calendar.
  // (e.g. http://www.googe.com/feeds/calendar/calendar-id/private/full)
  private static URL singleFeedUrl = null;

  
  private final long MILISECONDS_IN_HOUR = 60*60*1000;
 
  private String userName;
  private String userPassword;
  private String calendarId;
  
 public Converter(String userName, String userPassword, String calendarId)
 {
  	 this.userName 	  = userName ;
	 this.userPassword = userPassword;
	 this.calendarId   = calendarId;
  }
  
  public boolean convert(String starts, String ends)
  {
  	String fileName = this.userName + " TimeSheet From " + starts + " To " + ends + ".xls";	      		
    Vector<Event> events = null;
    
    events = getCalendarData(this.userName, this.userPassword, this.calendarId, starts, ends);
    
    if(events==null)
    	return false;
    
    return writeExcel(events, fileName);
  }
  
  private boolean writeExcel(Vector<Event> events, String fileName)
  {
  	try
	{
		WritableWorkbook workbook = Workbook.createWorkbook(new File(fileName));
		WritableSheet sheet = workbook.createSheet("The Sheet", 0);
		
		WritableFont labelFont = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD, true);
		WritableCellFormat labelFormat = new WritableCellFormat (labelFont); 

		Label labelTask = new Label(0, 0, "Task", labelFormat);
		sheet.addCell(labelTask);
		
		Label labelHours = new Label(2, 0, "Hours", labelFormat);
		sheet.addCell(labelHours);
		
		Label labelStart = new Label(4, 0, "Start", labelFormat);
		sheet.addCell(labelStart);
		
		Label labelEnds = new Label(6, 0, "Ends", labelFormat);
		sheet.addCell(labelEnds);
	
		int row = 2;
		ListIterator iterator = events.listIterator();
		while (iterator.hasNext()) 
		{
		    Event event = (Event)iterator.next();
		    
		    jxl.write.Label task = new jxl.write.Label(0, row, event.getTask());
			sheet.addCell(task);
			
		    jxl.write.Number hours = new jxl.write.Number(2, row, event.getHours());
			sheet.addCell(hours);
			
			jxl.write.Label starts = new jxl.write.Label(4, row, event.getStarts());
			sheet.addCell(starts);
			
			jxl.write.Label ends = new jxl.write.Label(6, row, event.getEnds());
			sheet.addCell(ends);
			
			row++;
		}
		
		Label labelTotal = new Label(0, (row+2), "Total Hours Billed: ", labelFormat);
		sheet.addCell(labelTotal);
		
		jxl.write.Formula formula = new jxl.write.Formula( 2, (row+2), "SUM(C3:C"+row+")" );
		sheet.addCell(formula);
		
		
	    workbook.write();
		workbook.close();
		
		return true;
	}
	catch(IOException e)
	{
	  System.err.println("Oh no - can't write the file, may be the file is in use.");
      e.printStackTrace();
      return false;
	}
	catch(jxl.write.WriteException e)
	{
	  System.err.println("Oh no - can't write the excel file.");
      e.printStackTrace();
      return false;
	}
  }
  
  private Vector<Event> getCalendarData(
  											String userName, String userPassword,
  											String calendarId,
  											String starts, String ends
  										)
  {
  	CalendarService myService = new CalendarService("Anupom-Gtalk2Spreadsheet-1");

    Vector<Event> events = null;
	 
  	// Create the necessary URL objects.
  	try 
    {
      singleFeedUrl = new URL(METAFEED_URL_BASE + calendarId + SINGLE_FEED_URL_SUFFIX);
      System.out.println(singleFeedUrl.toString()); 
    }
    catch (MalformedURLException e) 
    {
      // Bad URL
      System.err.println("Uh oh - you've got an invalid URL.");
      e.printStackTrace();
      return null;
    }

    try 
    {
      myService.setUserCredentials(userName, userPassword);
	  
      System.out.println("Date Range query");
      events = dateRangeQuery(myService, 
		      		DateTime.parseDate(starts), 
		      		DateTime.parseDate(ends));
    }
    catch (IOException e) 
    {
      // Communications error
      System.err.println("There was a problem communicating with the service.");
      e.printStackTrace();
    }
    catch (ServiceException e)
    {
      // Server side error
      System.err.println("The server had a problem handling your request.");
      e.printStackTrace();
    }
    
    return events;
  }
  
  /**
   * Prints the titles and start and end times of all the events.
   * 
   * @param service An authenticated CalendarService object.
   * @param startTime Start time (inclusive) of events to print.
   * @param endTime End time (exclusive) of events to print.
   * @throws ServiceException If the service is unable to handle the request.
   * @throws IOException Error communicating with the server.
   */
  private Vector<Event> dateRangeQuery(CalendarService service,
      DateTime startTime, DateTime endTime) throws ServiceException,
      IOException
  {
    
    CalendarQuery myQuery = new CalendarQuery(singleFeedUrl);
    myQuery.setMinimumStartTime(startTime);
	myQuery.setMaximumStartTime(endTime);
	myQuery.addCustomParameter(new Query.CustomParameter("orderby", "starttime"));
	myQuery.addCustomParameter(new Query.CustomParameter("sortorder", "ascending"));
    myQuery.addCustomParameter(new Query.CustomParameter("singleevents", "true"));
    myQuery.addCustomParameter(new Query.CustomParameter("max-results", "10000"));
    
    // Send the request and receive the response:
    CalendarEventFeed resultFeed = service.query(myQuery, CalendarEventFeed.class);
	
    System.out.println("Events from " + startTime.toString() + " to "
        + endTime.toString() + ":");
    System.out.println();
    
    int count = 1;
    Vector<Event> events = new Vector<Event>();
    
    for (int i = 0; i < resultFeed.getEntries().size(); i++) 
    {
      CalendarEventEntry entry = resultFeed.getEntries().get(i);
      
      String title = entry.getTitle().getPlainText();
      java.util.List<When> eventTimes = entry.getTimes();
      
 	  String starts = null;
      String ends = null;
      double duration = 0;
      
      Iterator<When> iterator = eventTimes.iterator();
      
      while(iterator.hasNext()) 
      {
        When when = iterator.next();
      	starts = when.getStartTime().toUiString();
      	ends = when.getEndTime().toUiString();
      	duration = getWhenDiff(when);
      }
      
      
      events.addElement( new Event(title, duration, starts, ends));
      System.out.println("#"+count);
      System.out.println("----------------------");
      System.out.println("\t" + starts);
      System.out.println("\t" + title);
      System.out.println("\t" + ends);
      System.out.println("\t" + duration );
      System.out.println("----------------------");
      count++;
   
   }
   System.out.println();
   
   return events;
  }
  
  private double getWhenDiff(When when)
  {
  	 double hours = 0.0;
  	 
  	 try 
  	 {
  	 	long diff  = when.getEndTime().getValue()-when.getStartTime().getValue();
  	 	if(diff>0)
  	 	{
  	 		hours = (double) diff/(MILISECONDS_IN_HOUR);
  	 	}
  	 }
  	 catch (Exception e)
  	 {
  	 	System.err.println("Problem occured while calcualting time difference");
      	e.printStackTrace();
  	 }
  	 return hours;
  }
}