public class Event
{
  private String task;
  private double hours;
  private String starts;
  private String ends;
  	
  public Event(String task, double hours, String starts, String ends)
  {
  	this.task 	= task;
  	this.hours 	= hours;
  	this.starts = starts;
  	this.ends	= ends;
  }
  
  public String getTask()
  {
  	return this.task;
  }
  public double getHours()
  {
  	return this.hours;
  }
  public String getStarts()
  {
  	return this.starts;
  }
  public String getEnds()
  {
  	return this.ends;
  }
}