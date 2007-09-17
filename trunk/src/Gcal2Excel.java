import javax.swing.*;
import java.awt.*;
import java.awt.event.*;
import javax.swing.JFrame;
import javax.swing.plaf.metal.MetalLookAndFeel;
import javax.swing.plaf.metal.*;
import java.io.*;
import java.util.Arrays;
import java.util.Date;
import java.text.SimpleDateFormat;
import com.Ostermiller.util.Base64;
import com.toedter.calendar.JDateChooser;

/**
 * Frame to get user input
 * 
 * @author Anupom
 */

public class Gcal2Excel extends JFrame
{

  private JTextField nameTxt;
  private JPasswordField passwordTxt;
  private JTextField calendarIdTxt;
  
  private JDateChooser startsDateChooser;
  private JDateChooser endsDateChooser;
  
  private JLabel nameLbl;
  private JLabel passwordLbl;
  private JLabel calendarIdLbl;
  private JLabel startsLbl;
  private JLabel endsLbl;
  
  private JLabel nameLblHelp;
  private JLabel passwordLblHelp;
  private JLabel calendarIdLblHelp;
  
  private JLabel startsLblHelp;
  private JLabel endsLblHelp;
  
  private JButton okBtn;
  
  private SimpleDateFormat dateFormat;
  
  private final String DATE_FORMAT_PATTERN = "yyyy-MM-dd";
  
  /**
   * The constructor.
   */
  public Gcal2Excel()
  {
    Container container = getContentPane();
    container.setLayout(null);
    
    // Text Fields
    nameTxt = new JTextField();
  	nameTxt.setBounds(86,20,200,25);
    container.add(nameTxt);
    
  	passwordTxt = new JPasswordField();
  	passwordTxt.setBounds(86,50,200,25);
    container.add(passwordTxt);
    
  	calendarIdTxt = new JTextField();
  	calendarIdTxt.setBounds(86,80,200,25);
    container.add(calendarIdTxt);
    
    //Date Choosers
  	startsDateChooser = new JDateChooser(new Date(), DATE_FORMAT_PATTERN);
  	startsDateChooser.setBounds(86,110,200,25);
    container.add(startsDateChooser);
    
  	endsDateChooser = new JDateChooser(new Date(), DATE_FORMAT_PATTERN);
  	endsDateChooser.setBounds(86,140,200,25);
    container.add(endsDateChooser);
    
    // Labels
    nameLbl = new JLabel("UserName");
  	nameLbl.setBounds(10,20,70,25);
    container.add(nameLbl);
    
  	passwordLbl = new JLabel("Password");
  	passwordLbl.setBounds(10,50,70,25);
    container.add(passwordLbl);
    
  	calendarIdLbl = new JLabel("Calendar Id");
  	calendarIdLbl.setBounds(10,80,70,25);
    container.add(calendarIdLbl);
    
  	startsLbl = new JLabel("Starts");
  	startsLbl.setBounds(10,110,70,25);
    container.add(startsLbl);
    
  	endsLbl = new JLabel("Ends");
  	endsLbl.setBounds(10,140,70,25);
    container.add(endsLbl);
  	
  	// Help Labels
    nameLblHelp = new JLabel("example.user@gmail.com");
  	nameLblHelp.setBounds(292,20,160,25);
    nameLblHelp.setForeground(Color.gray);
    container.add(nameLblHelp);
    
  	passwordLblHelp = new JLabel("google password");
  	passwordLblHelp.setBounds(292,50,160,25);
    passwordLblHelp.setForeground(Color.gray);
    container.add(passwordLblHelp);
    
  	calendarIdLblHelp = new JLabel("from calendar settings");
  	calendarIdLblHelp.setBounds(292,80,160,25);
  	calendarIdLblHelp.setForeground(Color.gray);
    container.add(calendarIdLblHelp);
    
  	startsLblHelp = new JLabel("yyyy-mm-dd");
  	startsLblHelp.setBounds(292,110,160,25);
  	startsLblHelp.setForeground(Color.gray);
    container.add(startsLblHelp);
    
  	endsLblHelp = new JLabel("yyyy-mm-dd");
  	endsLblHelp.setBounds(292,140,160,25);
  	endsLblHelp.setForeground(Color.gray);
  	container.add(endsLblHelp);
  	
  	//Button
  	okBtn = new JButton("Create");
    okBtn.setActionCommand("Ok");
    okBtn.setBounds(10,186,430,40);
    container.add(okBtn);
    
    //{{INIT_CONTROLS
    setTitle("Gcal2Excel");
    getContentPane().setLayout(null);
    setSize(460,280);
    setLocation(80,40);
    setVisible(false);
    setResizable(false);
    
    dateFormat = new SimpleDateFormat(DATE_FORMAT_PATTERN);
    this.loadUserInfo();
    
    //{{REGISTER_LISTENERS
    SymAction lSymAction = new SymAction();
    okBtn.addActionListener(lSymAction);
  }

 /**
   * The ActionListener
   */  
  class SymAction implements ActionListener 
  {
    public void actionPerformed(ActionEvent event)
    {
      Object object = event.getSource();
      
      if ( object == okBtn )
        ok_actionPerformed(event);
    }
  }
  
  public static void main(String args[])
  {
    MetalLookAndFeel.setCurrentTheme(new DefaultMetalTheme());
    JFrame.setDefaultLookAndFeelDecorated(true);
    JDialog.setDefaultLookAndFeelDecorated(true);
    
    JFrame frame = new Gcal2Excel();
    frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
    frame.setVisible(true);
  }
  
  private void ok_actionPerformed(ActionEvent event)
  {
  	String userName 	= nameTxt.getText();
    char[] userPassword = passwordTxt.getPassword();
	String calendarId 	= calendarIdTxt.getText();
	String starts 		= dateFormat.format(startsDateChooser.getDate()); 
	String ends 		= dateFormat.format(endsDateChooser.getDate()); 
	
	Converter converter = new Converter(userName, new String(userPassword), calendarId);
	
	if(converter.convert(starts, ends))
	{
		JOptionPane.showMessageDialog(this,
                                  	"File is created successfully :)",
                                  	"Successfully Converted",
                                  	JOptionPane.PLAIN_MESSAGE);
	}
	else
	{
		JOptionPane.showMessageDialog(this,
                                  	"Error Occured :'(",
                                  	"An Error Occured!",
                                  	JOptionPane.PLAIN_MESSAGE);
	}
	
	this.saveUserInfo(userName, new String(userPassword), calendarId);
	
	Arrays.fill(userPassword, '\0');
  }
  
  private void saveUserInfo(String userName, String userPassword, String calendarId)
  {
  	try 
    {
      FileWriter f;// the actual file stream
      BufferedWriter w;// used to write

      f = new FileWriter( new File("userinf.dat") );
      w = new BufferedWriter(f);
      
      String dataStr = Base64.encode(userName) + "\n" +
      				   Base64.encode(userPassword) + "\n" +
      				   Base64.encode(calendarId);
      				   
      w.write(dataStr);
      
      w.close();
      f.close();
   }
   catch ( Exception e ) 
   {
      JOptionPane.showMessageDialog(this,"Save Error: " + 
                                    e,"Initialization",
                                    JOptionPane.ERROR_MESSAGE);                                     
   }
  }
  
  private void loadUserInfo()
  {
  	try 
    {
      FileReader f;// the actual file stream
      BufferedReader r;// used to read the file line by line

      f = new FileReader( new File("userinf.dat") );
      r = new BufferedReader(f);
      
      this.nameTxt.setText(Base64.decode( r.readLine()));
      this.passwordTxt.setText(Base64.decode( r.readLine()));
      this.calendarIdTxt.setText(Base64.decode( r.readLine()));
      
      r.close();
      f.close();
   }
   catch ( Exception e ) 
   {
      /*JOptionPane.showMessageDialog(this,"Load Error: " + 
                                    e,"Initialization",
                                    JOptionPane.ERROR_MESSAGE);*/                           
   }
   
  }
  
  
}