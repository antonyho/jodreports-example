package net.antonyho.jodmapper;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.util.HashMap;
import java.util.Map;

import javax.imageio.ImageIO;

import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;

import junit.framework.Test;
import junit.framework.TestCase;
import junit.framework.TestSuite;
import net.sf.jooreports.templates.DocumentTemplate;
import net.sf.jooreports.templates.DocumentTemplateException;
import net.sf.jooreports.templates.DocumentTemplateFactory;
import net.sf.jooreports.templates.image.ImageSource;
import net.sf.jooreports.templates.image.RenderedImageSource;

/**
 * Unit test for simple App.
 */
public class AppTest 
    extends TestCase
{
    /**
     * Create the test case
     *
     * @param testName name of the test case
     */
    public AppTest( String testName )
    {
        super( testName );
    }

    /**
     * @return the suite of tests being tested
     */
    public static Test suite()
    {
        return new TestSuite( AppTest.class );
    }

    /**
     * Rigourous Test :-)
     * @throws IOException 
     * @throws DocumentTemplateException 
     * @throws InterruptedException 
     */
    public void testApp() throws IOException, DocumentTemplateException, InterruptedException
    {
    	DocumentTemplateFactory docTemplateFactory = new DocumentTemplateFactory();
    	final URL resource = this.getClass().getResource("/template1.odt");
    	System.out.println(resource);
    	DocumentTemplate template = docTemplateFactory.getTemplate(resource.openStream());
    	System.out.println(resource.openStream());
//    	File odtTemplate = new File("C:\\projects\\jodreports-test\\template1.odt");
//    	DocumentTemplate template = docTemplateFactory.getTemplate(odtTemplate);
    	Map<String, Object> data = new HashMap<String, Object>();	// Referenced by variableName.fieldName using freemarker
    	data.put("sendername", "Antony Ho");
    	data.put("companyname", "Antony Workshop");
    	data.put("addrstreet", "11-13 Tai Yuen St");
    	data.put("addrpostalcode", "00000");
    	data.put("addrcity", "Kwai Chung");
    	data.put("addrstate", "HK");
    	
    	data.put("recipienttitle", "Lord");
    	data.put("recipientname", "Page Larry");
    	data.put("recipientstreet", "1600 Amphitheatre Parkway");
    	data.put("recipientpostalcode", "94043");
    	data.put("recipientcity", "Mountain View");
    	data.put("recipientstate", "CA");
    	
    	
    	/*
    	 * Add image mapping in template
    	 * 1. Add an image into proper position in the template
    	 * 2. Double click the image into the image property
    	 * 3. Go to Option tab
    	 * 4. The Name value should be: jooscript.image(image_map_key)
    	 */
//    	ImageSource flagImg = new RenderedImageSource(ImageIO.read(new File("C:\\projects\\jodreports-test\\flag2.png")));
    	ImageSource flagImg = new RenderedImageSource(ImageIO.read(this.getClass().getResource("/flag2.png")));
    	data.put("flag", flagImg);
    	
    	File odtOutput = new File("C:\\projects\\jodreports-test\\testoutput.odt");
    	template.createDocument(data, new FileOutputStream(odtOutput));
    	
    	// convert ODF to PDF
    	File pdfOutput = new File("C:\\projects\\jodreports-test\\testoutput.pdf");
//    	File officeDir = new File("C:/Program Files (x86)/LibreOffice 5/program/");
    	
    	ProcessBuilder procBuilder = new ProcessBuilder("C:/Program Files (x86)/LibreOffice 5/program/soffice.exe", "-headless", "-accept=\"socket,host=127.0.0.1,port=8100;urp;\"", "-nofirststartwizard");
//    	procBuilder.directory(officeDir);
    	Process proc = procBuilder.start();
//    	proc.waitFor();
    	OpenOfficeConnection connection = new SocketOpenOfficeConnection("127.0.0.1", 8100);
    	connection.connect();
    	
    	DocumentConverter converter = new OpenOfficeDocumentConverter(connection);
    	converter.convert(odtOutput, pdfOutput);
    	
    	proc.destroy();
    	
        assertTrue( true );
    }
}
