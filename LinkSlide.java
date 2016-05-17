import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFHyperlink;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

//import java.io.IOException;
public class LinkSlide{
   public static void main(String[] args) throws IOException{
   
    //create an empty presentation
      XMLSlideShow ppt = new XMLSlideShow();
      
      //getting the slide master object
      XSLFSlideMaster slideMaster = ppt.getSlideMasters().get(0);
      
      //select a layout from specified list
      XSLFSlideLayout slidelayout = slideMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);
      //XSLFSlideLayout slidelayout = slideMaster.getLayout(SlideLayout.BLANK);
     
      //creating a slide with title and content layout
      XSLFSlide slide = ppt.createSlide(slidelayout);
      
      //selection of title place holder
     XSLFTextShape body = slide.getPlaceholder(1);
      //XSLFTextShape body = slide.createTextBox();

      //clear the existing text in the slid
      body.clearText();
      
      //adding new paragraph
      XSLFTextRun textRun = body.addNewTextParagraph().addNewTextRun();
      
      //setting the text
      textRun.setText("Tutorials point");	
      
      //creating the hyperlink
      XSLFHyperlink link = textRun.createHyperlink();
      
      //setting the link address
      link.setAddress("http://www.tutorialspoint.com/");
      
      //create the file object            
      File file=new File("hyperlink.pptx");
      FileOutputStream out = new FileOutputStream(file);
      
      //save the changes in a file
      ppt.write(out);
      System.out.println("slide cretated successfully");
      out.close();              
                
   }
}