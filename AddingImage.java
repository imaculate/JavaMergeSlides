import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;

public class AddingImage {
   
   public static void main(String args[]) throws IOException{
   
      //creating a presentation 
      XMLSlideShow ppt = new XMLSlideShow();
      
      //creating a slide in it 
      XSLFSlide slide = ppt.createSlide();
      
      //reading an image
      File image=new File("Full.png");
      
      //converting it into a byte array
      byte[] picture = IOUtils.toByteArray(new FileInputStream(image));
      
      //adding the image to the presentation
      int idx = ppt.addPicture(picture, XSLFPictureData.PictureType.PNG);
      
      //creating a slide with given picture on it
      XSLFPictureShape pic = slide.createPicture(idx);
      
      //creating a file object 
      File file=new File("addingimage.pptx");
      FileOutputStream out = new FileOutputStream(file);
      
      //saving the changes to a file
      ppt.write(out);
      System.out.println("image added successfully");
      out.close();	
   }
}