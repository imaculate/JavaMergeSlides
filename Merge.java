import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

public class Merge
{
	public static void main(String[] args) throws Exception
	{
		String file = "example.pptx";
		System.out.println("Starting merge of file: " +file);
		XMLSlideShow ppt = new XMLSlideShow();
		FileInputStream is = new FileInputStream(file);
		XMLSlideShow src = new XMLSlideShow(is);
		is.close();
      int i = 0;

		for (XSLFSlide srcSlide : src.getSlides())
		{
          System.out.println(i);
			ppt.createSlide().importContent(srcSlide);
         i++;
        
		}

		File f = new File("merged.pptx");
		FileOutputStream out = new FileOutputStream(f.getAbsolutePath());
		ppt.write(out);
		System.out.println("Finished merge to file: " +f.getAbsolutePath());
	}
}