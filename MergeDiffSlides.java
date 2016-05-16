import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import java.sql.Timestamp;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Date;
import java.util.regex.Pattern;
import java.util.HashMap;

public class MergeDiffSlides {
    public static void main(String args[]) throws IOException{
         HashMap<String, File> files = new HashMap<String, File>();
         XMLSlideShow pptOut = new XMLSlideShow();
         File file= null;
         String target_dir = "";
         
        
         int n = args.length; //number of slides
         for(int i = 0; i< n; i++){
            System.out.println(i);
            String st = args[i];
            int idx = st.lastIndexOf('-');
            String finame = st.substring(0,idx) ; //gonna extract , finame contains directory_path too
            int slideNo = Integer.parseInt(st.substring(idx+1));
             
            if(!files.containsKey(finame)){
               file = new File(finame);  
               files.put(finame, file);             
            }else{
               file = files.get(finame);
             
            }
             if(i == 1)//same target directory for all the files
               target_dir = file.getParent();   
            
            
            XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(file));
            pptOut.createSlide().importContent(ppt.getSlides().get(slideNo));
            
         }
         

        //write to output
        Date date= new Date();
        String ts = new Timestamp(date.getTime()).toString().replace(' ', '-'); 
        ts = ts.substring(0, ts.lastIndexOf(':')).replace(':', '_');
      
        
        
        String fname = "merged-"+ts+".pptx";
        File fout = new File(target_dir, fname);
        

        FileOutputStream out = new FileOutputStream(fout);
        
        pptOut.write(out);
        out.close();

    }
    
   
}
