import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.commons.io.IOUtils;
import java.sql.Timestamp;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Date;
import java.util.regex.Pattern;
import java.util.HashMap;
import java.util.ArrayList;

public class MergeSlides{
    public static void main(String args[]) throws IOException{
    
         if(isOneFile(args)){
            //System.out.println("One file");
            mergeSame(args);
            
         }else{
            //System.out.println("Different files");
            mergeDifferent(args);
         }
         
    }
    public static void mergeSame(String[] args) throws IOException{
    
        String name = args[0];
        int first = 0;
        while((first< args.length-1) && name.substring(name.length()-4).equals(".png")){
         //do nothing
         first++;
         name = args[first];
         
        }
        
        String target_path = name.substring(0, name.lastIndexOf('-'));
        File file=new File(target_path);
        String target_dir = file.getParent();
        String fname = file.getName();

        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(file));

        //merging
        List<XSLFSlide> slides = ppt.getSlides();
        
        List<Integer> order = new ArrayList<Integer>();
        List<String> images  = new ArrayList<String>();
        
        for(int i = 0; i< args.length; i++){
            name = args[i]; 
            if(!name.substring(name.length()-4).equals(".png")){
               order.add(Integer.parseInt(name.substring(1+name.lastIndexOf('-'))));
               //System.out.println(name);
            }else{
               //image
               name = name+ "-"+i ;
               images.add(name);
               
            }
        }
        int sz = slides.size();
        int[] idc = new int[sz] ;    //new indices after reshuffling
        for(int i = 0; i<sz; i++){
            idc[i]=i;
        }
        for(int i = order.size()-1; i>= 0; i--){
           // System.out.println(order[i]);
            
            int idx =  order.get(i);
            XSLFSlide selectesdslide = slides.get(idc[idx]);
            ppt.setSlideOrder(selectesdslide, 0);
            for(int j = 0; j<idx; j++){
                idc[j]+=1;          //shift items to the right
            }
        }
        int outsize = order.size();
        for(int i = outsize; i< sz; i++){ //remove the excess slides in case order.length< sz
            //System.out.println(i);
            ppt.removeSlide(outsize);
        }
        
        for(int i= 0; i< images.size(); i++){
               XSLFSlide imageSlide = ppt.createSlide();
               name = images.get(i);
               int pos = Integer.parseInt(name.substring(1+name.lastIndexOf('-')));
               name = name.substring(0,name.lastIndexOf('-'));
               
               byte[] pictureData = IOUtils.toByteArray(new FileInputStream(name));

               XSLFPictureData pd = ppt.addPicture(pictureData, XSLFPictureData.PictureType.PNG);
               XSLFPictureShape pic = imageSlide.createPicture(pd);
               ppt.setSlideOrder(imageSlide, pos);


        }
        
        
        OPCPackage pkg = ppt.getPackage();
        for(PackagePart mediaPart :
                pkg.getPartsByName(Pattern.compile("/ppt/media/.*?"))){
            try {
                if(!isReferenced(mediaPart, pkg)) {
                    //System.out.println(mediaPart.getPartName() + " is not referenced. removing.... ");
                    pkg.removePart(mediaPart);
                }
            } catch (Exception e) {
                e.printStackTrace();  //To change body of catch statement use File | Settings | File Templates.
            }
        }
            for(PackagePart embPart :
                    pkg.getPartsByName(Pattern.compile("/ppt/embeddings/.*?"))){
                try {
                    if(!isReferenced(embPart, pkg)) {
                        //System.out.println(embPart.getPartName() + " is not referenced. removing.... ");
                        pkg.removePart(embPart);
                    }
                } catch (Exception e) {
                    e.printStackTrace();  //To change body of catch statement use File | Settings | File Templates.
                }
            }
        
        //write to output
        Date date= new Date();
        String ts = new Timestamp(date.getTime()).toString().replace(' ', '-'); 
        ts = ts.substring(0, ts.lastIndexOf(':')).replace(':', '_');
      
        
        fname = fname.substring(0, fname.lastIndexOf("."));
        fname = fname+"-"+ts+".pptx";
        File fout = new File(target_dir, fname);
        

        FileOutputStream out = new FileOutputStream(fout);
        
        ppt.write(out);
        out.close();


    }
    
    public static void mergeDifferent(String args[]) throws IOException{
      HashMap<String, File> files = new HashMap<String, File>();
         XMLSlideShow pptOut = new XMLSlideShow();
         File file= null;
         String target_dir = "";
         
        
         int n = args.length; //number of slides
         for(int i = 0; i< n; i++){
            //System.out.println(i);
            String st = args[i];
            if(!st.substring(st.length()-4).equals(".png")){
            int idx = st.lastIndexOf('-');
            String finame = st.substring(0,idx) ; //gonna extract , finame contains directory_path too
            //System.out.println(st);
            int slideNo = Integer.parseInt(st.substring(idx+1));
             
            if(!files.containsKey(finame)){
               file = new File(finame);  
               files.put(finame, file);             
            }else{
               file = files.get(finame);
             
            }
                         
            
            XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(file));
            pptOut.createSlide().importContent(ppt.getSlides().get(slideNo));
           }else{
               //System.out.println(st);
               XSLFSlide imageSlide = pptOut.createSlide();
               file = new File(st);
               
               byte[] pictureData = IOUtils.toByteArray(new FileInputStream(file));

               XSLFPictureData pd = pptOut.addPicture(pictureData, XSLFPictureData.PictureType.PNG);
               XSLFPictureShape pic = imageSlide.createPicture(pd);
              
           }
           if(i == 0)//same target directory for all the files
               target_dir = file.getAbsoluteFile().getParent();   

            
         }
         

        //write to output
        Date date= new Date();
        String ts = new Timestamp(date.getTime()).toString().replace(' ', '-'); 
        ts = ts.substring(0, ts.lastIndexOf(':')).replace(':', '_');
      
        
        
        String fname = "merged-"+ts+".pptx";
        //System.out.println("Writing to "+ fname);
        //System.out.println(target_dir);
        File fout = new File(target_dir, fname);
        

        FileOutputStream out = new FileOutputStream(fout);
        
        pptOut.write(out);
        out.close();
   
    }
    public static boolean isOneFile(String pieces[]){
      String prev = pieces[0];
      String curr = "";
      int idxp = prev.lastIndexOf('-');
      int idx = 0;
      int imageCount = 0;
      for(int i = 0; i< pieces.length; i++){
         curr = pieces[i];
         
         //System.out.println(curr);
         idx = curr.lastIndexOf('-');
         boolean image = curr.substring(curr.length()-4).equals(".png");
         if(image){
            //System.out.println("Image");
            imageCount++;
            continue;
          }
         if(idxp>= 0 && !curr.substring(0, idx).equals(curr.substring(0, idxp))){
            return false;
         }
         prev = curr;
         idxp = prev.lastIndexOf('-');
         
      }
      if(imageCount == pieces.length)
         return false;
      return true;
      
    }
    
    
    public static boolean isReferenced(PackagePart mediaPart,
                                       OPCPackage pkg) throws Exception {
        for(PackagePart part : pkg.getParts()){
            if(part.isRelationshipPart()) continue;

            for(PackageRelationship rel : part.getRelationships()){
                if(
                        mediaPart.getPartName().getURI().equals(rel.getTargetURI())){
                    System.out.println("mediaPart[" + mediaPart.getPartName() + "] is referenced by " + part.getPartName());
                    return true;
                }
            }
        }
        return false;
    }

   
}
