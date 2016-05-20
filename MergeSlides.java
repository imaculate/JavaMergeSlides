import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFHyperlink;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
//import org.apache.commons.io.IOUtils;
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


import java.awt.Rectangle;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.FileInputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;
import java.text.DecimalFormat;

import javax.imageio.ImageIO;
import javax.xml.namespace.QName;

import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.apache.poi.sl.usermodel.PictureData.PictureType;
import org.apache.poi.util.IOUtils;

import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.presentationml.x2006.main.CTExtension;
import org.openxmlformats.schemas.drawingml.x2006.main.CTHyperlink;
import org.openxmlformats.schemas.officeDocument.x2006.relationships.STRelationshipId;
import org.openxmlformats.schemas.presentationml.x2006.main.CTApplicationNonVisualDrawingProps;
 import org.openxmlformats.schemas.presentationml.x2006.main.CTPicture;
import org.openxmlformats.schemas.presentationml.x2006.main.CTSlide;
import org.openxmlformats.schemas.presentationml.x2006.main.CTTLCommonMediaNodeData;
import org.openxmlformats.schemas.presentationml.x2006.main.CTTLCommonTimeNodeData;
import org.openxmlformats.schemas.presentationml.x2006.main.CTTimeNodeList;
import org.openxmlformats.schemas.presentationml.x2006.main.STTLTimeIndefinite;
import org.openxmlformats.schemas.presentationml.x2006.main.STTLTimeNodeFillType;
import org.openxmlformats.schemas.presentationml.x2006.main.STTLTimeNodeRestartType;
import org.openxmlformats.schemas.presentationml.x2006.main.STTLTimeNodeType;

import com.xuggle.mediatool.IMediaReader;
import com.xuggle.mediatool.MediaListenerAdapter;
import com.xuggle.mediatool.ToolFactory;
import com.xuggle.mediatool.event.IVideoPictureEvent;
import com.xuggle.xuggler.Global;
import com.xuggle.xuggler.IContainer;
import com.xuggle.xuggler.io.InputOutputStreamHandler;


public class MergeSlides{
     static DecimalFormat df_time = new DecimalFormat("0.####");
    public static void main(String args[]) throws Exception{
    
         if(isOneFile(args)){
            //System.out.println("One file");
            mergeSame(args);
            
         }else{
            //System.out.println("Different files");
            mergeDifferent(args);
         }
         
    }
    
    public static boolean isImage(String in){
      String[] exts = {"jpg", "jpeg", "png", "gif"};
      for(int i= 0; i< exts.length; i++){
         if(!isLink(in) && in.indexOf(exts[i])>=0 )
            return true;
      }  
      
      return false; 
     
      
    }
    
     public static boolean isAudio(String in){
      String[] exts = { "mp3"};
      for(int i= 0; i< exts.length; i++){
         if(!isLink(in) && in.indexOf(exts[i])>=0 )
            return true;
      }  
      
      return false; 
    }
    
    public static boolean isVideo(String in){
      String[] exts = { "mp4", "avi"};
      for(int i= 0; i< exts.length; i++){
         if(!isLink(in) && in.indexOf(exts[i])>=0 )
            return true;
      }  
      
      return false; 
    }
    
    public static boolean isLink(String in){
      return (in.indexOf(':')>= 0);
         
    }
    
    
    
    
    public static void mergeSame(String[] args) throws Exception{
    
        String name = args[1];
        int first = 1;

        while((first< args.length-1) && (name.indexOf(".ppt")<0) || (name.indexOf(".ppt")>=0 && name.indexOf(':')>=0)){
         //do nothing
         first++;
         name = args[first];
         
        }
         String target_dir = args[0];
        String fname = name.substring(0, name.lastIndexOf('-'));
        File file=new File(target_dir, fname);
       
       

        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(file));

        //merging
        List<XSLFSlide> slides = ppt.getSlides();
        
        List<Integer> order = new ArrayList<Integer>();
        List<String> media  = new ArrayList<String>();
        
        for(int i = 1; i< args.length; i++){
            name = args[i]; 
            if( name.indexOf(".ppt")>=0 && name.indexOf(":")<0){//neither image nor link ppt
               order.add(Integer.parseInt(name.substring(1+name.lastIndexOf('-'))));
               //System.out.println(name);
            }else{
               //image or link or audio
               name = name+ "-"+i ;
               media.add(name);
               
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
        //System.out.println(outsize);
        for(int i = outsize; i< sz; i++){ //remove the excess slides in case order.length< sz
            //System.out.println(i);
            ppt.removeSlide(outsize);
        }
        
        for(int i= 0; i< media.size(); i++){
               XSLFSlide imageSlide = ppt.createSlide();
               name = media.get(i);
               int pos = Integer.parseInt(name.substring(1+name.lastIndexOf('-')));
               name = name.substring(0,name.lastIndexOf('-'));
               if(isLink(name)){
                  
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
                        textRun.setText(name);	    
                           //creating the hyperlink
                        XSLFHyperlink link = textRun.createHyperlink();     
                           //setting the link address
                        link.setAddress(name);
                        
                        ppt.setSlideOrder(slide, pos);
               }else if(isImage(name)){
                  
              
                  byte[] pictureData = IOUtils.toByteArray(new FileInputStream(name));

                  XSLFPictureData pd = ppt.addPicture(pictureData, XSLFPictureData.PictureType.PNG);
                  XSLFPictureShape pic = imageSlide.createPicture(pd);
                  ppt.setSlideOrder(imageSlide, pos);
               }else if(isVideo(name)){
                  //System.out.println("Playable");
                  
                  PackagePartName partName = PackagingURIHelper.createPartName("/ppt/media/"+name);
                  PackagePart part = ppt.getPackage().createPart(partName, "video/mpeg");
      
                   OutputStream partOs = part.getOutputStream();

                  FileInputStream fis = new FileInputStream(name);
                  byte buf[] = new byte[1024];
                  for (int readBytes; (readBytes = fis.read(buf)) != -1; partOs.write(buf, 0, readBytes));
                  fis.close();
                  partOs.close();

                  XSLFSlide slide = ppt.createSlide();
                  XSLFPictureShape pv1 = addPreview(ppt, slide, part, 5, 50, 80);
                  addVideo(ppt, slide, part, pv1, 5);
                  addTimingInfo(slide, pv1);
                  ppt.setSlideOrder(slide, pos);
                  
               }else if(isAudio(name)){
                    
                  PackagePartName partName = PackagingURIHelper.createPartName("/ppt/media/"+name);
                  PackagePart part = ppt.getPackage().createPart(partName, "audio/mpeg");
      
                   OutputStream partOs = part.getOutputStream();

                  FileInputStream fis = new FileInputStream(name);
                  byte buf[] = new byte[1024];
                  for (int readBytes; (readBytes = fis.read(buf)) != -1; partOs.write(buf, 0, readBytes));
                  fis.close();
                  partOs.close();

                  XSLFSlide slide = ppt.createSlide();
                  
                  byte[] picture = IOUtils.toByteArray(new FileInputStream("audio.png"));
      
                  //adding the image to the presentation
                  XSLFPictureData idx = ppt.addPicture(picture, XSLFPictureData.PictureType.PNG);
      
            //creating a slide with given picture on it
                  XSLFPictureShape pv1 = slide.createPicture(idx);
                   addAudio(ppt, slide, part, pv1, 5);
                  addTimingInfo(slide, pv1);
                  ppt.setSlideOrder(slide, pos);
               }


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
    
    public static void mergeDifferent(String args[]) throws Exception{
      HashMap<String, File> files = new HashMap<String, File>();
         XMLSlideShow pptOut = new XMLSlideShow();
         File file= null;
         String target_dir = args[0];
         
        
         int n = args.length; //number of slides
         for(int i = 1; i< n; i++){
            //System.out.println(i);
            String st = args[i];
            if(st.indexOf(".ppt")>=0 && st.indexOf(":")<0){
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
           }else{//image, link or media
               if(isImage(st)){
               //System.out.println(st);
                  XSLFSlide imageSlide = pptOut.createSlide();
                  file = new File(st);
               
                  byte[] pictureData = IOUtils.toByteArray(new FileInputStream(file));

                  XSLFPictureData pd = pptOut.addPicture(pictureData, XSLFPictureData.PictureType.PNG);
                  XSLFPictureShape pic = imageSlide.createPicture(pd);
              }else if(isLink(st)){
                  XSLFSlideMaster slideMaster = pptOut.getSlideMasters().get(0);      
                     //select a layout from specified list
                     XSLFSlideLayout slidelayout = slideMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);
                       //XSLFSlideLayout slidelayout = slideMaster.getLayout(SlideLayout.BLANK);    
                        //creating a slide with title and content layout
                     XSLFSlide slide = pptOut.createSlide(slidelayout);    
                        //selection of title place holder
                       XSLFTextShape body = slide.getPlaceholder(1);
                        //XSLFTextShape body = slide.createTextBox();
                          //clear the existing text in the slid
                        body.clearText();      
                           //adding new paragraph
                        XSLFTextRun textRun = body.addNewTextParagraph().addNewTextRun();     
                           //setting the text
                        textRun.setText(st);	    
                           //creating the hyperlink
                        XSLFHyperlink link = textRun.createHyperlink();     
                           //setting the link address
                        link.setAddress(st);

                  
              }else if(isVideo(st)){
              
                     
                  PackagePartName partName = PackagingURIHelper.createPartName("/ppt/media/"+st);
                  PackagePart part = pptOut.getPackage().createPart(partName, "video/mpeg");
      
                   OutputStream partOs = part.getOutputStream();

                  FileInputStream fis = new FileInputStream(st);
                  byte buf[] = new byte[1024];
                  for (int readBytes; (readBytes = fis.read(buf)) != -1; partOs.write(buf, 0, readBytes));
                  fis.close();
                  partOs.close();

                  XSLFSlide slide = pptOut.createSlide();
                  XSLFPictureShape pv1 = addPreview(pptOut, slide, part, 5, 50, 80);
                  addVideo(pptOut, slide, part, pv1, 5);
                  addTimingInfo(slide, pv1);
                  
              }else if(isAudio(st)){
              
               PackagePartName partName = PackagingURIHelper.createPartName("/ppt/media/"+st);
               PackagePart part = pptOut.getPackage().createPart(partName, "audio/mpeg");
               OutputStream partOs = part.getOutputStream();
        //InputStream fis = video.openStream();
               FileInputStream fis = new FileInputStream(st);
                byte buf[] = new byte[1024];
                 for (int readBytes; (readBytes = fis.read(buf)) != -1; partOs.write(buf, 0, readBytes));
               fis.close();
               partOs.close();

               XSLFSlide slide = pptOut.createSlide();
        
               byte[] picture = IOUtils.toByteArray(new FileInputStream("audio.png"));
      
      //adding the image to the presentation
               XSLFPictureData idx = pptOut.addPicture(picture, XSLFPictureData.PictureType.PNG);
      
      //creating a slide with given picture on it
             XSLFPictureShape pv1 = slide.createPicture(idx);
      
               addAudio(pptOut, slide, part, pv1, 5);
               addTimingInfo(slide, pv1);

              }
              
           }
           

            
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
      for(int i = 1; i< pieces.length; i++){//first argument is the directory
         curr = pieces[i];
         
         //System.out.println(curr);
         idx = curr.lastIndexOf('-');
         boolean image = (curr.indexOf(".ppt")<0) || (curr.indexOf(".ppt")>=0 && curr.indexOf(':')>=0) ;//ie not pptx 
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
      if(imageCount == pieces.length-1)
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
    
    
    static XSLFPictureShape addPreview(XMLSlideShow pptx, XSLFSlide slide1, PackagePart videoPart, double seconds, int x, int y) throws IOException {
        // get preview after 5 sec.
        IContainer ic = IContainer.make();
        InputOutputStreamHandler iosh = new InputOutputStreamHandler(videoPart.getInputStream());
        if (ic.open(iosh, IContainer.Type.READ, null) < 0) return null;

        IMediaReader mediaReader = ToolFactory.makeReader(ic);

        // stipulate that we want BufferedImages created in BGR 24bit color space
        mediaReader.setBufferedImageTypeToGenerate(BufferedImage.TYPE_3BYTE_BGR);

        ImageSnapListener isl = new ImageSnapListener(seconds);
        mediaReader.addListener(isl);

        // read out the contents of the media file and
        // dispatch events to the attached listener
        while (!isl.hasFired && mediaReader.readPacket() == null) ;

        mediaReader.close();
        ic.close();

        // add snapshot
        BufferedImage image1 = isl.image;
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        ImageIO.write(image1, "jpeg", bos);
        XSLFPictureData snap = pptx.addPicture(bos.toByteArray(), PictureType.JPEG);
        XSLFPictureShape pic1 = slide1.createPicture(snap);
        pic1.setAnchor(new Rectangle(x, y, image1.getWidth(), image1.getHeight()));
        return pic1;
    }

     
    static void addAudio(XMLSlideShow pptx, XSLFSlide slide1, PackagePart videoPart, XSLFPictureShape pic1, double seconds) throws IOException {

        // add video shape
        PackagePartName partName = videoPart.getPartName();
        PackageRelationship prsEmbed1 = slide1.getPackagePart().addRelationship(partName, TargetMode.INTERNAL, "http://schemas.microsoft.com/office/2007/relationships/media");
        PackageRelationship prsExec1 = slide1.getPackagePart().addRelationship(partName, TargetMode.INTERNAL, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio");
        CTPicture xpic1 = (CTPicture)pic1.getXmlObject();
        CTHyperlink link1 = xpic1.getNvPicPr().getCNvPr().addNewHlinkClick();
        link1.setId("");
        link1.setAction("ppaction://media");

        // add video relation
        CTApplicationNonVisualDrawingProps nvPr = xpic1.getNvPicPr().getNvPr();
        nvPr.addNewVideoFile().setLink(prsExec1.getId());
        CTExtension ext = nvPr.addNewExtLst().addNewExt();
        // see http://msdn.microsoft.com/en-us/library/dd950140(v=office.12).aspx
        ext.setUri("{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}");
        String p14Ns = "http://schemas.microsoft.com/office/powerpoint/2010/main";
        XmlCursor cur = ext.newCursor();
        cur.toEndToken();
        cur.beginElement(new QName(p14Ns, "media", "p14"));
        cur.insertNamespace("p14", p14Ns);
        cur.insertAttributeWithValue(new QName(STRelationshipId.type.getName().getNamespaceURI(), "embed"), prsEmbed1.getId());
        cur.beginElement(new QName(p14Ns, "trim", "p14"));
        cur.insertAttributeWithValue("st", df_time.format(seconds*1000.0));
        cur.dispose();

    }
    static void addVideo(XMLSlideShow pptx, XSLFSlide slide1, PackagePart videoPart, XSLFPictureShape pic1, double seconds) throws IOException {

        // add video shape
        PackagePartName partName = videoPart.getPartName();
        PackageRelationship prsEmbed1 = slide1.getPackagePart().addRelationship(partName, TargetMode.INTERNAL, "http://schemas.microsoft.com/office/2007/relationships/media");
        PackageRelationship prsExec1 = slide1.getPackagePart().addRelationship(partName, TargetMode.INTERNAL, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video");
        CTPicture xpic1 = (CTPicture)pic1.getXmlObject();
        CTHyperlink link1 = xpic1.getNvPicPr().getCNvPr().addNewHlinkClick();
        link1.setId("");
        link1.setAction("ppaction://media");

        // add video relation
        CTApplicationNonVisualDrawingProps nvPr = xpic1.getNvPicPr().getNvPr();
        nvPr.addNewVideoFile().setLink(prsExec1.getId());
        CTExtension ext = nvPr.addNewExtLst().addNewExt();
        // see http://msdn.microsoft.com/en-us/library/dd950140(v=office.12).aspx
        ext.setUri("{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}");
        String p14Ns = "http://schemas.microsoft.com/office/powerpoint/2010/main";
        XmlCursor cur = ext.newCursor();
        cur.toEndToken();
        cur.beginElement(new QName(p14Ns, "media", "p14"));
        cur.insertNamespace("p14", p14Ns);
        cur.insertAttributeWithValue(new QName(STRelationshipId.type.getName().getNamespaceURI(), "embed"), prsEmbed1.getId());
        cur.beginElement(new QName(p14Ns, "trim", "p14"));
        cur.insertAttributeWithValue("st", df_time.format(seconds*1000.0));
        cur.dispose();

    }

    static void addTimingInfo(XSLFSlide slide1, XSLFPictureShape pic1) {
        // add slide timing information, so video can be controlled
        CTSlide xslide = slide1.getXmlObject();
        CTTimeNodeList ctnl;
        if (!xslide.isSetTiming()) {
            CTTLCommonTimeNodeData ctn = xslide.addNewTiming().addNewTnLst().addNewPar().addNewCTn();
            ctn.setDur(STTLTimeIndefinite.INDEFINITE);
            ctn.setRestart(STTLTimeNodeRestartType.NEVER);
            ctn.setNodeType(STTLTimeNodeType.TM_ROOT);
            ctnl = ctn.addNewChildTnLst();
        } else {
            ctnl = xslide.getTiming().getTnLst().getParArray(0).getCTn().getChildTnLst();
        }

        CTTLCommonMediaNodeData cmedia = ctnl.addNewVideo().addNewCMediaNode();
        cmedia.setVol(80000);
        CTTLCommonTimeNodeData ctn = cmedia.addNewCTn();
        ctn.setFill(STTLTimeNodeFillType.HOLD);
        ctn.setDisplay(false);
        ctn.addNewStCondLst().addNewCond().setDelay(STTLTimeIndefinite.INDEFINITE);
        cmedia.addNewTgtEl().addNewSpTgt().setSpid(""+pic1.getShapeId());
    }


    static class ImageSnapListener extends MediaListenerAdapter {
        final double SECONDS_BETWEEN_FRAMES;
        final long MICRO_SECONDS_BETWEEN_FRAMES;
        boolean hasFired = false;
        BufferedImage image = null;

        // The video stream index, used to ensure we display frames from one and
        // only one video stream from the media container.
        int mVideoStreamIndex = -1;

        // Time of last frame write
        long mLastPtsWrite = Global.NO_PTS;

        public ImageSnapListener(double seconds) {
            SECONDS_BETWEEN_FRAMES = seconds;
            MICRO_SECONDS_BETWEEN_FRAMES =
                    (long)(Global.DEFAULT_PTS_PER_SECOND * SECONDS_BETWEEN_FRAMES);
        }


        @Override
        public void onVideoPicture(IVideoPictureEvent event) {

            if (event.getStreamIndex() != mVideoStreamIndex) {
                // if the selected video stream id is not yet set, go ahead an
                // select this lucky video stream
                if (mVideoStreamIndex != -1) return;
                mVideoStreamIndex = event.getStreamIndex();
            }

            long evtTS = event.getTimeStamp();

            // if uninitialized, back date mLastPtsWrite to get the very first frame
            if (mLastPtsWrite == Global.NO_PTS)
                mLastPtsWrite = Math.max(0, evtTS - MICRO_SECONDS_BETWEEN_FRAMES);

            // if it's time to write the next frame
            if (evtTS - mLastPtsWrite >= MICRO_SECONDS_BETWEEN_FRAMES) {
                if (!hasFired) {
                    image = event.getImage();
                    hasFired = true;
                }
                // update last write time
                mLastPtsWrite += MICRO_SECONDS_BETWEEN_FRAMES;
            }
        }
    }


   
}
