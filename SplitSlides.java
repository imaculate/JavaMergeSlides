import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.regex.Pattern;

/**
 * Created with IntelliJ IDEA.
 * User: DELL
 * Date: 4/17/16
 * Time: 5:13 PM
 * To change this template use File | Settings | File Templates.
 */
public class SplitSlides {
    public static void main(String args[]) throws IOException {

        //Opening an existing slide
        String target_path = args[0];
        File file=new File(target_path);
        String target_dir = file.getParent();
        String fname = file.getName();
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(file));

        //splitting
        List<XSLFSlide> slides = ppt.getSlides();
        int sz =  slides.size();
         Dimension pgsize = ppt.getPageSize();
        XSLFSlide slide;
        int num = 0;
        while (num< sz){

            XMLSlideShow ppte = new XMLSlideShow(new FileInputStream(file));
            List<XSLFSlide> slidese = ppte.getSlides();
            XSLFSlide selectesdslide = slidese.get(num);

            //bringing it to the top
            ppte.setSlideOrder(selectesdslide, 0);
            for(int i = 1; i< sz; i++){
                    ppte.removeSlide(1);

            }

        OPCPackage pkg = ppte.getPackage();
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
            num++;
            //writing to image
            BufferedImage img = new BufferedImage(pgsize.width, pgsize.height,BufferedImage.TYPE_INT_RGB);
            Graphics2D graphics = img.createGraphics();
            graphics.setPaint(Color.WHITE);
            graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));
            slide = slidese.get(0);
            slide.draw(graphics);
            FileOutputStream out = new FileOutputStream( new File(target_dir, fname+ "-"+num+".png"));
            javax.imageio.ImageIO.write(img, "png", out);
            out.close();



            //writing to pptx
            /*File fi=new File("example-"+num+".pptx");
            FileOutputStream out = new FileOutputStream(fi);
            //Saving the changes to the presentation
            ppte.write(out);
            out.close();  */

        }
        System.out.println("Done removing packages");
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
