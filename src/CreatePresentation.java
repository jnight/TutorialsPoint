import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

public class CreatePresentation {
   
   public static void main(String args[]) throws IOException{
   
      //creating a new empty slide show
     // XMLSlideShow ppt = new XMLSlideShow();	     
      
      //creating an FileOutputStream object
      File file =new File("D:\\My\\example1.pptx");
     // FileOutputStream out = new FileOutputStream(file);
      FileInputStream inputstream =new FileInputStream(file);
      XMLSlideShow ppt = new XMLSlideShow(inputstream);
        
      //adding slides to the slodeshow
      XSLFSlide slide1 = ppt.createSlide();
      XSLFSlide slide2 = ppt.createSlide();
      
      FileOutputStream out = new FileOutputStream(file);
      
      //saving the changes to a file
      ppt.write(out);
      System.out.println("Presentation created successfully");
      out.close();
   }
}