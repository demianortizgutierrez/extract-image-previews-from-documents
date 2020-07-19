package main;

import java.io.File;
import org.jodconverter.LocalConverter;
import org.jodconverter.filter.text.PageSelectorFilter;
import org.jodconverter.office.LocalOfficeManager;
import org.jodconverter.office.LocalOfficeUtils;
import org.jodconverter.office.OfficeException;


public class DocPreview {

    public static void main(String[] args) throws OfficeException {

     // Create an office manager using the default configuration.
     // The default port is 2002. Note that when an office manager
     // is installed, it will be the one used by default when
     // a converter is created.
     final LocalOfficeManager officeManager = LocalOfficeManager.install(); 
     try {

         // Start an office process and connect to the started instance (on port 2002).
         officeManager.start();

         final File inputFile = new File("file1.docx");
         
         final File outputFile1 = new File("resources/document1.png");
         final File outputFile2 = new File("resources/document2.png");
         final File outputFile3 = new File("resources/document3.png");

         // Create a page selector filter in order to
         // convert only the first page.
         final PageSelectorFilter page1 = new PageSelectorFilter(1);
         final PageSelectorFilter page2 = new PageSelectorFilter(2);
         final PageSelectorFilter page3 = new PageSelectorFilter(3);

         LocalConverter
             .builder()
             .filterChain(page1)
             .build()
             .convert(inputFile)
             .to(outputFile1)
             .execute();
         
         LocalConverter
             .builder()
             .filterChain(page2)
             .build()
             .convert(inputFile)
             .to(outputFile2)
             .execute();         
         
         LocalConverter
             .builder()
             .filterChain(page3)
             .build()
             .convert(inputFile)
             .to(outputFile3)
             .execute();
     } finally {
         // Stop the office process
         LocalOfficeUtils.stopQuietly(officeManager);
     }
    }
}