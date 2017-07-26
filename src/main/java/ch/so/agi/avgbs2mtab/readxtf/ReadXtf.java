package ch.so.agi.avgbs2mtab.readxtf;

import ch.interlis.iox.*;

import java.io.File;
import java.util.Map;

/**
 *
 */
public class ReadXtf {

    private Map map;

    /** Ili-Name des avgbs Modells.
     */
    private final String ILI_MODEL="GB2AV";
    private final String ILI_GRUDA_MODEL="GrudaTrans";
    /** Qualifizierter Ili-Name des Topics Mutationstabelle.
     */
    private final String ILI_MUT=ILI_MODEL+".Mutationstabelle";
    private final String ILI_GRUDA_MUT=ILI_GRUDA_MODEL+".Mutationstabelle";

    /** Input-Stream.
     * Wrapper um iomFile.
     */
    private IoxReader ioxReader=null;


    public void readFile(String xtffilepath) {
        try{
            // open xml file
            ioxReader=new ch.interlis.iom_j.xtf.XtfReader(new java.io.File(xtffilepath));
            // loop threw baskets
            IoxEvent event;
            while(true){
                event=ioxReader.read();;
                if(event instanceof ObjectEvent){
                }else if(event instanceof StartBasketEvent){
                    StartBasketEvent se=(StartBasketEvent)event;
                    getParcelAndNewArea(se);
                }else if(event instanceof EndBasketEvent){
                }else if(event instanceof StartTransferEvent){
                    StartTransferEvent se=(StartTransferEvent)event;
                    String sender=se.getSender();
                }else if(event instanceof EndTransferEvent){
                    System.out.flush();
                    ioxReader.close();
                    ioxReader=null;
                    break;
                }else {
                }
            }
        }catch(Exception e){
            System.out.println("FEHLER 1");
        }finally{
            if(ioxReader!=null){
                try{
                    ioxReader.close();
                }catch(IoxException ex){
                    System.out.println(ex.getMessage());
                }
                ioxReader=null;
            }
        }
    }

    public Map getParcelAndNewArea (StartBasketEvent basket) {
        System.out.println(basket.getType()+" (BID "+basket.getBid()+")");

        return map;
    }
}
