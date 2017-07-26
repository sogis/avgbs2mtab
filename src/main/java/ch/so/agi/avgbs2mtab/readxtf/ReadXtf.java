package ch.so.agi.avgbs2mtab.readxtf;


import ch.interlis.iom.IomObject;
import ch.interlis.iox.*;

import java.io.File;
import java.util.HashMap;
import java.util.Iterator;
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
            throw new RuntimeException(e);
        }
        finally{
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
        HashMap objv=new HashMap();
        String parcelnumber;
        String area;
        String oldparcelnumber;
        String oldarea;

        try{

            // loop threw basket
            IoxEvent event;
            while(true){
                event=ioxReader.read();
                if(event instanceof ObjectEvent){
                    IomObject iomObj=((ObjectEvent)event).getIomObject();
                        objv.put(iomObj.getobjectoid(),iomObj);
                        String aclass=iomObj.getobjecttag();
                        //EhiLogger.debug("object "+aclass);
                        if(aclass.equals(ILI_MUT+".Liegenschaft") || aclass.equals(ILI_GRUDA_MUT+".Liegenschaft")){
                            area = iomObj.getattrvalue("Flaechenmass");
                            parcelnumber = iomObj.getattrobj("Nummer",0).getattrvalue("Nummer");
                            System.out.println("ParcelNumber = "+parcelnumber);
                            System.out.println("ParcelArea = "+area);
                            try {
                                System.out.println("Objekte: "+iomObj.getattrvaluecount("Zugang"));
                                for(int i = 0; i < iomObj.getattrvaluecount("Zugang"); i++) {
                                    oldparcelnumber = iomObj.getattrobj("Zugang",i).getattrobj("von",0).getattrvalue("Nummer");
                                    oldarea = iomObj.getattrobj("Zugang",i).getattrvalue("Flaechenmass");

                                    System.out.println("Zugang = "+oldparcelnumber);
                                    System.out.println("Fläche = "+oldarea);
                                }
                            } catch (NullPointerException e) {
                            }
                        }
                        if(aclass.equals(ILI_MUT+".AVMutationsAnnulation") || aclass.equals(ILI_GRUDA_MUT+".AVMutationsAnnulation")){
                            //mutannulationv.add(iomObj);
                        }
                        if(aclass.equals(ILI_MUT+".AVMutationBetroffeneGrundstuecke") || aclass.equals(ILI_GRUDA_MUT+".AVMutationBetroffeneGrundstuecke")){
                            //EhiLogger.debug(iomObj.toString());
                            //mutBetrfGs.add(iomObj);
                        }
                        //if(aclass.equals(ILI_GSBESCHR+".Anteil")){
                            //anteilv.add(iomObj);
                        //}
                }else if(event instanceof EndBasketEvent){
                    break;
                }else{
                    throw new IllegalStateException("unexpected event "+event.getClass().getName());
                }
            }
        }catch(IoxException ex){
            System.out.println("Fehler!");
        }
        return map;
    }
}
