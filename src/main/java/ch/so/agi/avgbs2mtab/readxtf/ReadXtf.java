package ch.so.agi.avgbs2mtab.readxtf;


import ch.interlis.iom.IomObject;
import ch.interlis.iox.*;

import java.util.HashMap;
import java.util.Map;

import ch.so.agi.avgbs2mtab.mutdat.GetAndSetParcel;

/**
 * This Class contains methods to read xtf-files and write specific content to a hashtable
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

    public Map readFile(String xtffilepath) {
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
                    map = getParcelAndNewArea(se);
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
        return map;
    }

    public Map getParcelAndNewArea (StartBasketEvent basket) {
        HashMap objv=new HashMap();
        int parcelnumber;
        int area;
        int oldparcelnumber;
        int oldarea;
        GetAndSetParcel getandsetter = new GetAndSetParcel();

        try{
            // loop threw basket
            IoxEvent event;
            while(true){
                event=ioxReader.read();
                if(event instanceof ObjectEvent){
                    IomObject iomObj=((ObjectEvent)event).getIomObject();
                        objv.put(iomObj.getobjectoid(),iomObj);
                        String aclass=iomObj.getobjecttag();
                        if(aclass.equals(ILI_MUT+".Liegenschaft") || aclass.equals(ILI_GRUDA_MUT+".Liegenschaft")){
                            area = Integer.parseInt(iomObj.getattrvalue("Flaechenmass"));
                            parcelnumber = Integer.parseInt(iomObj.getattrobj("Nummer",0).getattrvalue("Nummer"));
                            //System.out.println("ParcelNumber = "+parcelnumber);
                            //System.out.println("ParcelArea = "+area);
                            getandsetter.setParcelNewArea(parcelnumber,area);
                            if(iomObj.getattrvaluecount("Zugang")>0) {
                                //System.out.println("Zugänge: "+iomObj.getattrvaluecount("Zugang"));
                                for(int i = 0; i < iomObj.getattrvaluecount("Zugang"); i++) {
                                    oldparcelnumber = Integer.parseInt(iomObj.getattrobj("Zugang",i).getattrobj("von",0).getattrvalue("Nummer"));
                                    oldarea = Integer.parseInt(iomObj.getattrobj("Zugang",i).getattrvalue("Flaechenmass"));
                                    //System.out.println("Zugang = "+oldparcelnumber);
                                    //System.out.println("Fläche = "+oldarea);
                                    getandsetter.setParcelAddition(parcelnumber,oldparcelnumber,oldarea);
                                }
                            }
                        }
                }else if(event instanceof EndBasketEvent){
                    break;
                }else{
                    throw new IllegalStateException("unexpected event "+event.getClass().getName());
                }
            }
        }catch(IoxException ex){
            System.out.println("Fehler!");
        }
        System.out.println(map.toString());
        return map;
    }
}
