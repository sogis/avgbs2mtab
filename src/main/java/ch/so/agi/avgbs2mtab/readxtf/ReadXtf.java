package ch.so.agi.avgbs2mtab.readxtf;

import ch.interlis.iom.IomObject;
import ch.interlis.iox.*;
import java.util.HashMap;
import ch.so.agi.avgbs2mtab.mutdat.SetParcel;

/**
 * This Class contains methods to read xtf-files and write specific content to a hashtable
 */
public class ReadXtf {

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
    private SetParcel parceldump;

    public ReadXtf(SetParcel parceldump) {
        this.parceldump = parceldump;
    }

    public void readFile(String xtffilepath) {
        HashMap<String,String> map = readMetadata(xtffilepath);
        readValues(xtffilepath, map);
    }

    private void readValues(String xtffilepath, HashMap<String,String> metadatamap) {
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
                    //Hier beginnt das eigentliche Auslesen!
                    transferParcelAndNewArea(se, metadatamap);
                    ///////////////////////////////////////
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

    private HashMap<String,String> readMetadata(String xtffilepath) {
        HashMap<String, String> map = new HashMap<String, String>();
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
                    //Hier beginnt das Auslesen der Metadaten!
                    map = getMetadata(se);
                    ///////////////////////////////////////
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

    public void transferParcelAndNewArea(StartBasketEvent basket, HashMap<String,String> metadatamap) {

        HashMap objv=new HashMap();

        try {
            //loop threw basket, find things and write them to the Container
            IoxEvent event2;
            while(true){
                event2=ioxReader.read();
                if(event2 instanceof ObjectEvent){
                    IomObject iomObj=((ObjectEvent)event2).getIomObject();
                        objv.put(iomObj.getobjectoid(),iomObj);
                        String aclass=iomObj.getobjecttag();
                        if(aclass.equals(ILI_MUT+".Liegenschaft") || aclass.equals(ILI_GRUDA_MUT+".Liegenschaft")){
                            if(iomObj.getattrvalue("GrundstueckArt").equals("Liegenschaft")) {
                                if (metadatamap.containsKey(iomObj.getobjectoid())) {
                                    int parcelnumber = Integer.parseInt(iomObj.getattrobj("Nummer", 0).getattrvalue("Nummer"));
                                    int area = Integer.parseInt(iomObj.getattrvalue("Flaechenmass"));
                                    parceldump.setParcelNewArea(parcelnumber, area);

                                    try {
                                        int roundingdifference = Integer.parseInt(iomObj.getattrvalue("Korrektur"));
                                        parceldump.setParcelRoundingDifference(parcelnumber, roundingdifference);
                                    } catch (NumberFormatException e) {
                                    }
                                    ;
                                    try {
                                        int oldarea = Integer.parseInt(iomObj.getattrvalue("korrigiertesFlaechenmass"));
                                        parceldump.setParcelOldArea(parcelnumber, oldarea);
                                    } catch (NumberFormatException e) {
                                    }


                                    if (iomObj.getattrvaluecount("Zugang") > 0) {
                                        for (int i = 0; i < iomObj.getattrvaluecount("Zugang"); i++) {
                                            int oldparcelnumber = Integer.parseInt(iomObj.getattrobj("Zugang", i).getattrobj("von", 0).getattrvalue("Nummer"));
                                            int additionarea = Integer.parseInt(iomObj.getattrobj("Zugang", i).getattrvalue("Flaechenmass"));
                                            parceldump.setParcelAddition(parcelnumber, oldparcelnumber, additionarea);
                                        }
                                    }
                                }
                            }
                        }
                }else if(event2 instanceof EndBasketEvent){
                    break;
                }else{
                    throw new IllegalStateException("unexpected event "+event2.getClass().getName());
                }
            }
        }catch(IoxException ex){
            System.out.println("Fehler!");
        }
    }

    private HashMap<String,String> getMetadata(StartBasketEvent basket) throws IoxException {
        // loop threw basket and find all "betroffeneGrundstuecke"
        HashMap objv=new HashMap();
        HashMap<String,String> betroffenegrundstuecke = new HashMap<String,String>();
        IoxEvent event;
        while (true) {
            event = ioxReader.read();
            if (event instanceof ObjectEvent) {
                IomObject iomObj = ((ObjectEvent) event).getIomObject();
                objv.put(iomObj.getobjectoid(), iomObj);
                String aclass = iomObj.getobjecttag();
                if (aclass.equals(ILI_MUT + ".AVMutationBetroffeneGrundstuecke") || aclass.equals(ILI_GRUDA_MUT + ".AVMutationBetroffeneGrundstuecke")) {
                    String ref = iomObj.getattrobj("betroffeneGrundstuecke",0).getobjectrefoid();
                    betroffenegrundstuecke.put(ref,ref);
                    System.out.println("ref = "+ref);
                }
            }
            else if(event instanceof EndBasketEvent){
                break;
            }else{
                throw new IllegalStateException("unexpected event "+event.getClass().getName());
            }
        }
        return betroffenegrundstuecke;
    }
}
