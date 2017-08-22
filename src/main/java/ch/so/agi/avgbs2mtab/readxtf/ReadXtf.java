package ch.so.agi.avgbs2mtab.readxtf;

import ch.interlis.iom.IomObject;
import ch.interlis.iox.*;

import java.util.HashMap;
import java.util.Map;

import ch.so.agi.avgbs2mtab.mutdat.SetDPR;
import ch.so.agi.avgbs2mtab.mutdat.SetParcel;

/**
 * This Class contains methods to read xtf-files and write specific content to a hashtable
 */
public class ReadXtf {

    private final String ILI_MODEL="GB2AV";
    private final String ILI_GRUDA_MODEL="GrudaTrans";
    private final String ILI_MUT=ILI_MODEL+".Mutationstabelle";
    private final String ILI_GRUDA_MUT=ILI_GRUDA_MODEL+".Mutationstabelle";

    private IoxReader ioxReader=null;
    private SetParcel parceldump;
    private SetDPR drpdump;

    public ReadXtf(SetParcel parceldump, SetDPR drpdump) {

        this.parceldump = parceldump;
        this.drpdump = drpdump;
    }

    public void readFile(String xtffilepath) {
        HashMap<String,String> parcelmetadatamap = readParcelMetadata(xtffilepath);
        HashMap<String,HashMap> drpmetadatamap = readDRPMetadata(xtffilepath);
        readValues(xtffilepath, parcelmetadatamap, drpmetadatamap);
    }

    /**
     * Main Method. Loop over Baskets and start the transfer of Parcel- and DPR-Values.
     * @param xtffilepath
     * @param parcelmetadatamap
     * @param drpmetadatamap
     */
    private void readValues(String xtffilepath, HashMap<String,String> parcelmetadatamap, HashMap<String, HashMap> drpmetadatamap) {
        try{
            //Get Parcel-Information
            ioxReader=new ch.interlis.iom_j.xtf.XtfReader(new java.io.File(xtffilepath));
            IoxEvent event;
            while(true){
                event=ioxReader.read();;
                if(event instanceof ObjectEvent){
                }else if(event instanceof StartBasketEvent){
                    StartBasketEvent se=(StartBasketEvent)event;
                    //Main Parcel-Value-Function
                    transferParcelAndNewArea(se, parcelmetadatamap);
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
            //Get DPR-Information
            ioxReader=new ch.interlis.iom_j.xtf.XtfReader(new java.io.File(xtffilepath));
            IoxEvent event3;
            while(true){
                event3=ioxReader.read();
                if(event3 instanceof ObjectEvent){
                }else if(event3 instanceof StartBasketEvent){
                    StartBasketEvent se=(StartBasketEvent)event3;
                    //Main DPR-Value-Function
                    transferDRP(se, drpmetadatamap);
                    ///////////////////////////////////////
                }else if(event3 instanceof EndBasketEvent){
                }else if(event3 instanceof StartTransferEvent){
                    StartTransferEvent se=(StartTransferEvent)event3;
                    String sender=se.getSender();
                }else if(event3 instanceof EndTransferEvent){
                    System.out.flush();
                    ioxReader.close();
                    ioxReader=null;
                    break;
                }else {
                }
            }
        }catch(Exception e){
            System.out.println("Error reading Values");
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

    private void transferDRP(StartBasketEvent se, HashMap<String, HashMap> drpmetadatamap) {
        HashMap objv=new HashMap();

        try {
            //loop threw basket, find things and write them to the Container
            IoxEvent event;
            while (true) {
                event = ioxReader.read();
                if (event instanceof ObjectEvent) {
                    IomObject iomObj = ((ObjectEvent) event).getIomObject();
                    objv.put(iomObj.getobjectoid(), iomObj);
                    String aclass = iomObj.getobjecttag();
                    if (aclass.equals(ILI_MUT + ".Flaeche") || aclass.equals(ILI_GRUDA_MUT + ".Flaeche")) {
                        if (iomObj.getattrvalue("GrundstueckArt").startsWith("SelbstRecht")) { //Selbstrecht.* ?
                            if (drpmetadatamap.containsKey(iomObj.getobjectoid())) {
                                int parcelnumber = Integer.parseInt(iomObj.getattrobj("Nummer", 0).getattrvalue("Nummer"));
                                int newarea = Integer.parseInt(iomObj.getattrvalue("Flaechenmass"));
                                String parcelref = iomObj.getobjectoid();
                                Map internalmap = drpmetadatamap.get(iomObj.getobjectoid());
                                for (Object key : internalmap.keySet()) {
                                    String fromparcelref = key.toString();
                                    Integer area = Integer.parseInt(internalmap.get(key).toString());
                                    drpdump.setDPRWithAdditions(parcelnumber,fromparcelref,area);
                                }
                                drpdump.setDPRNumberAndRef(parcelref,parcelnumber);
                                drpdump.setDPRNewArea(parcelnumber,newarea);
                            }
                        }
                    }
                    if (aclass.equals(ILI_MUT + ".Liegenschaft") || aclass.equals(ILI_GRUDA_MUT + ".Liegenschaft")) {
                        int parcelnumber = Integer.parseInt(iomObj.getattrobj("Nummer", 0).getattrvalue("Nummer"));
                        String parcelref = iomObj.getobjectoid();
                        drpdump.setDPRNumberAndRef(parcelref,parcelnumber);
                    }
                    if (aclass.equals(ILI_MUT + ".AVMutation") || aclass.equals(ILI_GRUDA_MUT + ".AVMutation")) {
                        if (iomObj.getattrobj("geloeschteGrundstuecke",0) != null) {
                            Integer nummer = Integer.parseInt(iomObj.getattrobj("geloeschteGrundstuecke",0).getattrvalue("Nummer"));
                            drpdump.setDPRNewArea(nummer,0);
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
    }

    private HashMap<String,String> readParcelMetadata(String xtffilepath) {
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
                    map = getParcelMetadata(se);
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
            System.out.println("FEHLER 2");
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

    private HashMap<String, HashMap> readDRPMetadata(String xtffilepath) {
        HashMap<String, HashMap> map = new HashMap<>();
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
                    map = getDRPMetadata(se);
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
            System.out.println("FEHLER 3");
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
                                    parceldump.setParcelOldArea(parcelnumber,area); //vorläufig wird die alte Fläche = neue Fläche gesetzt.

                                    try {
                                        int roundingdifference = Integer.parseInt(iomObj.getattrvalue("Korrektur"));
                                        parceldump.setParcelRoundingDifference(parcelnumber, roundingdifference);
                                    } catch (NumberFormatException e) {
                                    };                                    ;

                                    if (iomObj.getattrvaluecount("Zugang") > 0) {
                                        int additionsum = 0;
                                        for (int i = 0; i < iomObj.getattrvaluecount("Zugang"); i++) {
                                            int oldparcelnumber = Integer.parseInt(iomObj.getattrobj("Zugang", i).getattrobj("von", 0).getattrvalue("Nummer"));
                                            int additionarea = Integer.parseInt(iomObj.getattrobj("Zugang", i).getattrvalue("Flaechenmass"));
                                            parceldump.setParcelAddition(parcelnumber, oldparcelnumber, additionarea);
                                            additionsum += additionarea;
                                        }
                                        if(area != additionsum){
                                            parceldump.setParcelOldArea(parcelnumber,area-additionsum);
                                        }
                                        else{
                                            parceldump.delParcelOldArea(parcelnumber);
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



    private HashMap<String,String> getParcelMetadata(StartBasketEvent basket) throws IoxException {
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

    private HashMap<String, HashMap> getDRPMetadata(StartBasketEvent basket) throws IoxException {
        // loop threw basket and find all "betroffeneGrundstuecke"
        HashMap objv=new HashMap();

        HashMap<String,HashMap> anteil = new HashMap<String, HashMap>();
        HashMap<String, Integer> liegt_auf_map = new HashMap<String, Integer>();
        IoxEvent event;
        while (true) {
            event = ioxReader.read();
            if (event instanceof ObjectEvent) {
                IomObject iomObj = ((ObjectEvent) event).getIomObject();
                objv.put(iomObj.getobjectoid(), iomObj);
                String aclass = iomObj.getobjecttag();
                if (aclass.equals(ILI_MODEL + ".Grundstuecksbeschrieb.Anteil") || aclass.equals(ILI_MODEL + ".Grundstuecksbeschrieb.Anteil")) {
                    String drpnumber = iomObj.getattrobj("flaeche",0).getobjectrefoid();
                    String liegt_auf = iomObj.getattrobj("liegt_auf",0).getobjectrefoid();
                    Integer area = Integer.parseInt(iomObj.getattrvalue("Flaechenmass"));
                    Map test = anteil.get(drpnumber);
                    if (test != null) {
                        liegt_auf_map = anteil.get(drpnumber);
                        liegt_auf_map.put(liegt_auf,area);
                    } else {
                        try {
                            liegt_auf_map.put(liegt_auf, area);
                        } catch (Exception h) {
                            System.out.println("Scheiss Fehler");
                        }
                    }
                    anteil.put(drpnumber,liegt_auf_map);

                    anteil.put(drpnumber, liegt_auf_map);
                }
            }
            else if(event instanceof EndBasketEvent){
                break;
            }else{
                throw new IllegalStateException("unexpected event "+event.getClass().getName());
            }
        }
        return anteil;
    }
}
