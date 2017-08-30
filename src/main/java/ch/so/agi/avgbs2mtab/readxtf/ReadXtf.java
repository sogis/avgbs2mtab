package ch.so.agi.avgbs2mtab.readxtf;

import ch.interlis.iom.IomObject;
import ch.interlis.iox.*;

import java.io.IOException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;

import ch.so.agi.avgbs2mtab.mutdat.DataExtractionParcel;
import ch.so.agi.avgbs2mtab.mutdat.SetDPR;
import ch.so.agi.avgbs2mtab.mutdat.SetParcel;
import ch.so.agi.avgbs2mtab.util.Avgbs2MtabException;


//todo Der XtfReader wird sowieso viermal aufgerufen. Bitte mit mir anschauen ob eine Aufteilung von ReadXtf in ReadDPR und ReadParcel den code lesbarer machen würde
//todo Aufgrund der Eventstruktur des Readers von Claude haben wir eine hohe Codeverdoppelung - Bitte die Lösungsmuster dafür bei mir abholen
/**
 * This Class contains methods to read xtf-files and write specific content to a hashtable
 */
public class ReadXtf {

    private static final String ILI_MODELNAME ="GB2AV";
    private final String ILI_MUT= ILI_MODELNAME +".Mutationstabelle";
    private IoxReader ioxReader=null;
    private SetParcel parceldump;
    private SetDPR drpdump;

    private static final Logger LOGGER = Logger.getLogger(ReadXtf.class.getName());

    public ReadXtf(SetParcel parceldump, SetDPR drpdump) {

        this.parceldump = parceldump;
        this.drpdump = drpdump;
    }

    public void readFile(String xtffilepath) throws IOException {
        LOGGER.log(Level.CONFIG,"Start reading the file");

        //The parcelmetadatamap contains a set of Ref-keys from parcels affected by the mutation.
        HashSet<String> parcelmetadatamap = readParcelMetadata(xtffilepath);
        //The dprmetadatamap contains the ref-key as Key and a Map of ref-keys and areas from parcels, affected by the dpr.
        HashMap<String,HashMap> dprmetadatamap = readDRPMetadata(xtffilepath);

        readValues(xtffilepath, parcelmetadatamap, dprmetadatamap);
    }


    /////////////////////////////////////
    //PARCEL META-DATA //////////////////
    /////////////////////////////////////
    private HashSet<String> readParcelMetadata(String xtffilepath) {

        HashSet<String> parcelmetadatamap = new HashSet<>();
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

                    assertModelIsAvGbs(se);
                    //Hier beginnt das Auslesen der Metadaten!
                    parcelmetadatamap = getParcelMetadata(se);
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
            if(e instanceof Avgbs2MtabException) {
                throw (Avgbs2MtabException) e;
            }
            else {
                throw new Avgbs2MtabException("error in reading xtf file", e);
            }
        }
        finally{
            if(ioxReader!=null){
                try{
                    ioxReader.close();
                }catch(IoxException ex){
                    LOGGER.log(Level.WARNING,"Got a IoxException: "+ex);
                }
                ioxReader=null;
            }
        }
        return parcelmetadatamap;
    }

    private HashSet<String> getParcelMetadata(StartBasketEvent basket) throws IoxException {
        // loop threw basket and find all "betroffeneGrundstuecke"

        HashSet<String> oidbetroffenegrundstuecke = new HashSet<>();
        IoxEvent event;
        while (true) {
            event = ioxReader.read();
            if (event instanceof ObjectEvent) {
                IomObject iomObj = ((ObjectEvent) event).getIomObject();
                String aclass = iomObj.getobjecttag();

                if (aclass.equals(ILI_MUT + ".AVMutationBetroffeneGrundstuecke")) {
                    String ref = iomObj.getattrobj("betroffeneGrundstuecke",0).getobjectrefoid();

                    oidbetroffenegrundstuecke.add(ref);
                }
            }
            else if(event instanceof EndBasketEvent){
                break;
            }else{
                throw new IllegalStateException("unexpected event "+event.getClass().getName());
            }
        }
        return oidbetroffenegrundstuecke;
    }

    /////////////////////////////////////
    //DPR META-DATA ////////////////////
    /////////////////////////////////////
    private HashMap<String, HashMap> readDRPMetadata(String xtffilepath) {
        HashMap<String, HashMap> map = new HashMap<>();
        try{
            ioxReader=new ch.interlis.iom_j.xtf.XtfReader(new java.io.File(xtffilepath));
            // loop threw baskets
            IoxEvent event;
            while(true){
                event=ioxReader.read();;
                if(event instanceof ObjectEvent){
                }else if(event instanceof StartBasketEvent){
                    StartBasketEvent se=(StartBasketEvent)event;

                    assertModelIsAvGbs(se);
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
            if(e instanceof Avgbs2MtabException){
                throw (Avgbs2MtabException)e;
            }
            else{
                throw new Avgbs2MtabException("Error reading xtf file ...", e);
            }
        }
        finally{
            if(ioxReader!=null){
                try{
                    ioxReader.close();
                }catch(IoxException ex){
                    LOGGER.log(Level.WARNING,"Got a IoxException: "+ex);
                }
                ioxReader=null;
            }
        }
        return map;
    }

    private HashMap<String, HashMap> getDRPMetadata(StartBasketEvent basket) throws IoxException {
        // loop threw basket and find all "betroffeneGrundstuecke"

        HashMap<String,HashMap> dpranteilanliegenschaft = new HashMap<String, HashMap>();

        HashMap<String, Integer> liegtaufmap = new HashMap<String, Integer>();
        IoxEvent event;
        while (true) {
            event = ioxReader.read();
            if (event instanceof ObjectEvent) {
                IomObject iomObj = ((ObjectEvent) event).getIomObject();
                String aclass = iomObj.getobjecttag();

                if (aclass.equals(ILI_MODELNAME + ".Grundstuecksbeschrieb.Anteil")) {
                    String drpnumber = iomObj.getattrobj("flaeche",0).getobjectrefoid();
                    String liegt_auf = iomObj.getattrobj("liegt_auf",0).getobjectrefoid();
                    Integer area = Integer.parseInt(iomObj.getattrvalue("Flaechenmass"));

                    if (dpranteilanliegenschaft.get(drpnumber) != null) {
                        liegtaufmap = dpranteilanliegenschaft.get(drpnumber);
                        liegtaufmap.put(liegt_auf,area);
                    } else {
                        liegtaufmap.put(liegt_auf, area);
                    }
                    dpranteilanliegenschaft.put(drpnumber,liegtaufmap);

                }
            }
            else if(event instanceof EndBasketEvent){
                break;
            }else{
                throw new IllegalStateException("unexpected event "+event.getClass().getName());
            }
        }
        return dpranteilanliegenschaft;
    }

    /////////////////////////////////////
    //Get Data /////////////////////////
    /////////////////////////////////////

    /**
     * Main Method. Loop over Baskets and start the transfer of Parcel- and DPR-Values.
     * @param xtffilepath
     * @param parcelmetadatamap
     * @param drpmetadatamap
     */
    private void readValues(String xtffilepath, HashSet<String> parcelmetadatamap, HashMap<String, HashMap> drpmetadatamap) {
        try{
            //Get Parcel-Information
            // transferDPR(StartBasketEvent se, HashMap<String, HashMap> drpmetadatamap)
            // transferParcel(StartBasketEvent basket, HashMap<String,String> metadatamap) {
            ioxReader=new ch.interlis.iom_j.xtf.XtfReader(new java.io.File(xtffilepath));

            IoxEvent event = ioxReader.read();
            while(event != null){
                event=ioxReader.read();;
                if(event instanceof ObjectEvent){
                }else if(event instanceof StartBasketEvent){
                    StartBasketEvent se=(StartBasketEvent)event;

                    assertModelIsAvGbs(se);

                    //Main Parcel-Value-Function
                    transferParcel(se, parcelmetadatamap);
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
                event = ioxReader.read();
            }
            //Get DPR-Information
            ioxReader=new ch.interlis.iom_j.xtf.XtfReader(new java.io.File(xtffilepath));
            IoxEvent event3 = ioxReader.read();
            while(event3 != null){
                event3=ioxReader.read();
                if(event3 instanceof ObjectEvent){
                }else if(event3 instanceof StartBasketEvent){
                    StartBasketEvent se=(StartBasketEvent)event3;

                    assertModelIsAvGbs(se);
                    //Main DPR-Value-Function
                    transferDPR(se, drpmetadatamap);
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
                event3 = ioxReader.read();
            }
        }catch(Exception e){
            LOGGER.log(Level.WARNING,"Error reading Values");
            throw new RuntimeException(e);
        }
        finally{
            if(ioxReader!=null){
                try{
                    ioxReader.close();
                }catch(IoxException ex){
                    LOGGER.log(Level.WARNING,"Got a IoxException: "+ex);
                }
                ioxReader=null;
            }
        }
    }

    /////////////////////////////////////
    //Get ParcelData ////////////////////
    /////////////////////////////////////

    private void transferParcel(StartBasketEvent basket, HashSet<String> metadataset) {
        try {
            //loop threw basket, find things and write them to the Container
            IoxEvent event2;
            while(true){
                event2=ioxReader.read();
                if(event2 instanceof ObjectEvent){
                    IomObject iomObj=((ObjectEvent)event2).getIomObject();
                    String aclass=iomObj.getobjecttag();
                    if(aclass.equals(ILI_MUT+".Liegenschaft")){
                        getParcelLiegenschaft(metadataset, iomObj);
                    }
                }else if(event2 instanceof EndBasketEvent){
                    break;
                }else{
                    throw new IllegalStateException("unexpected event "+event2.getClass().getName());
                }
            }
        }catch(IoxException ex){
            LOGGER.log(Level.WARNING,"Got a Error in transferParcel: "+ex);
        }
    }

    private void getParcelLiegenschaft(HashSet<String> metadataset, IomObject iomObj) {
        if(iomObj.getattrvalue("GrundstueckArt").equals("Liegenschaft")) {
            if (metadataset.contains(iomObj.getobjectoid())) {
                int parcelnumber = Integer.parseInt(iomObj.getattrobj("Nummer", 0).getattrvalue("Nummer"));
                int area = Integer.parseInt(iomObj.getattrvalue("Flaechenmass"));
                parceldump.setParcelNewArea(parcelnumber, area);
                parceldump.setParcelOldArea(parcelnumber,area); //vorläufig wird die alte Fläche = neue Fläche gesetzt.

                try {
                    int roundingdifference = Integer.parseInt(iomObj.getattrvalue("Korrektur"));
                    parceldump.setParcelRoundingDifference(parcelnumber, roundingdifference);
                } catch (NumberFormatException e) {
                    //throw new Avgbs2MtabException(Avgbs2MtabException.TYPE_NUMBERFORMAT, "NumberFormatException");
                };                                    ;

                getZugaenge(iomObj, parcelnumber, area);
            }
        }
    }

    private void getZugaenge(IomObject iomObj, int parcelnumber, int area) {
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

    /////////////////////////////////////
    //Get DPRData    ////////////////////
    /////////////////////////////////////

    private void transferDPR(StartBasketEvent se, HashMap<String, HashMap> drpmetadatamap) {

        try {
            //loop threw basket, find things and write them to the Container
            IoxEvent event;
            while (true) {
                event = ioxReader.read();
                if (event instanceof ObjectEvent) {
                    IomObject iomObj = ((ObjectEvent) event).getIomObject();
                    String aclass = iomObj.getobjecttag();
                    getDPR(drpmetadatamap, iomObj, aclass);
                    getDPRLiegenschaft(iomObj, aclass);
                    getDeletedDPR(iomObj, aclass);
                }else if(event instanceof EndBasketEvent){
                    break;
                }else{
                    throw new IllegalStateException("unexpected event "+event.getClass().getName());
                }
            }
        }catch(IoxException ex){
            LOGGER.log(Level.WARNING,"Got an Error in transferDPR: "+ex);
        }
    }

    private void getDeletedDPR(IomObject iomObj, String aclass) {
        if (aclass.equals(ILI_MUT + ".AVMutation")) {
            Integer numberofdeletedparcels = iomObj.getattrvaluecount("geloeschteGrundstuecke");
            for(Integer i=0;i<numberofdeletedparcels;++i) {
                Integer nummer = Integer.parseInt(iomObj.getattrobj("geloeschteGrundstuecke", i).getattrvalue("Nummer"));
                DataExtractionParcel parceldatagetter = (DataExtractionParcel)parceldump;
                if(!parceldatagetter.getOldParcelNumbers().contains(nummer)&&!parceldatagetter.getNewParcelNumbers().contains(nummer)) {
                    drpdump.setDPRNewArea(nummer, 0);
                }
            }
        }
    }

    private void getDPRLiegenschaft(IomObject iomObj, String aclass) {
        if (aclass.equals(ILI_MUT + ".Liegenschaft")) {
            Integer numbercount = iomObj.getattrvaluecount("Nummer");
            for (int i = 0;i<numbercount;++i) {
                int parcelnumber = Integer.parseInt(iomObj.getattrobj("Nummer", i).getattrvalue("Nummer"));
                String parcelref = iomObj.getobjectoid();
                drpdump.setDPRNumberAndRef(parcelref, parcelnumber);
            }
        }
    }

    private void getDPR(HashMap<String, HashMap> drpmetadatamap, IomObject iomObj, String aclass) {
        if (aclass.equals(ILI_MUT + ".Flaeche")) {
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
    }

    /////////////////////////////////////
    //Utility        ////////////////////
    /////////////////////////////////////

    private static void assertModelIsAvGbs(StartBasketEvent se){

        String namev[] = se.getType().split("\\.");
        String modelName = namev[0];

        if(!modelName.equals(ILI_MODELNAME)) {
            throw new Avgbs2MtabException(
                    Avgbs2MtabException.TYPE_TRANSFERDATA_NOT_FOR_AVGBS_MODEL,
                    String.format("Given transferfile references wrong model %s (should reference %s)", modelName, ILI_MODELNAME)
            );
        }
    }
}
