package ch.so.agi.avgbs2mtab.mutdat;

import java.util.*;

/**
 *
 */


public class DPRContainer implements SetDPR, MetadataOfDPRMutation, DataExtractionDPR {
    Map<Integer,Map> mainDprMap =new Hashtable<Integer,Map>(); //Main Map. Contains the DPR-Number as key and the laysOnRefAndAreaMap.
    HashMap<String, Integer> laysOnRefAndAreaMap = new HashMap<>(); //Map, containing the Ref of the Parcel and the Area concerning the DPR.
    HashMap<String, Integer> numberAndRefMap = new HashMap<>(); //Map contains the Ref-String and the Parcel- or DPR-Number.
    HashMap<String, String> affectedParcelsMap = new HashMap<>(); //A small mainParcelMap, containing the numbers of affected Parcels.
    HashMap<Integer, Integer> newAreaMap = new HashMap<>(); //Map, contains the area of the DPRs.

    ////////////////////////////////////////////////////////
    // SET- Methoden ///////////////////////////////////////
    ///////////////////////////////////////////////////////

    public void setDPRWithAdditions(Integer dprnumber, String laysonref, Integer area) {
        if(mainDprMap.get(dprnumber) != null) {
            Map laysonrefandarea = mainDprMap.get(dprnumber);
            laysonrefandarea.put(laysonref,area);
        } else {
            laysOnRefAndAreaMap.put(laysonref,area);
        }
        mainDprMap.put(dprnumber, laysOnRefAndAreaMap);
        affectedParcelsMap.put(laysonref,laysonref);
    }

    @Override
    public void setDPRNumberAndRef(String ref, int parcelnumber) {
        numberAndRefMap.put(ref,parcelnumber);
    }

    @Override
    public void setDPRNewArea(Integer parcelnumber, Integer area) {
        newAreaMap.put(parcelnumber,area);
    }

    ////////////////////////////////////////////////////////
    // GET- Methoden ///////////////////////////////////////
    ///////////////////////////////////////////////////////

    @Override
    public Integer getNumberOfDPRs() {
        int numberofdprs = mainDprMap.size();
        //Addiere noch die gelöschten dazu....
        int i = 0;
        for (Integer key : newAreaMap.keySet()) {
            if(newAreaMap.get(key).equals(0)) {
                i++;
            }
        }
        numberofdprs += i;
        return numberofdprs;
    }

    @Override
    public Integer getNumberOfParcelsAffectedByDPRs() {
        int numberofparcelsaffected = affectedParcelsMap.size();
        return numberofparcelsaffected;
    }

    @Override
    public List<Integer> getOrderedListOfParcelsAffectedByDPRs() {
        List<Integer> parcelsaffectedmydprs = new ArrayList<>();
        for(String key : affectedParcelsMap.keySet()) {
            int keyparcelnumber = numberAndRefMap.get(key);
            if(!parcelsaffectedmydprs.contains(keyparcelnumber)) {
                parcelsaffectedmydprs.add(keyparcelnumber);
            }
        }
        Collections.sort(parcelsaffectedmydprs);
        return parcelsaffectedmydprs;
    }

    @Override
    public List<Integer> getOrderedListOfNewDPRs() {
        List<Integer> newdprs = new ArrayList<>(mainDprMap.keySet());
        for (Integer key : newAreaMap.keySet()) {
            if(newAreaMap.get(key).equals(0)) {
                newdprs.add(key);
            }
        }
        Collections.sort(newdprs);
        return newdprs;
    }

    @Override
    public Integer getAddedAreaDPR(int parcelNumberAffectedByDPR, int dpr) {
        Map<String,Integer> innerdprmap = mainDprMap.get(dpr);
        String ref = getKeyFromValue(numberAndRefMap,parcelNumberAffectedByDPR);
        int addedarea = innerdprmap.get(ref);
        return addedarea;
    }

    @Override
    public Integer getNewAreaDPR(int dpr) {
        int newarea = newAreaMap.get(dpr);
        return newarea;
    }

    @Override
    public Integer getRoundingDifferenceDPR(int dpr) {
        Integer sumaddedareas = 0;
        Integer roundingdifference = null;
        Map<String, Integer> internalmap = mainDprMap.get(dpr);
        if (internalmap != null) {
            for (String key : internalmap.keySet()) {
                Integer area = internalmap.get(key);
                sumaddedareas += area;
            }
            roundingdifference = getNewAreaDPR(dpr) - sumaddedareas;
        }

        return roundingdifference;
    }

    //todo wieso nicht einfach hm.containsKey()? Was ist der Zweck der Funktion?
    //containsKey liefert einen Boolean zurück (Ja, der Key existiert). Was ich aber brauche, ist der Key als String.
    //Ich habe einen Wert (Nummer) und brauche den Key (String).
    public static String getKeyFromValue(Map<String, Integer> hm, Integer value) {
        for (String key : hm.keySet()) {
            if (hm.get(key).equals(value)) {
                return key;
            }
        }
        return null;
    }
}
