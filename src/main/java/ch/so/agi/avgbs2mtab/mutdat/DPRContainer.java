package ch.so.agi.avgbs2mtab.mutdat;

import java.util.*;

/**
 *
 */


public class DPRContainer implements SetDPR, MetadataOfDPRMutation, DataExtractionDPR {
    Map<Integer,Map> map=new Hashtable<Integer,Map>(); //Haupt Map
    HashMap<String, Integer> laysonrefandareamap = new HashMap<>(); //innere Map der Haupt-Map
    HashMap<String, Integer> numberandrefmap = new HashMap<>(); //Map mit ref (String) und Parzellennummer (int)
    HashMap<String, String> affectedparcelsmap = new HashMap<>(); //Eine kleine Map mit allen Parzellen die irgendwie betroffen sind.
    HashMap<Integer, Integer> newareamap = new HashMap<>(); //HashMap mit den gesamtflächen der DPRs.

    public void setDPRWithAdditions(Integer dprnumber, String laysonref, Integer area) {
        if(map.get(dprnumber) != null) {
            Map laysonrefandarea = map.get(dprnumber);
            laysonrefandarea.put(laysonref,area);
        } else {
            laysonrefandareamap.put(laysonref,area);
        }
        map.put(dprnumber, laysonrefandareamap);
        affectedparcelsmap.put(laysonref,laysonref);
    }

    @Override
    public void setDPRNumberAndRef(String ref, int parcelnumber) {
        numberandrefmap.put(ref,parcelnumber);
    }

    @Override
    public void setDPRNewArea(Integer parcelnumber, Integer area) {
        newareamap.put(parcelnumber,area);
    }

    ////////////////////////////////////////////////////////
    // GET- Methoden ///////////////////////////////////////
    ///////////////////////////////////////////////////////

    @Override
    public int getNumberOfDPRs() {
        int numberofdprs = map.size();
        //Addiere noch die gelöschten dazu....
        int i = 0;
        for (Integer key : newareamap.keySet()) {
            if(newareamap.get(key).equals(0)) {
                i++;
            }
        }
        numberofdprs += i;
        return numberofdprs;
    }

    @Override
    public int getNumberOfParcelsAffectedByDPRs() {
        int numberofparcelsaffected = affectedparcelsmap.size();
        return numberofparcelsaffected;
    }

    @Override
    public List<Integer> getParcelsAffectedByDPRs() {
        HashMap<Integer,Integer> affectedparcelsnumbermap = new HashMap<>();
        for(String key : affectedparcelsmap.keySet()) {
            int keyparcelnumber = numberandrefmap.get(key);
            affectedparcelsnumbermap.put(keyparcelnumber,keyparcelnumber);
        }
        Map<Integer,Integer> sortedmap = new TreeMap<>(affectedparcelsnumbermap); //sortiert die Map der grösse der keys nach
        List<Integer> parcelsaffectedmydprs = new ArrayList<>(sortedmap.keySet());
        return parcelsaffectedmydprs;
    }

    @Override
    public List<Integer> getNewDPRs() {
        Map<Integer,Map> sortedmap = new TreeMap<>(map);
        List<Integer> newdprs = new ArrayList<>(sortedmap.keySet());
        for (Integer key : newareamap.keySet()) {
            if(newareamap.get(key).equals(0)) {
                newdprs.add(key);
            }
        }
        return newdprs;
    }

    @Override
    public int getAddedAreaDPR(int parcelNumberAffectedByDPR, int dpr) {
        Map<String,Integer> innerdprmap = map.get(dpr);
        String ref = getKeyFromValue(numberandrefmap,parcelNumberAffectedByDPR);
        int addedarea = innerdprmap.get(ref);
        return addedarea;
    }

    @Override
    public int getNewAreaDPR(int dpr) {
        int newarea = newareamap.get(dpr);
        return newarea;
    }

    @Override
    public int getRoundingDifferenceDPR(int dpr) {
        int sumaddedareas = 0;
        Map<String, Integer> internalmap = map.get(dpr);
        for(String key : internalmap.keySet()) {
            int area = internalmap.get(key);
            sumaddedareas += area;
        }
        int roundingdiference = getNewAreaDPR(dpr)-sumaddedareas;
        return roundingdiference;
    }

    public static String getKeyFromValue(Map<String, Integer> hm, Integer value) {
        for (String key : hm.keySet()) {
            if (hm.get(key).equals(value)) {
                return key;
            }
        }
        return null;
    }
}
