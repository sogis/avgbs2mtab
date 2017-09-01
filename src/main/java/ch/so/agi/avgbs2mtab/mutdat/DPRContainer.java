package ch.so.agi.avgbs2mtab.mutdat;

import java.util.*;

/**
 *
 */


public class DPRContainer implements SetDPR, MetadataOfDPRMutation, DataExtractionDPR {
    Map<Integer,Map> map=new Hashtable<Integer,Map>(); //Haupt Map //todo namen und doku und camelcase
    HashMap<String, Integer> laysonrefandareamap = new HashMap<>(); //innere Map der Haupt-Map //todo namen und doku und camelcase
    HashMap<String, Integer> numberandrefmap = new HashMap<>(); //Map mit ref (String) und Parzellennummer (int) //todo namen und doku und camelcase
    HashMap<String, String> affectedparcelsmap = new HashMap<>(); //Eine kleine Map mit allen Parzellen die irgendwie betroffen sind. //todo guter name. camelcase beachten
    HashMap<Integer, Integer> newareamap = new HashMap<>(); //HashMap mit den gesamtflächen der DPRs. //todo namen und doku und camelcase

    ////////////////////////////////////////////////////////
    // SET- Methoden ///////////////////////////////////////
    ///////////////////////////////////////////////////////

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
    public Integer getNumberOfDPRs() {
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
    public Integer getNumberOfParcelsAffectedByDPRs() {
        int numberofparcelsaffected = affectedparcelsmap.size();
        return numberofparcelsaffected;
    }

    @Override
    public List<Integer> getOrderedListOfParcelsAffectedByDPRs() {
        List<Integer> parcelsaffectedmydprs = new ArrayList<>();
        for(String key : affectedparcelsmap.keySet()) {
            int keyparcelnumber = numberandrefmap.get(key);
            if(!parcelsaffectedmydprs.contains(keyparcelnumber)) {
                parcelsaffectedmydprs.add(keyparcelnumber);
            }
        }
        Collections.sort(parcelsaffectedmydprs);
        return parcelsaffectedmydprs;
    }

    @Override
    public List<Integer> getOrderedListOfNewDPRs() {
        List<Integer> newdprs = new ArrayList<>(map.keySet());
        for (Integer key : newareamap.keySet()) {
            if(newareamap.get(key).equals(0)) {
                newdprs.add(key);
            }
        }
        Collections.sort(newdprs);
        return newdprs;
    }

    @Override
    public Integer getAddedAreaDPR(int parcelNumberAffectedByDPR, int dpr) {
        Map<String,Integer> innerdprmap = map.get(dpr);
        String ref = getKeyFromValue(numberandrefmap,parcelNumberAffectedByDPR);
        int addedarea = innerdprmap.get(ref);
        return addedarea;
    }

    @Override
    public Integer getNewAreaDPR(int dpr) {
        int newarea = newareamap.get(dpr);
        return newarea;
    }

    @Override
    public Integer getRoundingDifferenceDPR(int dpr) {
        Integer sumaddedareas = 0;
        Integer roundingdifference = null;
        Map<String, Integer> internalmap = map.get(dpr);
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
    public static String getKeyFromValue(Map<String, Integer> hm, Integer value) {
        for (String key : hm.keySet()) {
            if (hm.get(key).equals(value)) {
                return key;
            }
        }
        return null;
    }
}
