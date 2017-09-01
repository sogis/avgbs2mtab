package ch.so.agi.avgbs2mtab.mutdat;

import java.util.*;

/**
 *
 */
public class ParcelContainer implements SetParcel, MetadataOfParcelMutation, DataExtractionParcel {

    Map<Integer,Map> map=new Hashtable<Integer,Map>(); //Haupt Map
    Map<Integer,Integer> parcelnewareamap = new Hashtable<>(); //Map mit den neuen Flächen
    Map<Integer,Integer> parcelrestareamap = new Hashtable<>(); //Map mit den Rest Flächen (Diagonale
    Map<Integer,Integer> parcelroundingdifferencemap = new Hashtable<>(); //Map mit den Rundungsdifferenzen

    ///////////////////////////////////////////////
    // SET-Methoden //////////////////////////////
    //////////////////////////////////////////////

    @Override
    public void setParcelAddition(int newparcelnumber, int oldparcelnumber, int area) {
        //Versuch die Parcelmap aus der main-map zu holen und füge den neuen Wert hinzu. Hat es noch keine Parcelmap von dieser Parzelle, dann leg eine neue an.
        //Schlussendlich füge die Parcemap (neu oder alt) wieder zur main-map hinzu.

        Map<Integer,Integer> parcelmap = map.get(newparcelnumber);
        if(parcelmap == null){
            parcelmap = new Hashtable<Integer, Integer>();
            map.put(newparcelnumber, parcelmap);
        }
        parcelmap.put(oldparcelnumber, area);
    }

    @Override
    public void setParcelNewArea(int newparcelnumber, int newarea) {
            parcelnewareamap.put(newparcelnumber,newarea);
    }

    @Override
    public void setParcelOldArea(int oldparcelnumber, int oldarea) {
            parcelrestareamap.put(oldparcelnumber,oldarea);
    }

    @Override
    public void delParcelOldArea(int oldparcelnumber) {
        parcelrestareamap.remove(oldparcelnumber);
    }

    @Override
    public void setParcelRoundingDifference(int parcel, int roundingdifference) {
            parcelroundingdifferencemap.put(parcel,roundingdifference); //todo camel case!!
    }

    ////////////////////////////////////
    //GET-Methoden  ///////////////////
    ///////////////////////////////////

    @Override
    public List<Integer> getOrderedListOfOldParcelNumbers() {
        List<Integer> oldparcelnumbers = new ArrayList<>();
        //Add all parcelnumbers in the inner-map from the main map to the oldparcelmap
        for(Integer key : map.keySet()) {
            Map internalmap = map.get(key);
            for(Object keyoldparcels : internalmap.keySet()) {
                if(!oldparcelnumbers.contains(keyoldparcels)) {
                    oldparcelnumbers.add((Integer) keyoldparcels);
                }
            }
        }
        //Add also all parcelnumbers from parcelrestareamap to the oldparcelmap
        for(Integer key : parcelrestareamap.keySet()) {
            if(!oldparcelnumbers.contains(key)) {
                oldparcelnumbers.add(key);
            }
        }
        Collections.sort(oldparcelnumbers);
        return oldparcelnumbers;
    }

    @Override
    public List<Integer> getOrderedListOfNewParcelNumbers() {
        List<Integer> newparcelnumbers = new ArrayList<>(parcelnewareamap.keySet());
        Collections.sort(newparcelnumbers);
        return newparcelnumbers;
    }
    @Override
    public Integer getAddedArea(int newparcel, int oldparcel) {
        Map addmap = map.get(newparcel);
        Integer areaadded = null; //Todo: Etwas unglücklich, aber bisher unvermeidlich!
        if (addmap!=null) {
            areaadded = (Integer) addmap.get(oldparcel);
        }
        return areaadded;
    }

    @Override
    public Integer getNewArea(int newParcelNumber) {
        Integer newarea = parcelnewareamap.get(newParcelNumber);
        return newarea;
    }

    @Override
    public Integer getRoundingDifference(int oldParcelNumber) {
        Integer roundingdifference = parcelroundingdifferencemap.get(oldParcelNumber);
        return roundingdifference;
    }

    @Override
    public Integer getNumberOfOldParcels() {
        Integer numberofoldparcels = getOrderedListOfOldParcelNumbers().size();
        return numberofoldparcels;
    }

    @Override
    public Integer getNumberOfNewParcels() {
        Integer numberofnewparcels = getOrderedListOfNewParcelNumbers().size();
        return numberofnewparcels;
    }

    @Override
    public Integer getRestAreaOfParcel(int oldParcelNumber) {
        Integer restarea = parcelrestareamap.get(oldParcelNumber);
        return restarea;
    }


}
