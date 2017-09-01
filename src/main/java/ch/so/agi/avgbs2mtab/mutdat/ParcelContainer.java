package ch.so.agi.avgbs2mtab.mutdat;

import java.util.*;

/**
 *
 */
public class ParcelContainer implements SetParcel, MetadataOfParcelMutation, DataExtractionParcel {

    Map<Integer,Map> mainParcelMap =new Hashtable<Integer,Map>(); //The main parcel-map.
    Map<Integer,Integer> parcelNewAreaMap = new Hashtable<>(); //Map contains the new area of a parcel.
    Map<Integer,Integer> parcelRemainingAreaMap = new Hashtable<>(); //Map contains the remaining-area of a Parcel (Diagonale).
    Map<Integer,Integer> parcelRoundingDifferenceMap = new Hashtable<>(); //Map contains the roundingdifference.

    ///////////////////////////////////////////////
    // SET-Methoden //////////////////////////////
    //////////////////////////////////////////////

    @Override
    public void setParcelAddition(int newparcelnumber, int oldparcelnumber, int area) {

        Map<Integer,Integer> parcelmap = mainParcelMap.get(newparcelnumber);
        if(parcelmap == null){
            parcelmap = new Hashtable<Integer, Integer>();
            mainParcelMap.put(newparcelnumber, parcelmap);
        }
        parcelmap.put(oldparcelnumber, area);
    }

    @Override
    public void setParcelNewArea(int newparcelnumber, int newarea) {
            parcelNewAreaMap.put(newparcelnumber,newarea);
    }

    @Override
    public void setParcelOldArea(int oldparcelnumber, int oldarea) {
            parcelRemainingAreaMap.put(oldparcelnumber,oldarea);
    }

    @Override
    public void delParcelOldArea(int oldparcelnumber) {
        parcelRemainingAreaMap.remove(oldparcelnumber);
    }

    @Override
    public void setParcelRoundingDifference(int parcel, int roundingdifference) {
            parcelRoundingDifferenceMap.put(parcel,roundingdifference);
    }

    ////////////////////////////////////
    //GET-Methoden  ///////////////////
    ///////////////////////////////////

    @Override
    public List<Integer> getOrderedListOfOldParcelNumbers() {
        List<Integer> oldparcelnumbers = new ArrayList<>();
        //Add all parcelnumbers in the inner-mainDprMap from the main mainDprMap to the oldparcelmap
        for(Integer key : mainParcelMap.keySet()) {
            Map internalmap = mainParcelMap.get(key);
            for(Object keyoldparcels : internalmap.keySet()) {
                if(!oldparcelnumbers.contains(keyoldparcels)) {
                    oldparcelnumbers.add((Integer) keyoldparcels);
                }
            }
        }
        //Add also all parcelnumbers from parcelRemainingAreaMap to the oldparcelmap
        for(Integer key : parcelRemainingAreaMap.keySet()) {
            if(!oldparcelnumbers.contains(key)) {
                oldparcelnumbers.add(key);
            }
        }
        Collections.sort(oldparcelnumbers);
        return oldparcelnumbers;
    }

    @Override
    public List<Integer> getOrderedListOfNewParcelNumbers() {
        List<Integer> newparcelnumbers = new ArrayList<>(parcelNewAreaMap.keySet());
        Collections.sort(newparcelnumbers);
        return newparcelnumbers;
    }
    @Override
    public Integer getAddedArea(int newparcel, int oldparcel) {
        Map addmap = mainParcelMap.get(newparcel);
        Integer areaadded = null; //Etwas unglücklich, aber bisher unvermeidlich!
        if (addmap!=null) {
            areaadded = (Integer) addmap.get(oldparcel);
        }
        return areaadded;
    }

    @Override
    public Integer getNewArea(int newParcelNumber) {
        Integer newarea = parcelNewAreaMap.get(newParcelNumber);
        return newarea;
    }

    @Override
    public Integer getRoundingDifference(int oldParcelNumber) {
        Integer roundingdifference = parcelRoundingDifferenceMap.get(oldParcelNumber);
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
        Integer restarea = parcelRemainingAreaMap.get(oldParcelNumber);
        return restarea;
    }


}
