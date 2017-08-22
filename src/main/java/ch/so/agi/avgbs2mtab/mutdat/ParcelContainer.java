package ch.so.agi.avgbs2mtab.mutdat;

import java.util.*;

/**
 *
 */
public class ParcelContainer implements SetParcel,MetadataOfParcelMutation, DataExtractionParcel {

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
        Map<Integer,Integer> parcelmap = new Hashtable<>(); //Maps der Zugänge der jeweiligen neuen Parzelle innerhalb der Haupt Map.
        if (map.get(newparcelnumber) != null) {
            parcelmap = map.get(newparcelnumber);
            parcelmap.put(oldparcelnumber,area);
        } else {
            parcelmap.put(oldparcelnumber,area);
        }
        map.put(newparcelnumber,parcelmap);
    }

    @Override
    public void setParcelNewArea(int newparcelnumber, int newarea) {
        try {
            parcelnewareamap.put(newparcelnumber,newarea);
        } catch (Exception e) {
            System.out.println("Fehler in der Methode setParcelNewArea");
            throw new RuntimeException(e);
        }

    }

    @Override
    public void setParcelOldArea(int oldparcelnumber, int oldarea) {
        try {
            parcelrestareamap.put(oldparcelnumber,oldarea);
        } catch (Exception e) {
            System.out.println("Fehler in der Methode setParcelNewArea");
            throw new RuntimeException(e);
        }

    }

    @Override
    public void delParcelOldArea(int oldparcelnumber) {
        parcelrestareamap.remove(oldparcelnumber);
    }

    @Override
    public void setParcelRoundingDifference(int parcel, int roundingdifference) {
        try {
            parcelroundingdifferencemap.put(parcel,roundingdifference);
        } catch (Exception e) {
            System.out.println("Fehler in der Methode setParcelRoundingDifference");
            throw new RuntimeException(e);
        }

    }

    ////////////////////////////////////
    //GET-Methoden  ///////////////////
    ///////////////////////////////////

    @Override
    public List<Integer> getOldParcelNumbers() {
        Map<Integer,Integer> sortedmap = new TreeMap<>(parcelrestareamap); //sortiert die Map der grösse der keys nach
        List<Integer> oldparcelnumbers = new ArrayList<>(sortedmap.keySet());
        return oldparcelnumbers;
    }

    @Override
    public List<Integer> getNewParcelNumbers() {
        Map<Integer,Integer> sortedmap = new TreeMap<>(parcelnewareamap); //sortiert die Map der grösse der keys nach
        List<Integer> newparcelnumbers = new ArrayList<>(sortedmap.keySet());
        return newparcelnumbers;
    }

    @Override
    public Integer getAddedArea(int newparcel, int oldparcel) {
        Map addmap = map.get(newparcel);
        Integer areaadded = (Integer) addmap.get(oldparcel);
        return areaadded;
    }

    @Override
    public Integer getNewArea(int newParcelNumber) {
        int newarea = parcelnewareamap.get(newParcelNumber);
        return newarea;
    }

    @Override
    public Integer getRoundingDifference(int oldParcelNumber) {
        Integer roundingdifference = parcelroundingdifferencemap.get(oldParcelNumber);
        return roundingdifference;
    }

    @Override
    public Integer getNumberOfOldParcels() {
        int numberofoldparcels = parcelrestareamap.size();
        return numberofoldparcels;
    }

    @Override
    public Integer getNumberOfNewParcels() {
        int numberofnewparcels = parcelnewareamap.size();
        return numberofnewparcels;
    }

    @Override
    public Integer getRestAreaOfParcel(int oldParcelNumber) {
        int restarea = parcelrestareamap.get(oldParcelNumber);
        return restarea;
    }


}
