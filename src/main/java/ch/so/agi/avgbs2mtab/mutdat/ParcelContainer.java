package ch.so.agi.avgbs2mtab.mutdat;

import java.util.ArrayList;
import java.util.Hashtable;
import java.util.List;
import java.util.Map;

/**
 *
 */
public class ParcelContainer implements SetParcel,MetadataOfParcelMutation, DataExtractionParcel {

    Map<Integer,Map> map=new Hashtable<Integer,Map>(); //Haupt Map
    Map<Integer,Integer> parcelmap = new Hashtable<>(); //Maps der jeweiligen Parzelle innerhalb der Haupt Map
    Map<Integer,Integer> parcelnewareamap = new Hashtable<>(); //Map mit den neuen Flächen
    Map<Integer,Integer> parcelrestareamap = new Hashtable<>(); //Map mit den Rest Flächen (Diagonale
    Map<Integer,Integer> parcelroundingdifferencemap = new Hashtable<>(); //Map mit den Rundungsdifferenzen

    @Override
    public void setParcelAddition(int newparcelnumber, int oldparcelnumber, int area) {
        //Versuch die Parcelmap aus der main-map zu holen und füge den neuen Wert hinzu. Hat es noch keine Parcelmap von dieser Parzelle, dann leg eine neue an.
        //Schlussendlich füge die Parcemap (neu oder alt) wieder zur main-map hinzu.
        try {
            Map parcelmap = map.get(newparcelnumber);
            parcelmap.put(oldparcelnumber,area);
        } catch (NullPointerException e) {
            parcelmap.put(oldparcelnumber,area);
        } finally {
            map.put(newparcelnumber,parcelmap);
        }
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
    //GET-Methoden
    ///////////////////////////////////

    @Override
    public List<Integer> getOldParcelNumbers() {
        List<Integer> oldparcelnumbers = new ArrayList<>(parcelrestareamap.keySet());
        return oldparcelnumbers;
    }

    @Override
    public List<Integer> getNewParcelNumbers() {
        List<Integer> newparcelnumbers = new ArrayList<>(parcelnewareamap.keySet());
        return newparcelnumbers;
    }

    @Override
    public int getAddedArea(int newparcel, int oldparcel) {
        Map addmap = map.get(newparcel);
        Integer areaadded = (Integer) addmap.get(oldparcel);
        return areaadded;
    }



    @Override
    public int getNewArea(int newParcelNumber) {
        int newarea = parcelnewareamap.get(newParcelNumber);
        return newarea;
    }

    @Override
    public int getRoundingDifference(int oldParcelNumber) {
        int roundingdifference = parcelroundingdifferencemap.get(oldParcelNumber);
        return roundingdifference;
    }

    @Override
    public int getNumberOfOldParcels() {
        int numberofoldparcels = parcelrestareamap.size();
        System.out.println(parcelrestareamap.toString());
        return numberofoldparcels;
    }

    @Override
    public int getNumberOfNewParcels() {
        int numberofnewparcels = parcelnewareamap.size();
        return numberofnewparcels;
    }


}
