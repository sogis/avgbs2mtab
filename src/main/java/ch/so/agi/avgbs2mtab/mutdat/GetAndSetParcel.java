package ch.so.agi.avgbs2mtab.mutdat;

import java.util.Hashtable;
import java.util.Map;

/**
 *
 */
public class GetAndSetParcel implements SetParcel {

    Map<Integer,Map> map=new Hashtable<Integer,Map>();
    Map<Integer,Integer> parcelmap = new Hashtable<>();
    Map<Integer,Integer> parcelnewareamap = new Hashtable<>();

    @Override
    public void setParcelAddition(int newparcelnumber, int oldparcelnumber, int area) {
        try {
            Map parcelmap = map.get(newparcelnumber);
            parcelmap.put(oldparcelnumber,area);
        } catch (NullPointerException e) {
            parcelmap.put(oldparcelnumber,area);
        } finally {
            map.put(newparcelnumber,parcelmap);
        }

        System.out.println("Values in setParcelAddition: "+newparcelnumber+" "+oldparcelnumber+" "+area);
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

    public int getAddedArea(int newparcel, int oldparcel) {
        Map addmap = map.get(newparcel);

        Integer areaadded = (Integer) addmap.get(oldparcel);

        return areaadded;
    }

    @Override
    public void setParcelRoundingDifference(int parcel, int roundingdifference) {

    }
}
