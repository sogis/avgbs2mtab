package ch.so.agi.avgbs2mtab.mutdat;

import java.util.Hashtable;
import java.util.Map;

/**
 *
 */
public class GetAndSetParcel implements SetParcel {

    Map<Integer,Map> map=new Hashtable<Integer,Map>();

    @Override
    public void setParcelAddition(int newparcelnumber, int oldparcelnumber, int area) {
        //Eintrag mit key newparcelnumber in map finden
        //neue Map generieren falls nicht vorhanden
        //...
        //map.put(newparcelnumber,add.map);
        System.out.println("Values in setParcelAddition: "+newparcelnumber+" "+oldparcelnumber+" "+area);
    }

    @Override
    public void setParcelNewArea(int newparcelnumber, int newarea) {

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
