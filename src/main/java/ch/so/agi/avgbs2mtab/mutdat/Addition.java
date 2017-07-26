package ch.so.agi.avgbs2mtab.mutdat;


import java.util.Hashtable;
import java.util.Map;


public class Addition {
    public int addedfromnumber;
    public int area;

    Map<Integer,Integer> map=new Hashtable<Integer,Integer>();

    public Map Addition(int addedfromnumber, int area) {
        map.put(addedfromnumber, area);
        return map;
    }
}
