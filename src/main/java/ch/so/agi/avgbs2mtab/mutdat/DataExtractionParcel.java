package ch.so.agi.avgbs2mtab.writeexcel;

import java.util.List;

public interface DataExtraction {

    public List<Integer> getOldProperties();

    public List<Integer> getNewProperties();

    public int getAddedArea(int oldProperty, int newProperty);

    public int getNewArea(int newProperty);

    public int getOldArea(int oldProperty);

    public int getRoundingDifference(int oldProperty);

    public List<Integer> getProperty();

    public List<Integer> getNewDPR();

    public int getAddedAreaDPR(int property, int dpr);

    public int getNewAreaDPR(int dpr);

    public int getRoundingDifferenceDPR(int dpr);

}
