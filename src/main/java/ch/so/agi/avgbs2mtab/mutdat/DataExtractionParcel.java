package ch.so.agi.avgbs2mtab.mutdat;

import java.util.List;

public interface DataExtractionParcel {

    public List<Integer> getOldParcels();

    public List<Integer> getNewParcels();

    public int getAddedArea(int oldParcelNumber, int newParcelNumber);

    public int getNewArea(int newParcelNumber);

    public int getRoundingDifference(int oldParcelNumber);



}
