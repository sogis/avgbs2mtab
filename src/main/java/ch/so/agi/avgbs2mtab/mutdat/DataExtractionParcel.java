package ch.so.agi.avgbs2mtab.mutdat;

import java.util.List;

public interface DataExtractionParcel {

    public List<Integer> getOldParcelNumbers();

    public List<Integer> getNewParcelNumbers();

    public int getAddedArea(int oldParcelNumber, int newParcelNumber);

    public int getNewArea(int newParcelNumber);

    public int getRoundingDifference(int oldParcelNumber);



}
