package ch.so.agi.avgbs2mtab.mutdat;

import java.util.List;

public interface DataExtractionParcel {

    public List<Integer> getOrderedListOfOldParcelNumbers();

    public List<Integer> getOrderedListOfNewParcelNumbers();

    public Integer getAddedArea(int oldParcelNumber, int newParcelNumber);

    public Integer getNewArea(int newParcelNumber);

    public Integer getRoundingDifference(int oldParcelNumber);

    //todo Rename: getRemainingAreaOfParcel(int parcelNumber);
    public Integer getRestAreaOfParcel(int oldParcelNumber);

}
