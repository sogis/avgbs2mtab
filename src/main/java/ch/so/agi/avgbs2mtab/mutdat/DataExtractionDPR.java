package ch.so.agi.avgbs2mtab.mutdat;

import java.util.List;

public interface DataExtractionDPR {

    public List<Integer> getParcelsAffectedByDPRs();

    public List<Integer> getNewDPRs();

    public int getAddedAreaDPR(int parcelNumberAffectedByDPR, int dpr);

    public int getNewAreaDPR(int dpr);

    public int getRoundingDifferenceDPR(int dpr);

}
