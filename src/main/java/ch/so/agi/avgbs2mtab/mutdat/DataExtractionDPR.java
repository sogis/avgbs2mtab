package ch.so.agi.avgbs2mtab.mutdat;

import java.util.List;

public interface DataExtractionDPR {

    public List<Integer> getParcelsAffectedByDPRs();

    public List<Integer> getNewDPRs();

    public Integer getAddedAreaDPR(int parcelNumberAffectedByDPR, int dpr);

    public Integer getNewAreaDPR(int dpr);

    public Integer getRoundingDifferenceDPR(int dpr);

}
