package ch.so.agi.avgbs2mtab.mutdat;

import java.util.HashMap;

/**
 *
 */
public interface SetDPR {

    public void setDPRWithAdditions(Integer dprnumber, String laysonref, Integer area);

    void setDPRNumberAndRef(String ref, int parcelnumber);

    void setDPRNewArea(Integer parcelnumber, Integer area);

}
