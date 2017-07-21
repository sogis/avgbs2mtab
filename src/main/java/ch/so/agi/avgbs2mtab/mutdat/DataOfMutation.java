package ch.so.agi.avgbs2mtab.mutdat;

/**
 * Defines which informations will be needed to create a excel-template
 */
public interface DataOfMutation {

    public int getNumberOfOldProperties();

    public int getNumberOfNewProperties();

    public int getNumberOfDPR();

    public int getNumberOfProperties();
}
