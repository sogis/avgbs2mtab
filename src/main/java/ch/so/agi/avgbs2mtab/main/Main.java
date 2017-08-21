package ch.so.agi.avgbs2mtab.main;

import ch.so.agi.avgbs2mtab.mutdat.DPRContainer;
import ch.so.agi.avgbs2mtab.mutdat.ParcelContainer;
import ch.so.agi.avgbs2mtab.readxtf.ReadXtf;

/**
 * Main class
 */
public class Main {
    public void runConversion(String inputfile, String outputfilename) {

        ParcelContainer parceldump = new ParcelContainer();
        DPRContainer dprdump = new DPRContainer();

        ReadXtf xtfreader = new ReadXtf(parceldump, dprdump);
        xtfreader.readFile(inputfile);
    }

}
