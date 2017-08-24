package ch.so.agi.avgbs2mtab.main;

import ch.so.agi.avgbs2mtab.mutdat.*;
import ch.so.agi.avgbs2mtab.readxtf.ReadXtf;
import ch.so.agi.avgbs2mtab.writeexcel.XlsxWriter;

import java.io.IOException;

/**
 * Main class
 */
public class Main {
    public void runConversion(String inputfile, String outputfilename) throws IOException {

        ParcelContainer parceldump = new ParcelContainer();
        DPRContainer dprdump = new DPRContainer();
        XlsxWriter xlsxWriter = new XlsxWriter(parceldump, dprdump, parceldump, dprdump);

        ReadXtf xtfreader = new ReadXtf((SetParcel)parceldump, (SetDPR)dprdump, (DataExtractionParcel)parceldump);
        xtfreader.readFile(inputfile);
        xlsxWriter.writeXlsx(outputfilename);


    }

}
