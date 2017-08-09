package ch.so.agi.avgbs2mtab.readxtf;

import ch.so.agi.avgbs2mtab.mutdat.GetParcel;
import ch.so.agi.avgbs2mtab.mutdat.ParcelContainer;
import ch.so.agi.avgbs2mtab.mutdat.SetParcel;
import org.junit.Test;

import java.util.Map;

import static org.junit.Assert.*;

/**
 * Created by bjsvwsch on 26.07.17.
 */
public class ReadXtfTest {

    ParcelContainer parceldump = new ParcelContainer();


    @Test
    public void readFile() throws Exception {
        ReadXtf xtfreader = new ReadXtf(parceldump);
        xtfreader.readFile("/home/bjsvwsch/codebasis_test/test.xml");
        int addedarea = parceldump.getAddedArea(90154,748);
        int numberofnewparcels = parceldump.getNumberOfNewParcels();
        int numberofoldparcels = parceldump.getNumberOfOldParcels();
        System.out.println("Addedarea from 748 to 90154 = "+addedarea);
        System.out.println("Number of new parcels: "+numberofnewparcels);
        System.out.println("Number of old parcels: "+numberofoldparcels);
    }

    @Test
    public void getParcelAndNewArea() throws Exception {
    }

}