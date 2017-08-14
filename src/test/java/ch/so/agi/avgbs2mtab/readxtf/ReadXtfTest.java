package ch.so.agi.avgbs2mtab.readxtf;

import ch.so.agi.avgbs2mtab.mutdat.ParcelContainer;
import org.junit.Test;

import java.util.List;

/**
 * Created by bjsvwsch on 26.07.17.
 */
public class ReadXtfTest {

    ParcelContainer parceldump = new ParcelContainer();


    @Test
    public void readFile() throws Exception {
        ReadXtf xtfreader = new ReadXtf(parceldump);
        xtfreader.readFile("/home/bjsvwsch/codebasis_test/test2.xml");
        //int addedarea = parceldump.getAddedArea(90154,748);
        int addedarea = parceldump.getAddedArea(4004,695);
        int numberofnewparcels = parceldump.getNumberOfNewParcels();
        int numberofoldparcels = parceldump.getNumberOfOldParcels();
        List<Integer> newparcels = parceldump.getNewParcelNumbers();
        List<Integer> oldparcels = parceldump.getOldParcelNumbers();
        System.out.println("Addedarea from 748 to 90154 = "+addedarea);
        System.out.println("Number of new parcels: "+numberofnewparcels);
        System.out.println("Number of old parcels: "+numberofoldparcels);
        System.out.println("New Parcels: "+newparcels);
        System.out.println("Old Parcels: "+oldparcels);
    }

    @Test
    public void getParcelAndNewArea() throws Exception {
    }

}