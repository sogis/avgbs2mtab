package ch.so.agi.avgbs2mtab.readxtf;

import ch.so.agi.avgbs2mtab.mutdat.DPRContainer;
import ch.so.agi.avgbs2mtab.mutdat.ParcelContainer;
import org.junit.Test;

import java.util.Arrays;
import java.util.List;

import static junit.framework.TestCase.assertTrue;

/**
 * Unit-Tests for avgbs2mtab
 */
public class ReadXtfTest {

    ParcelContainer parceldump = new ParcelContainer();
    DPRContainer dprdump = new DPRContainer();


    @Test
    public void readFile1() throws Exception {
        ReadXtf xtfreader = new ReadXtf(parceldump, dprdump);
        xtfreader.readFile("/home/bjsvwsch/codebasis_test/test.xml");
        int addedarea = parceldump.getAddedArea(90154,748);
        int numberofnewparcels = parceldump.getNumberOfNewParcels();
        int numberofoldparcels = parceldump.getNumberOfOldParcels();
        List<Integer> newparcels = parceldump.getNewParcelNumbers();
        List<Integer> oldparcels = parceldump.getOldParcelNumbers();
        int restarea = parceldump.getRestAreaOfParcel(751);
        int newarea = parceldump.getNewArea(751);
        int rundungsdifferenz = parceldump.getRoundingDifference(753);
        assertTrue(addedarea ==24);
        assertTrue(numberofnewparcels == 14);
        assertTrue(numberofoldparcels == 14);
        List<Integer> newandoldparcelsastheyshouldbe = Arrays.asList(748,749,750,751,753,755,756,757,758,1303,1799,2097,2098,90154);
        assertTrue(newparcels.containsAll(newandoldparcelsastheyshouldbe) && newparcels.size()==newandoldparcelsastheyshouldbe.size());
        assertTrue(oldparcels.containsAll(newandoldparcelsastheyshouldbe) && oldparcels.size()==newandoldparcelsastheyshouldbe.size());
        assertTrue(restarea == 1157);
        assertTrue(newarea == 1176);
        assertTrue(rundungsdifferenz == -1);
    }

    @Test
    public void readFile2() throws Exception {
        ReadXtf xtfreader = new ReadXtf(parceldump, dprdump);
        xtfreader.readFile("/home/bjsvwsch/codebasis_test/test2.xml");
        int addedarea = parceldump.getAddedArea(4004,695);
        int numberofnewparcels = parceldump.getNumberOfNewParcels();
        int numberofoldparcels = parceldump.getNumberOfOldParcels();
        List<Integer> newparcels = parceldump.getNewParcelNumbers();
        List<Integer> oldparcels = parceldump.getOldParcelNumbers();
        int restarea = parceldump.getRestAreaOfParcel(701);
        int newarea = parceldump.getNewArea(701);
        int rundungsdifferenz = parceldump.getRoundingDifference(701);
        assertTrue(addedarea == 242);
        assertTrue(numberofnewparcels == 7);
        assertTrue(numberofoldparcels == 6);
        List<Integer> newparcelsastheyshouldbe = Arrays.asList(695,696,697,701,870,874,4004);
        List<Integer> oldparcelsastheyshouldbe = Arrays.asList(695,696,697,701,870,874);
        assertTrue(newparcels.containsAll(newparcelsastheyshouldbe) && newparcels.size()==newparcelsastheyshouldbe.size());
        assertTrue(oldparcels.containsAll(oldparcelsastheyshouldbe) && oldparcels.size()==oldparcelsastheyshouldbe.size());
        assertTrue(restarea == 1112);
        assertTrue(newarea == 1114);
        assertTrue(rundungsdifferenz == -1);
    }

    @Test
    public void readFileWithDPR() throws Exception {
        ReadXtf xtfreader = new ReadXtf(parceldump, dprdump);
        xtfreader.readFile("/home/bjsvwsch/codebasis_test/test_mit_dpr.xml");
        int numberofdprs = dprdump.getNumberOfDPRs();
        int numberofareasafected = dprdump.getNumberOfParcelsAffectedByDPRs();
        List<Integer> parcelsaffectedbydprs = dprdump.getParcelsAffectedByDPRs();
        List<Integer> newdprs = dprdump.getNewDPRs();
        int getaddedarea = dprdump.getAddedAreaDPR(2141,40051);
        int newarea = dprdump.getNewAreaDPR(40051);
        int roundingdifference = dprdump.getRoundingDifferenceDPR(40051);
        assertTrue(numberofdprs==1);
        assertTrue(numberofareasafected==2);
        List<Integer> parcelsaffectedbydprsasitshouldbe = Arrays.asList(2141,2142);
        List<Integer> newdprsasitshouldbe = Arrays.asList(40051);
        assertTrue(parcelsaffectedbydprs.containsAll(parcelsaffectedbydprsasitshouldbe) && parcelsaffectedbydprs.size()==parcelsaffectedbydprsasitshouldbe.size());
        assertTrue(newdprs.containsAll(newdprsasitshouldbe) && newdprs.size()==newdprsasitshouldbe.size());
        assertTrue(getaddedarea == 1175);
        assertTrue(newarea == 3656);
        assertTrue(roundingdifference == 0);
    }

    @Test
    public void readFileWithDeleteDRP() throws Exception {
        ReadXtf xtfreader = new ReadXtf(parceldump, dprdump);
        xtfreader.readFile("/home/bjsvwsch/codebasis_test/test_loeschen_dpr.xml");
        int numberofdprs = dprdump.getNumberOfDPRs();
        List<Integer> parcelsaffectedbydprsa = dprdump.getParcelsAffectedByDPRs();
        List<Integer> newdprs = dprdump.getNewDPRs();
        int newarea = dprdump.getNewAreaDPR(40051);
        assertTrue(numberofdprs==1);
        assertTrue(parcelsaffectedbydprsa.size()==0);
        List<Integer> newdprsasitshouldbe = Arrays.asList(40051);
        assertTrue(newdprs.containsAll(newdprsasitshouldbe) && newdprs.size()==newdprsasitshouldbe.size());
        assertTrue(newarea==0);
    }

    @Test
    public void readComplexFile() throws Exception {
        ReadXtf xtfreader = new ReadXtf(parceldump, dprdump);
        xtfreader.readFile("/home/bjsvwsch/codebasis_test/komplex_test.xml");
        int numberofnewparcels = parceldump.getNumberOfNewParcels();
        int numberofoldparcels = parceldump.getNumberOfOldParcels();
        List<Integer> newparcels = parceldump.getNewParcelNumbers();
        List<Integer> oldparcels = parceldump.getOldParcelNumbers();
        int gotareafromoldparcel = parceldump.getAddedArea(1273,1864);
        int restarea = parceldump.getRestAreaOfParcel(1273);
        System.out.println("Number of New Parcels = "+numberofnewparcels);
        System.out.println("Number of Old Parcels = "+numberofoldparcels);
        System.out.println("New Parcels: "+newparcels);
        System.out.println("Old Parcels: "+oldparcels);
        System.out.println("Area that Parcel 1273 got from Parcel 1864 should be 69: "+gotareafromoldparcel);
        System.out.println("RestArea From Parcel 1273 Should be 352: "+restarea);
        int numberofdprs = dprdump.getNumberOfDPRs();
        int numberofareasafected = dprdump.getNumberOfParcelsAffectedByDPRs();
        List<Integer> parcelsaffectedbydprsa = dprdump.getParcelsAffectedByDPRs();
        List<Integer> newdprs = dprdump.getNewDPRs();
        System.out.println("Number of DPRs = "+numberofdprs);
        System.out.println("Number of Parcels affected by DPRs = "+numberofareasafected);
        System.out.println("New DPRs: "+newdprs);
        System.out.println("Affected Parcels: "+parcelsaffectedbydprsa);
    }



}