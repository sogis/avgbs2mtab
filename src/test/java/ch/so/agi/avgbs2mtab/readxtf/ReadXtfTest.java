package ch.so.agi.avgbs2mtab.readxtf;

import ch.so.agi.avgbs2mtab.mutdat.DPRContainer;
import ch.so.agi.avgbs2mtab.mutdat.ParcelContainer;
import org.junit.Ignore;
import org.junit.Test;

import java.io.IOException;
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
        ClassLoader classLoader = getClass().getClassLoader();
        xtfreader.readFile(classLoader.getResource("SO0200002407_4003_20150807.xtf").getPath());
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
        ClassLoader classLoader = getClass().getClassLoader();
        xtfreader.readFile(classLoader.getResource("SO0200002407_4004_20150810.xtf").getPath());
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
        ClassLoader classLoader = getClass().getClassLoader();
        xtfreader.readFile(classLoader.getResource("SO0200002407_40051_20150811.xtf").getPath());
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
        ClassLoader classLoader = getClass().getClassLoader();
        xtfreader.readFile(classLoader.getResource("SO0200002407_40061_20150814.xtf").getPath());
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
        ClassLoader classLoader = getClass().getClassLoader();
        xtfreader.readFile(classLoader.getResource("SO0200002427_809_20170529.xtf").getPath());
        int numberofnewparcels = parceldump.getNumberOfNewParcels();
        int numberofoldparcels = parceldump.getNumberOfOldParcels();
        List<Integer> newparcels = parceldump.getNewParcelNumbers();
        List<Integer> oldparcels = parceldump.getOldParcelNumbers();
        int gotareafromoldparcel = parceldump.getAddedArea(1273,1864);
        int restarea = parceldump.getRestAreaOfParcel(1273);
        assertTrue(numberofnewparcels==3);
        assertTrue(numberofoldparcels==2);
        List<Integer> newparcelsasitshouldbe = Arrays.asList(1273,1864,1973);
        List<Integer> oldparcelsasitshouldbe = Arrays.asList(1273,1864);
        assertTrue(newparcels.containsAll(newparcelsasitshouldbe) && newparcels.size()==newparcelsasitshouldbe.size());
        assertTrue(oldparcels.containsAll(oldparcelsasitshouldbe) && oldparcels.size()==oldparcelsasitshouldbe.size());
        assertTrue(gotareafromoldparcel==69);
        assertTrue(restarea==352);
        int numberofdprs = dprdump.getNumberOfDPRs();
        int numberofareasafected = dprdump.getNumberOfParcelsAffectedByDPRs();
        List<Integer> parcelsaffectedbydprsa = dprdump.getParcelsAffectedByDPRs();
        List<Integer> newdprs = dprdump.getNewDPRs();
        assertTrue(numberofdprs==1);
        assertTrue(numberofareasafected==1);
        List<Integer> newdprsastheyshouldbe = Arrays.asList(1941);
        List<Integer> parcelsaffectedasitshouldbe = Arrays.asList(1864);
        assertTrue(newdprs.containsAll(newdprsastheyshouldbe) && newdprs.size()==newdprsastheyshouldbe.size());
        assertTrue(parcelsaffectedbydprsa.containsAll(parcelsaffectedasitshouldbe) && parcelsaffectedbydprsa.size()==parcelsaffectedasitshouldbe.size());
    }

    @Test
    public void oensingentest() throws Exception {
        ReadXtf xtfreader = new ReadXtf(parceldump, dprdump);
        ClassLoader classLoader = getClass().getClassLoader();
        xtfreader.readFile(classLoader.getResource("SO0200002407_4002_20150807.xtf").getPath());
        List<Integer> newparcels = parceldump.getNewParcelNumbers();
        List<Integer> oldparcels = parceldump.getOldParcelNumbers();
        int numberofoldparcels = parceldump.getNumberOfOldParcels();
        List<Integer> newparcelsasitshouldbe = Arrays.asList(2199);
        List<Integer> oldparcelsasitshouldbe = Arrays.asList(681,682,2199);
        assertTrue(newparcels.containsAll(newparcelsasitshouldbe) && newparcels.size()==newparcelsasitshouldbe.size());
        assertTrue(oldparcels.containsAll(oldparcelsasitshouldbe) && oldparcels.size()==oldparcelsasitshouldbe.size());
        assertTrue(numberofoldparcels==3);
    }

    @Test
    public void readFileWithoutDRP() throws Exception {
        ReadXtf xtfreader = new ReadXtf(parceldump, dprdump);
        ClassLoader classLoader = getClass().getClassLoader();
        xtfreader.readFile(classLoader.getResource("SO0200002407_4002_20150807.xtf").getPath());
        int numberofdprs = dprdump.getNumberOfDPRs();
        List<Integer> parcelsaffectedbydprsa = dprdump.getParcelsAffectedByDPRs();
        List<Integer> newdprs = dprdump.getNewDPRs();
        assertTrue(numberofdprs==0);
    }

    @Test
    public void filewithwrongextension() {
        ReadXtf xtfreader = new ReadXtf(parceldump, dprdump);
        ClassLoader classLoader = getClass().getClassLoader();
        try {
            xtfreader.readFile(classLoader.getResource("SO0200002407_4001_20150806_wrong_type.xml").getPath());
        }catch (IOException e) {
            System.out.println("Got IOException as Expected. "+e);

        }
    }
}