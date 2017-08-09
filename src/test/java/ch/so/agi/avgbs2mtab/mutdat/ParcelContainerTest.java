package ch.so.agi.avgbs2mtab.mutdat;

import org.junit.Test;

/**
 * Created by bjsvwsch on 25.07.17.
 */
public class ParcelContainerTest {
    @Test
    public void setParcelWithAdditions() throws Exception {
        ParcelWithAddition additionnumber = new ParcelWithAddition();
        additionnumber.parcelid = 748;
        additionnumber.newarea = 463;

        Addition add = new Addition();
        add.addedfromnumber = 90154;
        add.area = 24;

        ParcelContainer getandset = new ParcelContainer();

        getandset.setParcelAddition(additionnumber.parcelid,add.addedfromnumber,add.area);

        int value = getandset.getAddedArea(748,90154);

        System.out.println(value);


    }

    @Test
    public void getAddedArea() throws Exception {

    }

    @Test
    public void setParcelRoundingDifference() throws Exception {
    }

}