package ch.so.agi.avgbs2mtab.mutdat;

import org.junit.Test;

/**
 * Created by bjsvwsch on 25.07.17.
 */
public class GetAndSetParcelTest {
    @Test
    public void setParcelWithAdditions() throws Exception {
        ParcelWithAddition additionnumber = new ParcelWithAddition();
        additionnumber.parcelid = 748;
        additionnumber.newarea = 463;

        Addition add = new Addition();
        add.addedfromnumber = 90154;
        add.area = 24;

        GetAndSetParcel getandset = new GetAndSetParcel();

        getandset.setParcelWithAdditions(additionnumber,add);

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