package ch.so.agi.avgbs2mtab.mutdat;


public interface SetParcel {


    public void setParcelAddition(int newparcelnumber, int oldparcelnumber, int area);

    public void setParcelNewArea(int newparcelnumber, int newarea);

    public void setParcelOldArea(int oldparcelnumber, int oldarea);

    void delParcelOldArea(int oldparcelnumber);

    public void setParcelRoundingDifference(int parcel, int roundingdifference);

}
