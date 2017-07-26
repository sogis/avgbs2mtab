package ch.so.agi.avgbs2mtab.readxtf;

import org.junit.Test;

import static org.junit.Assert.*;

/**
 * Created by bjsvwsch on 26.07.17.
 */
public class ReadXtfTest {
    @Test
    public void readFile() throws Exception {
        ReadXtf xtfreader = new ReadXtf();
        xtfreader.readFile("/home/bjsvwsch/codebasis_test/test.xml");
    }

    @Test
    public void getParcelAndNewArea() throws Exception {
    }

}