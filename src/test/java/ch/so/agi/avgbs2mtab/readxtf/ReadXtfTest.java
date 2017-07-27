package ch.so.agi.avgbs2mtab.readxtf;

import org.junit.Test;

import java.util.Map;

import static org.junit.Assert.*;

/**
 * Created by bjsvwsch on 26.07.17.
 */
public class ReadXtfTest {
    @Test
    public void readFile() throws Exception {
        ReadXtf xtfreader = new ReadXtf();
        Map map = xtfreader.readFile("/home/bjsvwsch/codebasis_test/test.xml");
        System.out.println(map.toString());
    }

    @Test
    public void getParcelAndNewArea() throws Exception {
    }

}