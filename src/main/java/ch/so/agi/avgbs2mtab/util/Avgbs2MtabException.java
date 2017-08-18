package ch.so.agi.avgbs2mtab.util;

public class Avgbs2MtabException  extends RuntimeException {

        public static final String TYPE_NO_FILE = "TYPE_NO_FILE";
        public static final String TYPE_WRONG_EXTENSION = "TYPE_WRONG_EXTENSION";
        public static final String TYPE_NO_XML_STYLING = "TYPE_NO_XML_STYLING";
        public static final String TYPE_NOT_MATCHING_TRANSFERDATA = "TYPE_NOT_MATCHING_TRANSFERDATA";
        public static final String TYPE_NO_ACCESS_TO_FILE = "TYPE_NO_ACCESS_TO_FILE";
        public static final String TYPE_NO_ACCESS_TO_FOLDER = "TYPE_NO_ACCESS_TO_FOLDER";
        public static final String TYPE_VALIDATION_FAILED = "TYPE_VALIDATION_FAILED";
        public static final String TYPE_MISSING_PARCEL_IN_EXCEL = "TYPE_MISSING_PARCEL_IN_EXCEL";

        private String type;

        public Avgbs2MtabException(){}

        public Avgbs2MtabException(String message) {
            super(message);
        }

        public Avgbs2MtabException(String message, Throwable cause) {
            super(message, cause);
        }

        public Avgbs2MtabException(String type, String message){
            super(message);
            this.type = type;

        }

        public String getType(){
            return this.type;
        }
}
