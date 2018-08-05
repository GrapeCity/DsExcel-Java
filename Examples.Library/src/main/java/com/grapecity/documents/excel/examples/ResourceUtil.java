package com.grapecity.documents.excel.examples;

import org.json.JSONObject;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Set;

public class ResourceUtil {
    private static JSONObject _codeJson = null;
    private static HashMap<String, JSONObject> _stringResources = new HashMap<String, JSONObject>();

    static {
        if (_codeJson == null) {
            InputStream jsonStream = getResourceStream("codeResource.json");
//            InputStream stream1 = getResourceStream("xlsx/AgingReport.xlsx");
//            String value = stream1.toString();
//            InputStream stream2 = getResourceStream("xlsx/Event budget.xlsx");
//            String value2 = stream2.toString();
            if (jsonStream != null) {
                _codeJson = new JSONObject(inputStreamToString(jsonStream));
            }
        }
    }

    public static String[] listAllExamples() {
        if (_codeJson != null) {
            return _codeJson.keySet().toArray(new String[]{});
        }
        return null;
    }

    public static String getCodeResource(String key) {
        if (_codeJson != null && _codeJson.has(key)) {
            return _codeJson.getString(key);
        }
        return null;
    }

    public static String getResourceByCulture(String culture, String key) {
        try {
            if (_stringResources.containsKey(culture)) {
                JSONObject jsonObject = _stringResources.get(culture);
                if (jsonObject.has(key)) {
                    return jsonObject.getString(key);
                }
            } else {
                InputStream inputStream = getResourceStream(String.format("stringResources_%s.json", culture));
                if (inputStream != null) {
                    JSONObject currentResource = new JSONObject(inputStreamToString(inputStream));
                    _stringResources.put(culture, currentResource);
                    if (currentResource.has(key)) {
                        return currentResource.getString(key);
                    }
                } else {
                    return getResourceForEnUS(key);
                }
            }
            return null;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private static String getResourceForEnUS(String key) {
        JSONObject enUSResource = _stringResources.get("en-US");
        if (enUSResource == null) {
            InputStream inputStream = getResourceStream("stringResources_en-US.json");
            enUSResource = new JSONObject(inputStreamToString(inputStream));

            _stringResources.put("en-US", enUSResource);
        }
        if (enUSResource.has(key)) {
            return enUSResource.getString(key);
        }
        return null;
    }

    public static InputStream getResourceStream(String resource) {
        return ExampleBase.class.getClassLoader().getResourceAsStream(resource);
    }


    public static String inputStreamToString(InputStream inputStream) {
        try {
            ByteArrayOutputStream result = inputStreamToByteArrayOutputStream(inputStream);
            // StandardCharsets.UTF_8.name() > JDK 7
            return result.toString("UTF-8");
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    public static ByteArrayOutputStream inputStreamToByteArrayOutputStream(InputStream inputStream) {
        ByteArrayOutputStream result = new ByteArrayOutputStream();
        byte[] buffer = new byte[1024];
        int length;
        try {
            while ((length = inputStream.read(buffer)) != -1) {
                result.write(buffer, 0, length);
            }
            return result;
        } catch (Exception e) {
            return null;
        }
    }

}
