package com.grapecity.documents.excel.examples;

import java.util.Comparator;
import java.util.HashMap;

public class ExampleComparer implements Comparator<ExampleBase> {

    private HashMap<String, String> _sortOrders = new HashMap<String, String>();

    public ExampleComparer() {
        // root children orders
        _sortOrders.put("tutorial", "a");
        _sortOrders.put("features", "b");
        _sortOrders.put("spreadsheetsviewer", "c");
        _sortOrders.put("excelreporting", "d");
        _sortOrders.put("exceltemplates", "e");

        // Features children orders
        _sortOrders.put("rangeoperations", "a");
        _sortOrders.put("formatting", "b");
        _sortOrders.put("tables", "c");
        _sortOrders.put("conditionalformatting", "d");
        _sortOrders.put("datavalidation", "e");
        _sortOrders.put("formulas", "f");
        _sortOrders.put("grouping", "g");
        _sortOrders.put("filtering", "h");
        _sortOrders.put("sorting", "i");
        _sortOrders.put("sparklines", "j");
        _sortOrders.put("charts", "k");
        _sortOrders.put("shape", "l");
        _sortOrders.put("picture", "m");
        _sortOrders.put("slicer", "n");
        _sortOrders.put("comments", "o");
        _sortOrders.put("pivotTable", "p");
        _sortOrders.put("hyperlinks", "q");
        _sortOrders.put("theme", "r");
        _sortOrders.put("workbook", "s");
        _sortOrders.put("worksheets", "t");
    }

    @Override
    public int compare(ExampleBase o1, ExampleBase o2) {
        if (o1 instanceof Tutorial) {
            return -1;
        } else if (o2 instanceof Tutorial) {
            return 1;
        }

        String xSortKey;
        if (_sortOrders.containsKey(o1.getShortID())) {
            xSortKey = _sortOrders.get(o1.getShortID());
        } else {
            xSortKey = o1.getID();

        }

        String ySortKey;
        if (_sortOrders.containsKey(o2.getShortID())) {
            ySortKey = _sortOrders.get(o2.getShortID());
        } else {
            ySortKey = o2.getID();

        }

        if (o1 instanceof FolderExample) {
            if (o1 instanceof FolderExample) {
                return xSortKey.compareTo(ySortKey);
            } else {
                return -1;
            }
        } else {
            if (o1 instanceof FolderExample) {
                return 1;
            } else {
                return xSortKey.compareTo(ySortKey);
            }
        }
    }

}
