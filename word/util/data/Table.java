package com.gentlesoft.workform.word.util.data;

import java.util.ArrayList;
import java.util.List;

public class Table {

    private List<Row> _row = new ArrayList<>();;

    public List<Row> get_row() {
        return _row;
    }

    public Row addRow(){
        Row row = new Row();
        this._row.add(row);
        return row;
    }
    
    public void addRow(Row row){
        this._row.add(row);
    }
}
