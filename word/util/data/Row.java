package com.gentlesoft.workform.word.util.data;

import java.util.ArrayList;
import java.util.List;


import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


public class Row {
    private static final Logger log = LoggerFactory.getLogger(Row.class);

    private List<Cell> _cells = new ArrayList<>();

    public List<Cell> get_cells() {
        return _cells;
    }

    public Cell addCell(){
        Cell cell = new Cell();
        this._cells.add(cell);
        return cell;
    }

    public void addCell(Cell cell){
        this._cells.add(cell);
    }

}
