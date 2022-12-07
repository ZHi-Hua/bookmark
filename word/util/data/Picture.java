package com.gentlesoft.workform.word.util.data;

import java.io.InputStream;

public class Picture {
    private InputStream pictureData;
    /*
    * 图片类型
    * */
    private int pictureType;
    private String fileName;
    private int width;
    private int height;

    public InputStream getPictureData() {
        return pictureData;
    }

    public void setPictureData(InputStream pictureData) {
        this.pictureData = pictureData;
    }

    public int getPictureType() {
        return pictureType;
    }

    public void setPictureType(int pictureType) {
        this.pictureType = pictureType;
    }

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public int getWidth() {
        return width;
    }

    public void setWidth(int width) {
        this.width = width;
    }

    public int getHeight() {
        return height;
    }

    public void setHeight(int height) {
        this.height = height;
    }
}
