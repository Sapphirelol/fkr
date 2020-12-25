package com.fkr.schedule.domain;

import java.util.List;

public class WorkNames {

    private List<Integer> ids;
    private List<String> shortNames;
    private List<String> longNames;

    public WorkNames(List<Integer> ids, List<String> shortNames, List<String> longName) {
        this.ids = ids;
        this.shortNames = shortNames;
        this.longNames = longName;
    }

    public List<Integer> getIds() {
        return ids;
    }

    public void setIds(List<Integer> ids) {
        this.ids = ids;
    }

    public List<String> getShortNames() {
        return shortNames;
    }

    public void setShortNames(List<String> shortNames) {
        this.shortNames = shortNames;
    }

    public List<String> getLongNames() {
        return longNames;
    }

    public void setLongName(List<String> longName) {
        this.longNames = longName;
    }
}
