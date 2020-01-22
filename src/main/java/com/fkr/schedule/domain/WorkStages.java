package com.fkr.schedule.domain;

import java.util.List;

public class WorkStages {

    private String workName;
    private List<String> names;
    private List<Short> weeksTo;
    private List<Short> weeksFor;

    WorkStages(String workName,
               List<String> names,
               List<Short> weeksTo,
               List<Short> weeksFor
    ) {
        this.workName = workName;
        this.names = names;
        this.weeksTo = weeksTo;
        this.weeksFor = weeksFor;
    }

    public String getWorkName() {
        return workName;
    }

    public void setWorkName(String workName) {
        this.workName = workName;
    }

    public List<String> getNames() {
        return names;
    }

    public void setNames(List<String> names) {
        this.names = names;
    }

    public List<Short> getWeeksTo() {
        return weeksTo;
    }

    public void setWeeksTo(List<Short> weeksTo) {
        this.weeksTo = weeksTo;
    }

    public List<Short> getWeeksFor() {
        return weeksFor;
    }

    public void setWeeksFor(List<Short> weeksFor) {
        this.weeksFor = weeksFor;
    }
}
