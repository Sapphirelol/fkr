package com.fkr.schedule.domain;

import java.util.List;

public class Registry {

    private int maxTerm;
    private int designTerm;
    private String name;
    private List<String> addresses;
    private List<Boolean> status;
    private List<String> regNum;
    private List<String> workNames;
    private List<String> workTypes;
    private List<Integer> terms;
    private List<Double> cost;

    Registry(
            String name,
            List<String> addresses,
            List<Boolean> status,
            List<String> regNum,
            List<String> workNames,
            List<String> workTypes,
            List<Integer> terms,
            List<Double> cost)
    {
        this.name = name;
        this.addresses = addresses;
        this.status = status;
        this.regNum = regNum;
        this.workNames = workNames;
        this.workTypes = workTypes;
        this.terms = terms;
        this.cost = cost;
    }

    public int getMaxTerm() {
        return maxTerm;
    }

    public void setMaxTerm(int maxTerm) {
        this.maxTerm = maxTerm;
    }

    public int getDesignTerm() {
        return designTerm;
    }

    public void setDesignTerm(int designTerm) {
        this.designTerm=designTerm;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public List<String> getAddresses() {
        return addresses;
    }

    public void setAddresses(List<String> addresses) {
        this.addresses = addresses;
    }

    public List<Boolean> getStatus() {
        return status;
    }

    public void setStatus(List<Boolean> status) {
        this.status = status;
    }

    public List<String> getWorkNames() {
        return workNames;
    }

    public void setWorkNames(List<String> workNames) {
        this.workNames = workNames;
    }

    public List<String> getWorkTypes() {
        return workTypes;
    }

    public void setWorkTypes(List<String> workTypes) {
        this.workTypes = workTypes;
    }

    public List<Integer> getTerms() {
        return terms;
    }

    public void setTerms(List<Integer> terms) {
        this.terms = terms;
    }

    public List<Double> getCost() {
        return cost;
    }

    public void setCost(List<Double> cost) {
        this.cost = cost;
    }

    public List<String> getRegNum() {
        return regNum;
    }

    public void setRegNum(List<String> regNum) {
        this.regNum = regNum;
    }
}
