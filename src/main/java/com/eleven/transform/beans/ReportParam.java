package com.eleven.transform.beans;

import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import lombok.ToString;

import java.util.List;

/**
 * @Author: Evan
 * @CreateTime: 2021-01-07
 * @Description:
 */
@Setter
@Getter
@ToString
@NoArgsConstructor
public class ReportParam {

    private String ptName;
    private String ptAge;
    private String ptGender;
    private String diagnoseNo;
    private String department;
    private String bedNo;
    private String hospNo;
    private String doctor;
    private String sendDate;
    private String site;
    private String diagnose;
    private String view;
    private List<ScreenShot> shots;
    private String currentDiagnose;
}
