package com.eleven.transform.config;

import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import lombok.ToString;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

/**
 * @Author: Evan
 * @CreateTime: 2021-01-05
 * @Description:
 */
@Setter
@Getter
@ToString
@NoArgsConstructor
@Component("pathConfig")
public class PathConfig {

    @Value("${word.sourcePath}")
    private String sourcePath;

    @Value("${word.destPath}")
    private String wordDestPath;


    @Value("${pdf.destPathdestPath}")
    private String pdfDestPath;

}
